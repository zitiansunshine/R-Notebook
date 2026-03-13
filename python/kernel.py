#!/usr/bin/env python3
"""
kernel.py — Python execution kernel for RNotebook notebook controller.

Protocol: newline-delimited JSON over stdin / stdout.

Inbound messages
  {"type":"ping"}
  {"type":"exec","chunk_id":"nb-0","code":"...","fig_width":7,"fig_height":5,"dpi":120}
  {"type":"reset"}
  {"type":"vars"}

Outbound messages
  {"type":"pong"}
  {"type":"result","chunk_id":"nb-0","stdout":"...","stderr":"...",
   "plots":[<base64png>,...], "dataframes":[<DataFrameResult>,...], "error":null}
  {"type":"vars_result","vars":[{"name":"x","type":"int","size":"","value":"3"},…]}
"""

import sys
import os
import json
import io
import base64
import traceback
import ast
from contextlib import redirect_stdout, redirect_stderr

# Force Agg backend before any matplotlib import so no GUI window appears.
os.environ.setdefault("MPLBACKEND", "Agg")

# Monkey-patch plt.show() to a no-op so user code calling plt.show() neither
# raises an error nor emits "cannot show the figure" UserWarning under Agg.
try:
    import matplotlib.pyplot as _mpl_plt
    _mpl_plt.show = lambda *_a, **_k: None
except ImportError:
    pass

# Persistent execution namespace shared across all cells.
_ns: dict = {}


# ---------------------------------------------------------------------------
# Helpers

def _send(msg: dict) -> None:
    print(json.dumps(msg, ensure_ascii=False), file=sys.__stdout__, flush=True)


class _StreamingBuffer(io.TextIOBase):
    """Capture writes while forwarding them to the host as stream events."""

    def __init__(self, chunk_id: str, stream_name: str, console_parts: list[str] | None = None):
        super().__init__()
        self._chunk_id = chunk_id
        self._stream_name = stream_name
        self._buffer = io.StringIO()
        self._console_parts = console_parts

    def write(self, text: str) -> int:
        if not text:
            return 0
        self._buffer.write(text)
        if self._console_parts is not None:
            self._console_parts.append(text)
        _send({
            "type": "stream",
            "chunk_id": self._chunk_id,
            "stream": self._stream_name,
            "text": text,
        })
        return len(text)

    def flush(self) -> None:
        return None

    def getvalue(self) -> str:
        return self._buffer.getvalue()


def _safe_repr(v) -> str:
    try:
        r = repr(v)
        return r[:80] + ("…" if len(r) > 80 else "")
    except Exception:
        return "<repr failed>"


def _serialise_df(name: str, df, page: int = 0, page_size: int = 50) -> dict:
    """Convert a pandas DataFrame to the RNotebook DataFrameResult format."""
    try:
        import pandas as pd
        import numpy as np
    except ImportError:
        return {}

    nrow_total = len(df)
    ncol_total = len(df.columns)
    start = page * page_size
    end   = min(start + page_size, nrow_total)
    slc   = df.iloc[start:end]

    columns = [{"name": str(c), "type": str(df[c].dtype)} for c in df.columns]

    data = []
    for _, row in slc.iterrows():
        row_dict: dict = {}
        for col in df.columns:
            val = row[col]
            try:
                if pd.isna(val):
                    row_dict[str(col)] = None
                elif isinstance(val, (np.integer,)):
                    row_dict[str(col)] = int(val)
                elif isinstance(val, (np.floating,)):
                    row_dict[str(col)] = float(round(float(val), 6))
                else:
                    row_dict[str(col)] = str(val)
            except (TypeError, ValueError):
                row_dict[str(col)] = str(val)
        data.append(row_dict)

    return {
        "name":    name,
        "nrow":    nrow_total,
        "ncol":    ncol_total,
        "pages":   max(1, (nrow_total + page_size - 1) // page_size),
        "page":    page,
        "columns": columns,
        "data":    data,
    }


def _exec_and_last(code: str, ns: dict):
    """Execute code; return the value of the final expression (or None)."""
    try:
        tree = ast.parse(code, mode="exec")
    except SyntaxError:
        exec(code, ns)      # let exec raise the SyntaxError with full context
        return None

    if not tree.body:
        return None

    last = tree.body[-1]
    if not isinstance(last, ast.Expr):
        exec(compile(tree, "<cell>", "exec"), ns)
        return None

    # Split: run everything before the last expression, then eval it.
    rest = ast.Module(body=tree.body[:-1], type_ignores=[])
    if rest.body:
        exec(compile(rest, "<cell>", "exec"), ns)

    expr = ast.Expression(body=last.value)
    return eval(compile(expr, "<cell>", "eval"), ns)


# ---------------------------------------------------------------------------
# Cell execution

def _exec_cell(chunk_id: str, code: str,
               fig_width: float = 7, fig_height: float = 5,
               dpi: int = 120) -> dict:

    global _ns

    console_parts: list[str] = []
    stdout_buf = _StreamingBuffer(chunk_id, "stdout", console_parts)
    stderr_buf = _StreamingBuffer(chunk_id, "stderr", console_parts)
    plots:      list[str] = []
    plots_html: list[str] = []
    dataframes: list[dict] = []
    error: str | None = None

    # ---- matplotlib setup --------------------------------------------------
    plt = None
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        plt.rcParams["figure.figsize"] = [fig_width, fig_height]
        plt.rcParams["figure.dpi"]     = dpi
        plt.close("all")
    except ImportError:
        pass

    last_val = None

    # ---- execute -----------------------------------------------------------
    try:
        with redirect_stdout(stdout_buf), redirect_stderr(stderr_buf):
            last_val = _exec_and_last(code, _ns)
    except KeyboardInterrupt:
        error = "Interrupted by user"
        last_val = None
    except Exception:
        error = traceback.format_exc()

    # ---- detect DataFrames explicitly named as standalone expressions -------
    # Walk the top-level AST statements.  Any bare Name expression whose value
    # is a DataFrame is added in source order (e.g. `df` on its own line).
    # Loop/conditional intermediaries (like `subset = ...`) are ignored.
    try:
        import pandas as pd
        try:
            _tree = ast.parse(code, mode="exec")
            for _stmt in _tree.body:
                if isinstance(_stmt, ast.Expr) and isinstance(_stmt.value, ast.Name):
                    _name = _stmt.value.id
                    if not _name.startswith("_") and _name in _ns:
                        _v = _ns[_name]
                        if isinstance(_v, pd.DataFrame):
                            if not any(d.get("name") == _name for d in dataframes):
                                dataframes.append(_serialise_df(_name, _v, page_size=2000))
        except SyntaxError:
            pass
        # Fallback: last expression was a DataFrame not caught above
        if isinstance(last_val, pd.DataFrame):
            _name = _find_df_name(last_val, _ns) or "result"
            if not any(d.get("name") == _name for d in dataframes):
                dataframes.append(_serialise_df(_name, last_val, page_size=2000))
    except ImportError:
        pass

    # ---- capture matplotlib figures ----------------------------------------
    if plt is not None:
        try:
            for fn in plt.get_fignums():
                fig = plt.figure(fn)
                buf = io.BytesIO()
                fig.savefig(buf, format="png", bbox_inches="tight",
                            dpi=dpi, facecolor="white")
                buf.seek(0)
                plots.append(base64.b64encode(buf.read()).decode())
            plt.close("all")
        except Exception as exc:
            stderr_buf.write(f"\n[plot capture error: {exc}]\n")

    # ---- detect Plotly figures ---------------------------------------------
    # If the last expression is a Plotly figure, export interactive HTML.
    try:
        import plotly.graph_objs as _pgo
        if isinstance(last_val, _pgo.BaseFigure):
            html = last_val.to_html(
                include_plotlyjs="cdn",
                full_html=False,
                config={"responsive": True},
            )
            plots_html.append(html)
    except (ImportError, AttributeError):
        pass

    return {
        "type":       "result",
        "chunk_id":   chunk_id,
        "console":    "".join(console_parts),
        "stdout":     stdout_buf.getvalue(),
        "stderr":     stderr_buf.getvalue(),
        "plots":      plots,
        "plots_html": plots_html,
        "dataframes": dataframes,
        "error":      error,
    }


def _find_df_name(df, ns: dict) -> str | None:
    """Try to find a variable name for this exact DataFrame object."""
    for k, v in ns.items():
        if v is df and not k.startswith("_"):
            return k
    return None


# ---------------------------------------------------------------------------
# vars inspection

def _vars_result() -> dict:
    try:
        import pandas as pd
        has_pd = True
    except ImportError:
        has_pd = False

    rows = []
    for k, v in _ns.items():
        if k.startswith("_"):
            continue
        tname = type(v).__name__
        if has_pd and isinstance(v, pd.DataFrame):
            size = f"{len(v)} × {len(v.columns)}"
            val  = "DataFrame"
        elif isinstance(v, (list, tuple)):
            size = str(len(v))
            val  = _safe_repr(v)
        elif isinstance(v, dict):
            size = str(len(v))
            val  = _safe_repr(v)
        else:
            size = ""
            val  = _safe_repr(v)
        rows.append({"name": k, "type": tname, "size": size, "value": val})

    return {"type": "vars_result", "vars": rows}


# ---------------------------------------------------------------------------
# Main loop

def main() -> None:
    for raw in sys.stdin:
        raw = raw.strip()
        if not raw:
            continue
        try:
            msg = json.loads(raw)
        except json.JSONDecodeError:
            continue

        t = msg.get("type")

        try:
            if t == "ping":
                _send({"type": "pong"})

            elif t == "exec":
                result = _exec_cell(
                    msg["chunk_id"],
                    msg.get("code", ""),
                    fig_width  = float(msg.get("fig_width",  7)),
                    fig_height = float(msg.get("fig_height", 5)),
                    dpi        = int(msg.get("dpi", 120)),
                )
                _send(result)

            elif t == "reset":
                _ns.clear()
                _send({"type": "result", "chunk_id": "__reset__", "console": "",
                       "stdout": "", "stderr": "", "plots": [],
                       "plots_html": [], "dataframes": [], "error": None})

            elif t == "vars":
                _send(_vars_result())

        except KeyboardInterrupt:
            # SIGINT arrived outside _exec_cell (e.g. during reset/vars).
            # Send an interrupted result so the TS side can resolve the pending
            # promise, then continue the main loop.
            chunk_id = msg.get("chunk_id", "")
            _send({"type": "result", "chunk_id": chunk_id, "console": "",
                   "stdout": "", "stderr": "", "plots": [],
                   "plots_html": [], "dataframes": [], "error": "Interrupted by user"})


if __name__ == "__main__":
    main()
