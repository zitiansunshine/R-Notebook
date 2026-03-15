#!/usr/bin/env Rscript
# =============================================================================
# RNotebook R Kernel — persistent subprocess protocol
#
# Communicates via newline-delimited JSON on stdin/stdout.
# Each message is a single JSON object terminated by "\n".
#
# INCOMING message types:
#   { "type": "exec",    "chunk_id": "c1", "code": "..." }
#   { "type": "df_page", "chunk_id": "c1", "name": "mtcars", "page": 2, "page_size": 20 }
#   { "type": "complete","chunk_id": "c1", "code": "...", "cursor_pos": 10 }
#   { "type": "restore_workspace", "workspace_state": "..." }
#   { "type": "reset"  }    -- clears user environment
#   { "type": "snapshot" }  -- returns user var names
#   { "type": "ping"   }
#
# OUTGOING message types:
#   { "type": "result",  "chunk_id": "c1", "stdout": "...", "stderr": "...",
#     "plots": ["base64..."], "dataframes": [...], "error": null }
#   { "type": "df_data", "chunk_id": "c1", "name": "...", "page": 2,
#     "nrow": 32, "ncol": 4, "columns": [...], "data": [[...]] }
#   { "type": "complete_result", "chunk_id": "c1", "completions": [...] }
#   { "type": "snapshot_result", "names": [...], "workspace_state": "..." }
#   { "type": "workspace_restored" }
#   { "type": "pong" }
#   { "type": "error",   "chunk_id": "c1", "message": "..." }
# =============================================================================

suppressPackageStartupMessages({
  library(jsonlite)
})

# Null-coalescing operator
`%||%` <- function(a, b) if (!is.null(a)) a else b

# ---- helpers ----------------------------------------------------------------

# Capture the REAL stdout connection before any sink() calls change stdout().
# Inside capture_output(), stdout() returns the active sink, so progress_cb
# (and all kernel I/O) must write to this pre-captured connection instead.
REAL_CON_OUT <- stdout()

send_message <- function(msg) {
  cat(toJSON(msg, auto_unbox = TRUE, null = "null"), "\n", sep = "", file = REAL_CON_OUT)
  flush(REAL_CON_OUT)
}

# Capture output from R code, evaluating each statement individually so that
# visible results are auto-printed exactly as the R console would show them.
capture_output <- function(expr_text, progress_cb = NULL, stream_cb = NULL,
                           plot_cb = NULL, df_cb = NULL,
                           fig_w = 7, fig_h = 5, fig_dpi = 120) {
  stdout_lines <- character(0)
  stderr_lines <- character(0)
  stderr_condition_lines <- character(0)
  console_parts <- character(0)
  console_segments <- list()
  stdout_con   <- textConnection("stdout_lines", "w", local = TRUE)
  stderr_con   <- textConnection("stderr_lines", "w", local = TRUE)
  stdout_sent  <- 0L
  stderr_sent  <- 0L

  sink(stdout_con, type = "output", split = TRUE)
  sink(stderr_con, type = "message")

  last_value   <- NULL
  last_visible <- FALSE
  error_val    <- NULL
  # inline_plots: list of raw base64 PNG strings, in encounter order.
  # inline_dfs_serialized: list of serialised data.frame results, in encounter order.
  # inline_order: list of {type, index, name} for natural interleaved ordering (0-indexed).
  inline_plots          <- list()
  inline_dfs_serialized <- list()
  inline_order          <- list()

  append_console_segment_output <- function(text) {
    if (!nzchar(text) || length(console_segments) == 0L) return(invisible(NULL))
    idx <- length(console_segments)
    current_output <- console_segments[[idx]]$output %||% ""
    separator <- if (nzchar(current_output)) "\n" else ""
    console_segments[[idx]]$output <<- paste0(current_output, separator, text)
    invisible(NULL)
  }

  flush_streams <- function() {
    if (length(stdout_lines) > stdout_sent) {
      raw_text <- paste(stdout_lines[(stdout_sent + 1L):length(stdout_lines)], collapse = "\n")
      stream_text <- paste0(
        if (length(console_parts) > 0L) "\n" else "",
        raw_text
      )
      console_parts <<- c(console_parts, stream_text)
      append_console_segment_output(raw_text)
      if (!is.null(stream_cb)) {
        tryCatch(stream_cb("stdout", raw_text), error = function(e) NULL)
      }
      stdout_sent <<- length(stdout_lines)
    }
    if (length(stderr_lines) > stderr_sent) {
      raw_text <- paste(stderr_lines[(stderr_sent + 1L):length(stderr_lines)], collapse = "\n")
      stream_text <- paste0(
        if (length(console_parts) > 0L) "\n" else "",
        raw_text
      )
      console_parts <<- c(console_parts, stream_text)
      append_console_segment_output(raw_text)
      if (!is.null(stream_cb)) {
        tryCatch(stream_cb("stderr", raw_text), error = function(e) NULL)
      }
      stderr_sent <<- length(stderr_lines)
    }
  }

  emit_stderr_condition <- function(text) {
    stderr_condition_lines <<- c(stderr_condition_lines, text)
    stream_text <- paste0(if (length(console_parts) > 0L) "\n" else "", text)
    console_parts <<- c(console_parts, stream_text)
    append_console_segment_output(text)
    if (!is.null(stream_cb)) {
      tryCatch(stream_cb("stderr", text), error = function(e) NULL)
    }
  }

  parse_err <- NULL
  parsed    <- tryCatch(
    parse(text = expr_text),
    error = function(e) { parse_err <<- conditionMessage(e); NULL }
  )

  if (!is.null(parse_err)) {
    error_val <- paste("Parse error:", parse_err)
  } else if (length(parsed) > 0) {
    for (i in seq_along(parsed)) {
      expr_code <- paste(deparse(parsed[[i]], width.cutoff = 500L), collapse = "\n")
      console_segments[[length(console_segments) + 1L]] <- list(code = expr_code, output = "")
      # Send progress before each statement (bypasses sink via file=stdout())
      if (!is.null(progress_cb)) {
        tryCatch(progress_cb(i, length(parsed), expr_code), error = function(e) NULL)
      }
      vis <- tryCatch(
        withCallingHandlers(
          withVisible(eval(parsed[[i]], envir = .GlobalEnv)),
          packageStartupMessage = function(m) {
            emit_stderr_condition(conditionMessage(m))
            tryInvokeRestart("muffleMessage")
          },
          message = function(m) {
            emit_stderr_condition(conditionMessage(m))
            tryInvokeRestart("muffleMessage")
          },
          warning = function(w) {
            emit_stderr_condition(conditionMessage(w))
            tryInvokeRestart("muffleWarning")
          }
        ),
        error = function(e) list(
          value   = structure(conditionMessage(e), class = "exec_error"),
          visible = FALSE
        )
      )
      if (inherits(vis$value, "exec_error")) {
        error_val    <- as.character(vis$value)
        last_value   <- vis$value
        last_visible <- FALSE
        flush_streams()
        break
      }
      if (isTRUE(vis$visible)) {
        nm <- tryCatch(deparse(parsed[[i]]), error = function(e) paste0("expr", i))
        if (is.data.frame(vis$value)) {
          # Serialize and stream immediately; record in inline_order.
          serialized <- tryCatch(
            serialise_df(nm, vis$value, page_size = 2000),
            error = function(e) NULL
          )
          if (!is.null(serialized)) {
            idx <- length(inline_dfs_serialized)  # 0-indexed
            inline_dfs_serialized[[idx + 1L]] <- serialized
            inline_order <- c(inline_order, list(list(type = "df", index = idx, name = nm)))
            if (!is.null(df_cb)) tryCatch(df_cb(serialized, idx), error = function(e) NULL)
          }
        } else if (is.list(vis$value) && should_use_compact_summary(vis$value)) {
          # Large list-like objects can take a very long time to recursively
          # walk or serialise. Show a compact summary instead so follow-up
          # inspection cells stay responsive.
          cat(compact_object_summary_line(vis$value), "\n", sep = "")
        } else if (is_plain_list(vis$value)) {
          visible_outputs <- tryCatch(
            collect_visible_outputs(vis$value, nm),
            error = function(e) list(ordered = list())
          )
          if (has_visible_display_items(visible_outputs)) {
            # Print placeholder summary to console; capture each item in natural order.
            print_display_aware_list(vis$value)
            for (item in visible_outputs$ordered) {
              if (item$kind == "plot") {
                b64 <- render_plot_to_b64(item$value, fig_w, fig_h, fig_dpi)
                if (!is.null(b64)) {
                  idx <- length(inline_plots)  # 0-indexed
                  inline_plots[[idx + 1L]] <- b64
                  inline_order <- c(inline_order, list(list(type = "plot", index = idx, name = item$name)))
                  if (!is.null(plot_cb)) tryCatch(plot_cb(b64, idx, item$name), error = function(e) NULL)
                }
              } else if (item$kind == "df") {
                serialized <- tryCatch(
                  serialise_df(item$name, item$value, page_size = 2000),
                  error = function(e) NULL
                )
                if (!is.null(serialized)) {
                  idx <- length(inline_dfs_serialized)  # 0-indexed
                  inline_dfs_serialized[[idx + 1L]] <- serialized
                  inline_order <- c(inline_order, list(list(type = "df", index = idx, name = item$name)))
                  if (!is.null(df_cb)) tryCatch(df_cb(serialized, idx, item$name), error = function(e) NULL)
                }
              }
            }
          } else {
            # Plain list with no displayable items: use the structured list
            # renderer so large objects do not stall the kernel on raw print().
            print_display_aware_list(vis$value)
          }
        } else if (is_renderable_plot_object(vis$value)) {
          b64 <- render_plot_to_b64(vis$value, fig_w, fig_h, fig_dpi)
          if (is.null(b64)) {
            tryCatch(print(vis$value), error = function(e) NULL)
          } else {
            idx <- length(inline_plots)  # 0-indexed
            inline_plots[[idx + 1L]] <- b64
            inline_order <- c(inline_order, list(list(type = "plot", index = idx, name = nm)))
            if (!is.null(plot_cb)) tryCatch(plot_cb(b64, idx, nm), error = function(e) NULL)
          }
        } else {
          # Auto-print other visible values.
          tryCatch(print(vis$value), error = function(e) NULL)
        }
      }
      last_value   <- vis$value
      last_visible <- vis$visible
      flush_streams()
    }
  }

  flush_streams()

  # Explicit cleanup — no on.exit so we never accidentally double-remove.
  sink(type = "output")
  sink(type = "message")
  close(stdout_con)
  close(stderr_con)

  list(
    value                 = last_value,
    visible               = last_visible,
    console               = paste(console_parts, collapse = ""),
    console_segments      = console_segments,
    stdout                = paste(stdout_lines, collapse = "\n"),
    stderr                = paste(c(stderr_condition_lines, stderr_lines), collapse = "\n"),
    error                 = error_val,
    inline_plots          = inline_plots,
    inline_dfs_serialized = inline_dfs_serialized,
    inline_order          = inline_order
  )
}

is_plain_list <- function(obj) {
  is.list(obj) && !is.data.frame(obj) && (is.null(class(obj)) || identical(class(obj), "list"))
}

is_renderable_plot_object <- function(obj) {
  inherits(obj, "ggplot") ||
    inherits(obj, "trellis") ||
    inherits(obj, "recordedplot") ||
    inherits(obj, c("grob", "gTree", "gList", "gtable")) ||
    inherits(obj, "nativeRaster") ||
    isTRUE(tryCatch(is.raster(obj), error = function(e) FALSE)) ||
    inherits(obj, "magick-image")
}

has_visible_display_items <- function(outputs) {
  length(outputs$ordered %||% list()) > 0L
}

collect_visible_outputs <- function(value, label, max_depth = 4L, max_items = 256L) {
  ordered_items <- list()
  found <- 0L

  walk <- function(obj, path, depth) {
    if (found >= max_items || depth < 0L || is.null(obj) || is.function(obj) || is.environment(obj)) {
      return(invisible(NULL))
    }
    if (is.data.frame(obj)) {
      found <<- found + 1L
      ordered_items[[length(ordered_items) + 1L]] <<- list(kind = "df", name = path, value = obj)
      return(invisible(NULL))
    }
    if (is_renderable_plot_object(obj)) {
      found <<- found + 1L
      ordered_items[[length(ordered_items) + 1L]] <<- list(kind = "plot", name = path, value = obj)
      return(invisible(NULL))
    }
    if (depth == 0L) return(invisible(NULL))

    if (is_plain_list(obj)) {
      obj_names <- names(obj)
      for (idx in seq_along(obj)) {
        if (found >= max_items) break
        child_name <- if (!is.null(obj_names) && idx <= length(obj_names) && nzchar(obj_names[[idx]] %||% "")) {
          paste0(path, "$", obj_names[[idx]])
        } else {
          paste0(path, "[[", idx, "]]")
        }
        walk(obj[[idx]], child_name, depth - 1L)
      }
      return(invisible(NULL))
    }
  }

  walk(value, label, as.integer(max_depth))
  list(ordered = ordered_items)
}

write_indented_lines <- function(lines, indent = "") {
  if (length(lines) == 0L) {
    cat(indent, "\n", sep = "")
    return(invisible(NULL))
  }
  for (line in lines) {
    cat(indent, line, "\n", sep = "")
  }
  invisible(NULL)
}

compact_object_summary_line <- function(item) {
  cls <- class(item)
  cls_label <- if (length(cls) > 0L) paste(cls, collapse = "/") else typeof(item)
  dims <- tryCatch(dim(item), error = function(e) NULL)

  if (inherits(item, "Seurat")) {
    if (!is.null(dims) && length(dims) >= 2L) {
      return(paste0("<Seurat: ", dims[[1]], " features x ", dims[[2]], " cells>"))
    }
    return("<Seurat>")
  }

  if (!is.null(dims) && length(dims) > 0L) {
    return(paste0("<", cls_label, ": ", paste(dims, collapse = " x "), ">"))
  }

  if (isS4(item)) {
    slots <- tryCatch(length(slotNames(item)), error = function(e) NA_integer_)
    if (is.finite(slots)) {
      return(paste0("<", cls_label, ": ", slots, " slots>"))
    }
    return(paste0("<", cls_label, ">"))
  }

  if (is.atomic(item) || is.factor(item)) {
    len <- length(item)
    if (len == 0L) return(paste0("<", cls_label, ": empty>"))
    if (len == 1L) return(as.character(item))
    if (len <= 6L) return(paste(as.character(item), collapse = " "))
    return(paste0("<", cls_label, ": length ", len, ">"))
  }

  if (is.list(item)) {
    return(paste0("<", cls_label, ": length ", length(item), ">"))
  }

  paste0("<", cls_label, ">")
}

should_use_compact_summary <- function(item) {
  if (is.data.frame(item) || is_renderable_plot_object(item)) return(FALSE)
  len <- tryCatch(length(item), error = function(e) NA_integer_)
  is.finite(len) && len > 100L
}

print_display_aware_list <- function(obj, indent = "", depth = 0L, max_depth = 4L) {
  if (!is_plain_list(obj)) {
    tryCatch(print(obj), error = function(e) NULL)
    return(invisible(NULL))
  }
  if (length(obj) == 0L) {
    cat(indent, "<empty list>\n", sep = "")
    return(invisible(NULL))
  }
  if (depth >= max_depth) {
    cat(indent, "<list>\n", sep = "")
    return(invisible(NULL))
  }

  obj_names <- names(obj)
  for (idx in seq_along(obj)) {
    header <- if (!is.null(obj_names) && idx <= length(obj_names) && nzchar(obj_names[[idx]] %||% "")) {
      paste0(indent, "$", obj_names[[idx]])
    } else {
      paste0(indent, "[[", idx, "]]")
    }
    item <- obj[[idx]]
    child_indent <- paste0(indent, "  ")

    cat(header, "\n", sep = "")

    if (is.data.frame(item)) {
      cat(child_indent, "<data.frame: ", nrow(item), " x ", ncol(item), ">\n", sep = "")
      next
    }
    if (is_renderable_plot_object(item)) {
      plot_class <- class(item)[1] %||% "plot"
      cat(child_indent, "<plot: ", plot_class, ">\n", sep = "")
      next
    }
    if (is_plain_list(item)) {
      print_display_aware_list(item, child_indent, depth + 1L, max_depth)
      next
    }

    if (should_use_compact_summary(item)) {
      cat(child_indent, compact_object_summary_line(item), "\n", sep = "")
      next
    }

    rendered_lines <- tryCatch(
      capture.output(print(item)),
      error = function(e) paste0("<print error: ", conditionMessage(e), ">")
    )
    write_indented_lines(rendered_lines, child_indent)
  }

  invisible(NULL)
}

render_plot_to_b64 <- function(value, fig_w = 7, fig_h = 5, fig_dpi = 120) {
  if (!requireNamespace("base64enc", quietly = TRUE)) return(NULL)
  tmp <- tempfile(fileext = ".png")
  on.exit(unlink(tmp), add = TRUE)
  opened <- tryCatch({
    png(tmp, width = fig_w, height = fig_h, res = fig_dpi, units = "in")
    TRUE
  }, error = function(e) FALSE)
  if (!opened) return(NULL)
  ok <- tryCatch({
    render_visible_plot_object(value)
    dev.off()
    TRUE
  }, error = function(e) {
    try(dev.off(), silent = TRUE)
    FALSE
  })
  if (!ok || !file.exists(tmp) || file.info(tmp)$size == 0) return(NULL)
  base64enc::base64encode(tmp)
}

render_visible_plot_object <- function(obj) {
  if (inherits(obj, "recordedplot")) {
    replayPlot(obj)
    return(TRUE)
  }
  if (inherits(obj, "ggplot") || inherits(obj, "trellis") || inherits(obj, "magick-image")) {
    print(obj)
    return(TRUE)
  }
  if (inherits(obj, c("grob", "gTree", "gList", "gtable"))) {
    if (!requireNamespace("grid", quietly = TRUE)) return(FALSE)
    grid::grid.newpage()
    grid::grid.draw(obj)
    return(TRUE)
  }
  if (inherits(obj, "nativeRaster") || isTRUE(tryCatch(is.raster(obj), error = function(e) FALSE))) {
    if (!requireNamespace("grid", quietly = TRUE)) return(FALSE)
    grid::grid.newpage()
    grid::grid.raster(obj)
    return(TRUE)
  }
  FALSE
}

render_visible_plot_items <- function(plot_items) {
  if (length(plot_items) == 0L) return(invisible(NULL))
  for (plot_item in plot_items) {
    tryCatch(render_visible_plot_object(plot_item$value), error = function(e) FALSE)
  }
  invisible(NULL)
}

# Serialise a data.frame to a paged JSON-safe structure
serialise_df <- function(df_name, df, page = 0, page_size = 50) {
  nrow_total <- nrow(df)
  ncol_total <- ncol(df)
  start      <- page * page_size + 1
  end        <- min(start + page_size - 1, nrow_total)
  column_view_threshold <- 150L
  column_view_limit <- 150L

  visible_df <- if (ncol_total > column_view_threshold) {
    df[, seq_len(min(column_view_limit, ncol_total)), drop = FALSE]
  } else {
    df
  }
  slice <- visible_df[start:end, , drop = FALSE]
  columns <- lapply(seq_along(slice), function(i) {
    list(name = names(slice)[i], type = class(slice[[i]])[1])
  })
  row_names <- rownames(slice)
  if (is.null(row_names)) {
    row_names <- as.character(seq.int(start, end))
  } else {
    row_names <- as.character(row_names)
  }
  row_names <- unname(as.list(row_names))

  data_rows <- lapply(seq_len(nrow(slice)), function(i) {
    lapply(slice[i, , drop = FALSE], function(v) {
      if (is.na(v)) NULL
      else if (is.numeric(v)) signif(v, 6)
      else as.character(v)
    })
  })

  list(
    name    = df_name,
    nrow    = nrow_total,
    ncol    = ncol_total,
    pages   = ceiling(nrow_total / page_size),
    page    = page,
    row_names = row_names,
    columns = columns,
    data    = data_rows
  )
}

snapshot_workspace_state <- function(checkpoint_path = NULL) {
  user_vars <- setdiff(ls(envir = .GlobalEnv), KERNEL_RESERVED)
  attached_packages <- setdiff(grep("^package:", search(), value = TRUE), KERNEL_ATTACHED_PACKAGES)
  if (length(user_vars) == 0 && length(attached_packages) == 0) return(NULL)

  user_objects <- if (length(user_vars) > 0) {
    mget(user_vars, envir = .GlobalEnv, inherits = FALSE)
  } else {
    list()
  }

  payload <- list(
    objects = user_objects,
    attached_packages = sub("^package:", "", attached_packages)
  )
  if (!is.null(checkpoint_path) && nzchar(checkpoint_path)) {
    dir.create(dirname(checkpoint_path), recursive = TRUE, showWarnings = FALSE)
    temp_checkpoint_path <- paste0(
      checkpoint_path,
      ".tmp-",
      Sys.getpid(),
      "-",
      sprintf("%06d", sample.int(999999, 1))
    )
    on.exit(unlink(temp_checkpoint_path, force = TRUE), add = TRUE)
    saveRDS(payload, temp_checkpoint_path, compress = FALSE)
    if (!file.rename(temp_checkpoint_path, checkpoint_path)) {
      stop("Failed to replace workspace checkpoint")
    }
    return(invisible(checkpoint_path))
  }
  jsonlite::base64_enc(serialize(payload, NULL, xdr = FALSE))
}

restore_workspace_state <- function(workspace_state = NULL, checkpoint_path = NULL) {
  user_vars <- setdiff(ls(envir = .GlobalEnv), KERNEL_RESERVED)
  if (length(user_vars) > 0) rm(list = user_vars, envir = .GlobalEnv)
  if (!is.null(checkpoint_path) && nzchar(checkpoint_path)) {
    if (!file.exists(checkpoint_path)) stop("Workspace checkpoint missing")
    restored <- readRDS(checkpoint_path)
  } else {
    if (is.null(workspace_state) || !nzchar(workspace_state)) return(invisible(NULL))
    restored <- unserialize(jsonlite::base64_dec(workspace_state))
  }
  if (!is.list(restored)) stop("Invalid workspace snapshot")

  attached_packages <- character(0)
  objects <- restored

  if (!is.null(names(restored)) && all(c("objects", "attached_packages") %in% names(restored))) {
    attached_packages <- restored[["attached_packages"]] %||% character(0)
    objects <- restored[["objects"]] %||% list()
  }

  if (!is.list(objects) || (length(objects) > 0 && is.null(names(objects)))) {
    stop("Invalid workspace snapshot")
  }

  if (length(attached_packages) > 0) {
    for (pkg in rev(attached_packages)) {
      suppressPackageStartupMessages(
        library(pkg, character.only = TRUE, quietly = TRUE, warn.conflicts = FALSE)
      )
    }
  }

  if (length(objects) > 0) list2env(objects, envir = .GlobalEnv)
  invisible(NULL)
}

# ---- message dispatcher -----------------------------------------------------
# All per-message state lives inside this function so it is LOCAL, not in
# .GlobalEnv.  This means a "reset" that clears .GlobalEnv can never
# accidentally delete chunk_id, type, code, etc.

process_message <- function(msg) {
  type     <- msg[["type"]]
  chunk_id <- msg[["chunk_id"]] %||% ""

  # ---------- ping -----------------------------------------------------------
  if (type == "ping") {
    send_message(list(type = "pong"))
    return(invisible(NULL))
  }

  # ---------- reset ----------------------------------------------------------
  if (type == "reset") {
    # Only remove variables that the USER created, not kernel internals.
    user_vars <- setdiff(ls(envir = .GlobalEnv), KERNEL_RESERVED)
    if (length(user_vars) > 0) rm(list = user_vars, envir = .GlobalEnv)
    # chunk_id is LOCAL here — safe even after rm().
    send_message(list(type = "result", chunk_id = chunk_id,
                      console = "",
                      stdout = "", stderr = "", plots = list(),
                      dataframes = list(), error = NULL))
    return(invisible(NULL))
  }

  # ---------- snapshot -------------------------------------------------------
  if (type == "snapshot") {
    user_vars <- setdiff(ls(envir = .GlobalEnv), KERNEL_RESERVED)
    attached_packages <- setdiff(grep("^package:", search(), value = TRUE), KERNEL_ATTACHED_PACKAGES)
    had_state <- length(user_vars) > 0 || length(attached_packages) > 0
    snapshot_ok <- TRUE
    checkpoint_path <- msg[["checkpoint_path"]] %||% NULL
    workspace_state <- NULL
    tryCatch(
      {
        workspace_state <- snapshot_workspace_state(checkpoint_path)
      },
      error = function(e) {
        snapshot_ok <<- FALSE
      }
    )
    send_message(list(
      type = "snapshot_result",
      names = as.list(user_vars),
      workspace_state = workspace_state,
      captured = snapshot_ok,
      had_state = had_state
    ))
    return(invisible(NULL))
  }

  # ---------- restore_workspace ---------------------------------------------
  if (type == "restore_workspace") {
    restored_ok <- TRUE
    tryCatch(
      restore_workspace_state(
        workspace_state = msg[["workspace_state"]] %||% NULL,
        checkpoint_path = msg[["checkpoint_path"]] %||% NULL
      ),
      error = function(e) {
        restored_ok <<- FALSE
        send_message(list(type = "error", chunk_id = "__restore_workspace__",
                          message = paste("Workspace restore failed:", conditionMessage(e))))
      }
    )
    if (restored_ok) send_message(list(type = "workspace_restored"))
    return(invisible(NULL))
  }

  # ---------- df_page --------------------------------------------------------
  if (type == "df_page") {
    df_name   <- msg[["name"]]
    page      <- as.integer(msg[["page"]]      %||% 0)
    page_size <- as.integer(msg[["page_size"]] %||% 50)

    if (!exists(df_name, envir = .GlobalEnv, inherits = FALSE)) {
      send_message(list(type = "error", chunk_id = chunk_id,
                        message = paste("Object not found:", df_name)))
      return(invisible(NULL))
    }
    df_obj <- get(df_name, envir = .GlobalEnv, inherits = FALSE)
    if (!is.data.frame(df_obj)) {
      send_message(list(type = "error", chunk_id = chunk_id,
                        message = paste(df_name, "is not a data.frame")))
      return(invisible(NULL))
    }
    result <- serialise_df(df_name, df_obj, page, page_size)
    send_message(c(list(type = "df_data", chunk_id = chunk_id), result))
    return(invisible(NULL))
  }

  # ---------- vars -----------------------------------------------------------
  if (type == "vars") {
    user_vars <- setdiff(ls(envir = .GlobalEnv), KERNEL_RESERVED)
    var_info <- lapply(user_vars, function(nm) {
      obj <- tryCatch(get(nm, envir = .GlobalEnv, inherits = FALSE),
                     error = function(e) NULL)
      if (is.null(obj)) return(list(name = nm, type = "NULL", size = "", value = "NULL"))

      cls <- class(obj)[1]

      sz <- tryCatch({
        if (is.data.frame(obj))                    paste0(nrow(obj), " \u00d7 ", ncol(obj))
        else if (is.matrix(obj))                   paste0(nrow(obj), " \u00d7 ", ncol(obj))
        else if (is.array(obj) && length(dim(obj)) > 1)
                                                   paste(dim(obj), collapse = " \u00d7 ")
        else if (is.list(obj) && !is.function(obj)) paste0("length ", length(obj))
        else if (is.vector(obj) || is.factor(obj)) as.character(length(obj))
        else ""
      }, error = function(e) "")

      val <- tryCatch({
        if (is.function(obj)) {
          fmls <- names(formals(obj))
          if (length(fmls) == 0) "function()" else paste0("function(", paste(fmls, collapse = ", "), ")")
        } else if (is.data.frame(obj) || is.matrix(obj)) {
          paste0("[", cls, "]")
        } else if (is.list(obj)) {
          paste0("[list:", length(obj), "]")
        } else if (length(obj) == 1) {
          as.character(obj)
        } else if (length(obj) <= 6) {
          paste(as.character(obj), collapse = " ")
        } else {
          paste0(paste(as.character(obj[1:5]), collapse = " "), " \u2026")
        }
      }, error = function(e) "?")

      list(name = nm, type = cls, size = sz, value = val)
    })
    send_message(list(type = "vars_result", vars = var_info))
    return(invisible(NULL))
  }

  # ---------- complete -------------------------------------------------------
  if (type == "complete") {
    code       <- msg[["code"]]       %||% ""
    cursor_pos <- as.integer(msg[["cursor_pos"]] %||% nchar(code))
    completions <- tryCatch({
      utils:::.completeToken
      utils:::.assignLinebuffer(substr(code, 1, cursor_pos))
      utils:::.assignEnd(cursor_pos)
      utils:::.guessTokenFromLine()
      utils:::.completeToken()
      utils:::.retrieveCompletions()
    }, error = function(e) character(0))
    send_message(list(type = "complete_result", chunk_id = chunk_id,
                      completions = as.list(completions)))
    return(invisible(NULL))
  }

  # ---------- exec -----------------------------------------------------------
  if (type == "exec") {
    code <- msg[["code"]] %||% ""

    tryCatch({
      stdout_text  <- ""
      stderr_text  <- ""
      error_msg    <- NULL

      # Chunk dimensions from message options (fig.width/height in inches, dpi).
      fig_w   <- as.numeric(msg[["fig_width"]]  %||% 7)
      fig_h   <- as.numeric(msg[["fig_height"]] %||% 5)
      fig_dpi <- as.numeric(msg[["dpi"]]        %||% 120)

      # Progress callback — send_message uses file=stdout() which bypasses sink.
      local_chunk <- chunk_id
      progress_cb <- function(line, total, expr_code = NULL) {
        send_message(list(type = "progress", chunk_id = local_chunk,
                          line = line, total = total, expr_code = expr_code))
      }
      stream_cb <- function(stream, text) {
        send_message(list(type = "stream", chunk_id = local_chunk,
                          stream = stream, text = text))
      }

      # Per-plot/df callbacks — stream each output immediately as it is produced.
      plot_cb <- function(b64, idx, name = NULL) {
        send_message(list(type = "stream_output", chunk_id = local_chunk,
                          kind = "plot", index = idx, name = name, b64 = b64))
      }
      df_cb <- function(serialized, idx, name = NULL) {
        send_message(list(type = "stream_output", chunk_id = local_chunk,
                          kind = "df", index = idx, name = name, df = serialized))
      }

      # Execute code — each plot/df is rendered and streamed immediately.
      out         <- capture_output(code, progress_cb = progress_cb, stream_cb = stream_cb,
                                    plot_cb = plot_cb, df_cb = df_cb,
                                    fig_w = fig_w, fig_h = fig_h, fig_dpi = fig_dpi)
      stdout_text <- out$stdout
      stderr_text <- out$stderr
      error_msg   <- out$error

      send_message(list(
        type         = "result",
        chunk_id     = chunk_id,
        console      = out$console,
        console_segments = out$console_segments,
        stdout       = stdout_text,
        stderr       = stderr_text,
        plots        = out$inline_plots,
        dataframes   = out$inline_dfs_serialized,
        output_order = if (length(out$inline_order) > 0) out$inline_order else NULL,
        error        = error_msg
      ))
    }, error = function(e) {
      # Safety net: catch any unexpected error in the exec path and report it
      # rather than crashing the repeat loop.
      send_message(list(
        type       = "result",
        chunk_id   = chunk_id,
        console    = "",
        stdout     = "",
        stderr     = "",
        plots      = list(),
        dataframes = list(),
        error      = paste("Internal kernel error:", conditionMessage(e))
      ))
    })
    return(invisible(NULL))
  }

  # Unknown type
  send_message(list(type = "error", chunk_id = chunk_id,
                    message = paste("Unknown message type:", type)))
}

# ---- main loop --------------------------------------------------------------

message("R Notebook R Kernel ready")
flush(stderr())

# Open an explicit connection to fd 0 (stdin).  Using stdin() directly in
# Rscript mode on macOS returns EOF immediately because R has already
# consumed the script via the same file descriptor.
con <- file("stdin", open = "r")
on.exit(close(con), add = TRUE)

# Snapshot kernel-defined names AFTER con is created so that reset() can
# distinguish kernel internals from user-created variables.
# "KERNEL_RESERVED" is added to the list so it protects itself. "line" and
# "msg" are loop-local protocol temporaries that are assigned in .GlobalEnv.
# "KERNEL_ATTACHED_PACKAGES" is the baseline search-path package set that
# should survive interrupt recovery without being reattached.
KERNEL_RESERVED <- c(ls(), "KERNEL_RESERVED", "line", "msg")
KERNEL_ATTACHED_PACKAGES <- grep("^package:", search(), value = TRUE)

repeat {
  # Wrap readLines so that a stray SIGINT arriving between messages (e.g. a
  # second rapid signal from the host) is silently ignored rather than
  # propagating as an unhandled interrupt that kills the process.
  line <- tryCatch(
    readLines(con, n = 1, warn = FALSE),
    error     = function(e) NULL,
    interrupt = function(e) character(0)
  )
  if (is.null(line) || length(line) == 0) break
  line <- trimws(line)
  if (nchar(line) == 0) next

  msg <- tryCatch(fromJSON(line, simplifyVector = FALSE), error = function(e) NULL)
  if (is.null(msg)) next

  # Catch R's interrupt condition (SIGINT) so the kernel loop survives.
  # Without this, SIGINT in --slave mode terminates the process entirely.
  tryCatch(
    process_message(msg),
    interrupt = function(e) {
      # Clean up any lingering sinks from interrupted execution
      tryCatch(while (sink.number()               > 0) sink(),              error = function(e) NULL)
      tryCatch(while (sink.number(type="message") > 0) sink(type="message"), error = function(e) NULL)
      # Close any open graphics device
      tryCatch(if (dev.cur() > 1) dev.off(), error = function(e) NULL)
      chunk_id <- msg[["chunk_id"]] %||% ""
      send_message(list(type = "result", chunk_id = chunk_id,
                        console = "",
                        stdout = "", stderr = "", plots = list(),
                        dataframes = list(), error = "Interrupted by user"))
    }
  )
}
