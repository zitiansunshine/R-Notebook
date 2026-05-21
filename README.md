# R Notebook

R Notebook is a VSCode extension (https://marketplace.visualstudio.com/items?itemName=zitiansunsh1ne.r-notebook) for `.Rmd`, `.qmd`, and `.ipynb` workflows with in-line execution, notebook-style outputs, plots, and session-aware tooling for R/Python data science analysis.

## Features

- Native `.Rmd`, `.qmd`, and `.ipynb` support with in-line execution and notebook controls
- In-line console output, plot rendering, and ordered multi-output display
- Session controls for selecting, restarting, and interrupting kernels
- Session variable inspection for active notebooks
- Multi-session support so separate notebooks can keep their own runtime state
- Paginated data-frame viewing for larger tabular outputs
- Built to work well alongside AI-assisted editing workflows

## AI Editing Support

R Notebook works well with AI-powered editing and review tools, including GitHub Copilot, Cursor, and Antigravity. You can keep an `.Rmd` or `.qmd` notebook open, ask an AI tool to revise analysis code or documentation, and continue executing chunks with the same in-line notebook experience. 

## Preview

### Editing code with Cursor

![Editing code with Cursor](images/Editing%20codes%20with%20Cursor.png)

### In-line image rendering

![In-line image rendering](images/Support%20for%20in-line%20rendering%20of%20images.png)

### In-line data-frame viewing

![In-line data-frame viewing](images/Support%20for%20in-line%20view%20of%20data%20frames.png)

### Detailed console output

![Detailed console output](images/Detailed%20Console%20Output.png)

### Detailed error panel with Traceback

![Detailed error panel with Traceback](images/Detailed%20Error%20Panel%20with%20TraceBack.png)

### Multiple R sessions

![Multiple R sessions](images/Support%20for%20running%20multiple%20R%20sessions.png)

## Build From Source

```bash
npm install
npm run build
```

## Package

```bash
npm run package
```

Version `1.5.2` packages into the project root as `r-notebook-1.5.2.vsix`.
