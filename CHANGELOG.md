# Changelog

## 1.5.3
- Bug fixes

## 1.5.2
- Updated the descriptions of the extension

## 1.5.1
- Bug fixes

## 1.5.0
- Fixed bugs related to stability and performance issues
- Updated the Error panel to include more details on Traceback

## 1.4.0
- Made automatic variable refresh lazy so chunk execution no longer waits on workspace inspection after every run
- Kept interrupt recovery checkpointing unchanged

## 1.2.9
- Fixed bugs

## 1.2.8
- Fixed `.ipynb` activation on stable Cursor / VS Code builds by skipping unsupported proposed kernel-source registration, which restores notebook toolbar actions
- Fixed `.ipynb` notebook toolbar actions for `Restart Kernel`, `Interrupt Kernel`, `Show Session Variables`, and `Export`
- Displayed dataframe column type labels in `.ipynb` outputs, matching `.Rmd` and `.qmd`

## 1.2.5
- Added top-right manual executable-path entries for Python and R notebook kernel selection
- Improved kernel selection for `.Rmd`, `.qmd`, and `.ipynb` notebooks

## 1.2.4
- Fixed bugs related to selecting Python Kernels when working with `.ipynb` files
- Improved external edit syncing and automatic save handling for large `.Rmd`, `.qmd`, and `.ipynb` notebooks

## 1.2.3
- Fixed bugs

## 1.2.2
- Fixed bugs

## 1.2.1
- Fixed bugs

## 1.2.0
- Added dataframe column type labels to `.Rmd`, `.qmd`, and `.ipynb` outputs
- Fixed bugs
- Added `Export` actions for notebook and custom-editor workflows

## 1.1.0
- Added `.qmd` / Quarto notebook support using the current notebook UI, including Quarto-style `#|` chunk options

## 1.0.2
- Fixed bugs related to the following three toolbar buttons: `Interrupt Kernel`, `Restart Kernel`, and `Show Session Variables`

## 1.0.1
- Fixed bugs
