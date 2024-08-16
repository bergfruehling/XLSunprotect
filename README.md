# XLSunprotect
XLSunprotect is a commandline tool written in go to remove the workbook and worksheet protection from Excel files (.xlsx). It is powered by [excelize](https://github.com/360EntSecGroup-Skylar/excelize).

Powered by go, it is compiled and statically linked to a single binary. It does not have any dependencies (Python, JRE, bash, ...). Just distribute the .exe file.

- The results are written to `<filename>_unprotected.xlsx` in the same folder
- The original file remains unchanged

## Download
[Windows](https://github.com/bergfruehling/XLSunprotect/releases/download/v1.0/xlsunprotect.exe)
[Linux](https://github.com/bergfruehling/XLSunprotect/releases/download/v1.0/xlsunprotect)

## Compile
If you want to compile on your own, install [Go](https://golang.org/).
The repository contains the go-module information. All you need is to run
```
go build
```

## Use
You can use the tool from the commandline:
```
> unprotect.exe test.xlsx
===================== 💾 test.xlsx 💾 =====================
ℹ️ Removing protection from test.xlsx
ℹ️ Unprotecting Tabelle1 ...
ℹ️ Unprotecting Sheet1 ...
ℹ️ Removing workbook protection...
✅ Done --> Output in test_unprotected.xlsx

[Press any key to close...]
```
When you want to unprotect multiple files, simply add their filenames as additional parameters.

You can also drag-and-drop Excel-files to the binary.

## License (MIT)
Copyright 2021-2024 Henning Carstens

Use of this source code is governed by the MIT license that can be found in the LICENSE file.
