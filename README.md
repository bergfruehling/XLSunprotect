# XLSunprotect
XLSunprotect is a commandline tool written in go to remove the workbook and worksheet protection from Excel files (.xlsx). It is powered by [excelize](https://github.com/360EntSecGroup-Skylar/excelize).

Powered by go, it is compiled and statically linked to a single binary. It does not have any dependencies (Python, JRE, bash, ...). Just distribute the .exe file.

- The result is written into `<filename>_unprotected.xlsx`
- The original file remains unchanged

## Download
[Here](https://github.com/bergfruehling/XLSunprotect/releases/download/v1.0/unprotect.exe)

## Compile
If you want to compile on your own, install [Go](https://golang.org/).
```
# Get libraries
go get github.com/360EntSecGroup-Skylar/excelize
go get github.com/fatih/color
# Build
go build unprotect.go
```

## Use
```
> unprotect.exe test.xlsx
Removing protection from .\test.xlsx
Unprotecting Tabelle1 ...
Unprotecting Sheet1 ...
Removing workbook protection...
Done --> Output in .\test_unprotected.xlsx
```

## License (MIT)
Copyright 2021 Henning Carstens

Use of this source code is governed by the MIT license that can be found in the LICENSE file.
