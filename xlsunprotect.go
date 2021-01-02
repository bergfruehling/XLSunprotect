// Copyright 2021 Henning Carstens. All rights reserved.
// Use of this source code is governed by the MIT license that can be found in
// the LICENSE file.

package main

import (
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/fatih/color"
)

var printError = color.New(color.Bold, color.FgRed).PrintlnFunc()

func main() {
	if len(os.Args) != 2 {
		fmt.Println("This programs removes the protection from the workbook and all sheets in the XLSX")
		fmt.Println("https://github.com/bergfruehling/XLSunprotect")
		fmt.Println()
		fmt.Println("USAGE: xlsunprotect.exe <filename>.xlsx")
		fmt.Println()
		fmt.Println("- The result is written into <filename>_unprotected.xlsx")
		fmt.Println("- The original file remains unchanged")
		return
	}

	filename := os.Args[1]
	f, err := excelize.OpenFile(filename)
	fatalOnErr(err)

	fmt.Println("Removing protection from", filename)
	for _, name := range f.GetSheetMap() {
		fmt.Println("Unprotecting", name, "...")
		err := f.UnprotectSheet(name)
		fatalOnErr(err)
	}

	fmt.Println("Removing workbook protection...")
	f.WorkBook.WorkbookProtection.LockStructure = false
	f.WorkBook.WorkbookProtection.LockRevision = false
	f.WorkBook.WorkbookProtection.LockWindows = false

	outputFilename := strings.Replace(filename, ".xlsx", "", -1) + "_unprotected.xlsx"
	if _, err = os.Stat(outputFilename); os.IsNotExist(err) {
		err = f.SaveAs(outputFilename)
		fatalOnErr(err)
		color.Green("Done --> Output in " + outputFilename)
	} else {
		fatalOnErr(errors.New(outputFilename + " already exists...exiting."))
	}
}

func fatalOnErr(err error) {
	if err != nil {
		printError(err)
		os.Exit(1)
	}
}
