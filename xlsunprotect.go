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
	if len(os.Args) == 1 {
		fmt.Println("This programs removes the protection from the workbook and all sheets in the XLSX")
		fmt.Println("(2024) Henning Carstens")
		fmt.Println()
		fmt.Println("USAGE: xlsunprotect.exe <filename>.xlsx [<filename2>.xlsx ...]")
		fmt.Println()
		fmt.Println("- The result is written into <filename>_unprotected.xlsx")
		fmt.Println("- If <filename>_unprotected.xlsx already exists, it remains unchanged and the program stops")
		fmt.Println("- The original file remains unchanged")
		return
	}

	for i := 1; i < len(os.Args); i++ {
		println("===================== 💾", os.Args[i], "💾 =====================")
		err := unprotectFile(os.Args[i])
		if err != nil {
			printError("❌ Could not process file", os.Args[i], ":", err)
		}
		println()
	}

	fmt.Println("[Press any key to close...]")
	fmt.Scanf("h")
}

func unprotectFile(filename string) error {
	if strings.HasSuffix(filename, ".xlsb") {
		return errors.New("only .xlsx is supported: Please open .xlsb in Excel first an save as .xlsx")
	} else if !strings.HasSuffix(filename, ".xlsx") {
		return errors.New("only .xlsx is supported")
	}
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return err
	}

	fmt.Println("ℹ️ Removing protection from", filename)
	for _, name := range f.GetSheetMap() {
		fmt.Println("ℹ️ Unprotecting", name, "...")
		f.UnprotectSheet(name)
	}

	if f.WorkBook.WorkbookProtection != nil {
		fmt.Println("ℹ️ Removing workbook protection...")
		f.WorkBook.WorkbookProtection.LockStructure = false
		f.WorkBook.WorkbookProtection.LockRevision = false
		f.WorkBook.WorkbookProtection.LockWindows = false
	}

	outputFilename := strings.Replace(filename, ".xlsx", "", -1) + "_unprotected.xlsx"
	if _, err = os.Stat(outputFilename); os.IsNotExist(err) {
		err = f.SaveAs(outputFilename)
		if err != nil {
			return err
		}
		color.Green("✅ Done --> Output in " + outputFilename)
	} else {
		return errors.New("cannot write to" + outputFilename + " as it already exists.")
	}

	return nil
}
