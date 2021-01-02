package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	if len(os.Args) != 2 {
		fmt.Println("This programs removes the protection from all sheets in the XLSX")
		fmt.Println()
		fmt.Println("USAGE: unprotect.exe <filename>.xlsx")
		fmt.Println()
		fmt.Println("- The result is written into <filename>_unprotected.xlsx")
		fmt.Println("- The original file remains unchanged")
		return
	}

	filename := os.Args[1]
	f, err := excelize.OpenFile(filename)
	fatalOnErr(err, "Could not open file: ")

	log.Println("Removing protection from", filename)
	for _, name := range f.GetSheetMap() {
		log.Println("Unprotecting", name, "...")
		err := f.UnprotectSheet(name)
		fatalOnErr(err, "Could not remove protection for "+name)
	}

	log.Println("Removing workbook protection...")
	f.WorkBook.WorkbookProtection.LockStructure = false
	f.WorkBook.WorkbookProtection.LockRevision = false
	f.WorkBook.WorkbookProtection.LockWindows = false

	outputFilename := strings.Replace(filename, ".xlsx", "", -1) + "_unprotected.xlsx"
	err = f.SaveAs(outputFilename)
	fatalOnErr(err, "Could not write output file:")

	log.Println("Done --> Output in", outputFilename)
}

func fatalOnErr(err error, msg string) {
	if err != nil {
		log.Fatal(msg, err)
	}
}
