package main

import (
	"fmt"
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
	if err != nil {
		fmt.Println("Could not open file:", err)
		return
	}

	fmt.Println("Removing sheet protection from", filename)
	for _, name := range f.GetSheetMap() {
		fmt.Println("Unprotecting", name, "...")
		if err := f.UnprotectSheet(name); err != nil {
			fmt.Println("Could not remove protection for", name, ":", err)
		}
	}

	outputFilename := strings.Replace(filename, ".xlsx", "", -1) + "_unprotected.xlsx"
	if err := f.SaveAs(outputFilename); err != nil {
		fmt.Println("Could not write output file:", err)
		return
	}
	fmt.Println("Done --> Output in", outputFilename)
}
