package main

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintln(os.Stderr, "usage: excl2csv <file.xlsx> [sheet]")
		os.Exit(1)
	}

	filename := os.Args[1]

	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Fprintf(os.Stderr, "error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	sheet := f.GetSheetName(0)
	if len(os.Args) >= 3 {
		sheet = os.Args[2]
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Fprintf(os.Stderr, "error: %v\n", err)
		os.Exit(1)
	}

	w := csv.NewWriter(os.Stdout)
	for _, row := range rows {
		if err := w.Write(row); err != nil {
			fmt.Fprintf(os.Stderr, "error: %v\n", err)
			os.Exit(1)
		}
	}
	w.Flush()
	if err := w.Error(); err != nil {
		fmt.Fprintf(os.Stderr, "error: %v\n", err)
		os.Exit(1)
	}
}
