package main

import (
	"fmt"
	"os"
	"time"

	"github.com/mazzegi/xlsxscan"
)

func main() {
	var file string
	if len(os.Args) < 2 {
		file = "testsheet.xlsx"
	} else {
		file = os.Args[1]
	}

	t0 := time.Now()
	t, err := xlsxscan.OpenFile(file)
	if err != nil {
		fmt.Println("ERRRO: open file:", err)
		os.Exit(1)
	}

	sr, err := t.OpenSheetByName("Sheet1")
	if err != nil {
		fmt.Println("ERRRO: open sheet:", err)
		os.Exit(1)
	}

	rs, err := sr.OpenRowScanner()
	if err != nil {
		fmt.Println("ERRRO: open row scanner:", err)
		os.Exit(1)
	}
	defer rs.Close()

	for {
		row, ok := rs.Scan()
		if !ok {
			break
		}
		fmt.Println(row.String())
	}

	fmt.Println("done in", time.Since(t0))
}
