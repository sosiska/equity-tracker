package main

import (
	"fmt"

	"github.com/tealeg/xlsx/v3"
)

func main() {
	wb, err := xlsx.OpenFile("./sample.xlsx")
	if err != nil {
		panic(err)
	}
	fmt.Println("Sheets in this file:")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("----")
}
