package main

import (
	"errors"
	"fmt"
	"os"

	"github.com/tealeg/xlsx/v3"
)

func main() {
	rowStuff()
}

var isMoneyMoveBlock bool

func rowVisitor(r *xlsx.Row) error {
	if !isMoneyMoveBlock {
		c := r.GetCell(1)
		value, err := c.FormattedValue()
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
		if value == "Внешнее движение денежных средств в валюте счета" {
			isMoneyMoveBlock = true
			fmt.Println("Дата\t\tСумма\t\t\tОперация\t\tКомментарий")
			return nil
		}
	} else {
		cDate := r.GetCell(1)
		cDateValue, err := cDate.FormattedValue()
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
		switch cDateValue {
		case "Дата ":
			return nil
		case "":
			isMoneyMoveBlock = false
			fmt.Println("Вот и все, ребятки!")
			os.Exit(0)
		}
		// Если мы тут, то значит мы в блоке с движением средств
		cSum := r.GetCell(3)
		cSumValue, err := cSum.FormattedValue()
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
		cOpp := r.GetCell(5)
		cOppValue, err := cOpp.FormattedValue()
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
		cComment := r.GetCell(7)
		cCommentValue, err := cComment.FormattedValue()
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
		fmt.Println(cDateValue + "\t" + cSumValue + "\t\t" + cOppValue + "\t\t" + cCommentValue)
	}
	return nil
}

func rowStuff() {
	filename := "otchet2020.xlsx"
	wb, err := xlsx.OpenFile(filename)
	if err != nil {
		panic(err)
	}
	sh, ok := wb.Sheet["Отчет"]
	if !ok {
		panic(errors.New("Sheet not found"))
	}
	fmt.Println("Max row is", sh.MaxRow)
	sh.ForEachRow(rowVisitor)
}
