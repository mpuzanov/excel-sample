package main

import (
	"excel-sample/internal/domain/model"
	"excel-sample/internal/excelize"
	"excel-sample/internal/xlsx"
	"os"

	"fmt"
	"log"
)

var (
	p   *model.ListPayments
	err error
)

func main() {
	// готовим тестовые данные
	p = model.PrepareTestData(10000)

	fmt.Println("=======  tealeg/xlsx ========")
	fileName, err := xlsx.SaveToExcel1(p, ".", "file1*.xlsx")
	defer os.Remove(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Создали файл %q библиотека xlsx. 1 вариант.\n", fileName)
	pr, err := xlsx.ReadFile(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print(pr.String())

	fileName, err = xlsx.SaveToExcel2(p, ".", "file1*.xlsx")
	defer os.Remove(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Создали файл %q библиотека xlsx. 2 вариант.\n", fileName)

	fmt.Printf("Читаем файл %q библиотека xlsx\n", fileName)
	pr, err = xlsx.ReadFile(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print(pr.String())

	fmt.Println("\n=======  360EntSecGroup-Skylar/excelize ========")

	fileName, err = excelize.SaveToExcel1(p, ".", "file1*.xlsx")
	defer os.Remove(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Создали файл %q библиотека excelize. 1 вариант.\n", fileName)
	fmt.Printf("Читаем файл %q библиотека excelize.\n", fileName)
	pr2, err := excelize.ReadFile(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print(pr2.String())

	fileName, err = excelize.SaveToExcel2(p, ".", "file1*.xlsx")
	defer os.Remove(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Создали файл %q библиотека excelize. 2 вариант.\n", fileName)

	fmt.Printf("Читаем файл %q библиотека excelize.\n", fileName)
	pr2, err = excelize.ReadFile(fileName)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print(pr2.String())
}
