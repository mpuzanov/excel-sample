package xlsx

import (
	"io/ioutil"
	"reflect"
	"strings"
	"time"
	"unicode/utf8"

	"excel-sample/internal/domain/model"

	"github.com/tealeg/xlsx"
)

// ListPayments структура для хранения платежей
type ListPayments model.ListPayments

type fieldExcel struct {
	Name     string
	Position int
	With     int
	Style    int
	Type     string
}

var (
	//headerMap список полей в заголовке
	headerMap = map[int]*fieldExcel{
		0: {Name: "№", With: 10},
		1: {Name: "Fio", With: 20},
		2: {Name: "Date", With: 10, Type: "time.Time"},
		3: {Name: "Участок", With: 10},
		4: {Name: "Occ", With: 10, Type: "int"},
		5: {Name: "Address", With: 40},
		6: {Name: "Value", With: 10, Type: "float64"},
		7: {Name: "Код_услуги", With: 10},
		8: {Name: "Commission", With: 10, Type: "float64"},
		9: {Name: "PaymentAccount", With: 25},
	}
	headerName = make(map[string]*fieldExcel, 9)
)

func init() {
	for i := 0; i < len(headerMap); i++ {
		headerName[headerMap[i].Name] = &fieldExcel{Name: headerMap[i].Name, Position: i, With: headerMap[i].With}
	}
}

// SaveToExcel1 сохраняем данные в файл
// используем пакет reflect для определения полей структуры
func SaveToExcel1(s *model.ListPayments, path, templateFile string) (string, error) {

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	if templateFile == "" {
		templateFile = "file*.xlsx"
	}
	tmpfile, err := ioutil.TempFile(path, templateFile)
	if err != nil {
		return "", err
	}
	defer tmpfile.Close()
	fileName := tmpfile.Name()

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		return "", err
	}
	headerFont := xlsx.NewFont(12, "Calibri")
	headerFont.Bold = true
	headerFont.Underline = false
	headerStyle := xlsx.NewStyle()
	headerStyle.Font = *headerFont

	dataFont := xlsx.NewFont(11, "Calibri")
	dataStyle := xlsx.NewStyle()
	dataStyle.Font = *dataFont

	//заполняем заголовок
	row = sheet.AddRow()
	for index := 0; index < len(headerMap); index++ {
		cell = row.AddCell()
		cell.Value = headerMap[index].Name
		cell.SetStyle(headerStyle)
	}

	//данные
	for index := 0; index < len(s.Db); index++ {
		row = sheet.AddRow()
		// добавляем поля в строке
		values := reflect.ValueOf(s.Db[index])

		for i := 0; i < len(headerMap); i++ {
			cell = row.AddCell()
			f := values.FieldByName(strings.Title(headerMap[i].Name))
			if f.IsValid() {
				fieldValue := f.Interface()
				switch v := fieldValue.(type) {
				case float64:
					cell.SetFloatWithFormat(v, "#,##0.00")
				case int:
					cell.SetInt(int(v))
				case time.Time:
					cell.SetDate(v)
				default:
					cell.SetValue(v)
				}
				cell.SetStyle(dataStyle)
			}
		}
	}
	//Устанавливаем ширину колонок
	for i, col := range sheet.Cols {
		col.Width = float64(headerMap[i].With)
	}

	err = file.Save(fileName)
	if err != nil {
		return "", err
	}

	return fileName, nil
}

// SaveToExcel2 сохраняем данные в файл
// используем жестко заданные поля
func SaveToExcel2(s *model.ListPayments, path, templateFile string) (string, error) {

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	if templateFile == "" {
		templateFile = "file*.xlsx"
	}
	tmpfile, err := ioutil.TempFile(path, templateFile)
	if err != nil {
		return "", err
	}
	defer tmpfile.Close()
	fileName := tmpfile.Name()

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		return "", err
	}
	headerFont := xlsx.NewFont(12, "Calibri")
	headerFont.Bold = true
	headerFont.Underline = false
	headerStyle := xlsx.NewStyle()
	headerStyle.Font = *headerFont

	dataFont := xlsx.NewFont(11, "Calibri")
	dataStyle := xlsx.NewStyle()
	dataStyle.Font = *dataFont //*xlsx.DefaultFont()

	//Зададим наименование колонок
	row = sheet.AddRow()
	for index := 0; index < len(headerMap); index++ {
		cell = row.AddCell()
		cell.Value = headerMap[index].Name
		cell.SetStyle(headerStyle)
		if utf8.RuneCountInString(headerMap[index].Name) > headerMap[index].With {
			headerMap[index].With = utf8.RuneCountInString(headerMap[index].Name)
		}
	}

	//данные
	for index := 0; index < len(s.Db); index++ {
		row = sheet.AddRow()
		// добавляем поля в строке

		for colNo := 0; colNo < len(headerMap); colNo++ {
			cell = row.AddCell()
			cell.SetStyle(dataStyle)
			switch headerMap[colNo].Name {
			case "Occ":
				cell.SetInt(s.Db[index].Occ)
			case "Address":
				cell.Value = s.Db[index].Address
			case "Date":
				cell.SetDate(s.Db[index].Date)
			case "Value":
				cell.SetFloatWithFormat(s.Db[index].Value, "#,##0.00")
			case "Commission":
				cell.SetFloatWithFormat(s.Db[index].Commission, "#,##0.00")
			case "Fio":
				cell.Value = s.Db[index].Fio
			case "PaymentAccount":
				cell.Value = s.Db[index].PaymentAccount
			case "№":
				cell.SetInt(index + 1)
			}
			if utf8.RuneCountInString(cell.Value) > headerMap[colNo].With {
				headerMap[colNo].With = utf8.RuneCountInString(cell.Value)
			}
		}
	}
	//Устанавливаем ширину колонок
	for i, col := range sheet.Cols {
		col.Width = float64(headerMap[i].With)
	}

	err = file.Save(fileName)
	if err != nil {
		return "", err
	}

	return fileName, nil
}

// ReadFile читаем файл в структуру
func ReadFile(fileName string) (model.ListPayments, error) {
	res := model.ListPayments{}
	sheetName := "Sheet1"
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		return res, err
	}
	sheet := xlFile.Sheet[sheetName]
	for _, row := range sheet.Rows[1:] {
		p := model.Payment{}
		p.Occ, _ = row.Cells[headerName["Occ"].Position].Int()
		p.Address = row.Cells[headerName["Address"].Position].String()
		p.Date, _ = row.Cells[headerName["Date"].Position].GetTime(true)
		p.Value, _ = row.Cells[headerName["Value"].Position].Float()
		p.Commission, _ = row.Cells[headerName["Commission"].Position].Float()
		p.Fio = row.Cells[headerName["Fio"].Position].String()
		p.PaymentAccount = row.Cells[headerName["PaymentAccount"].Position].String()
		res.Db = append(res.Db, p)
	}
	return res, nil
}
