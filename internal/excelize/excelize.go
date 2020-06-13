package excelize

import (
	"io/ioutil"
	"reflect"
	"strconv"
	"strings"
	"time"

	"excel-sample/internal/domain/model"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"go.uber.org/zap"
)

// ListPayments структура для хранения платежей
type ListPayments model.Payments

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
		2: {Name: "Date", With: 10},
		3: {Name: "Участок", With: 10},
		4: {Name: "Occ", With: 10},
		5: {Name: "Address", With: 40},
		6: {Name: "Value", With: 10},
		7: {Name: "Код_услуги", With: 10},
		8: {Name: "Commission", With: 10},
		9: {Name: "PaymentAccount", With: 25},
	}
	headerName = make(map[string]*fieldExcel, 9)
)

func init() {
	for i := 0; i < len(headerMap); i++ {
		headerName[headerMap[i].Name] = &fieldExcel{Name: headerMap[i].Name, Position: i, With: headerMap[i].With}
	}
}

//SaveToExcel1 lib Excelize используем пакет reflect для определения полей структуры
func SaveToExcel1(s *model.ListPayments, path, templateFile string) (string, error) {
	if templateFile == "" {
		templateFile = "file*.xlsx"
	}
	tmpfile, err := ioutil.TempFile(path, templateFile)
	if err != nil {
		return "", err
	}
	defer tmpfile.Close()
	fileName := tmpfile.Name()
	zap.S().Infof("SaveToExcel2: %s", fileName)

	file := excelize.NewFile()

	sheetName := "Sheet1"
	indexSheet := file.NewSheet(sheetName)
	file.SetActiveSheet(indexSheet)

	expDate := "dd.MM.yyyy"
	styleDate, err := file.NewStyle(&excelize.Style{CustomNumFmt: &expDate})
	if err != nil {
		return "", err
	}
	styleHeader, err := file.NewStyle(`{"font":{"bold":true,"family":"Times New Roman","size":12}}`)
	if err != nil {
		return "", err
	}
	styleFloat, err := file.NewStyle(`{"number_format": 4}`)
	if err != nil {
		return "", err
	}
	//Зададим наименование колонок
	for index := 1; index <= len(headerMap); index++ {
		axis, err := excelize.CoordinatesToCellName(index, 1)
		if err != nil {
			return "", err
		}
		if err := file.SetCellValue(sheetName, axis, headerMap[index-1].Name); err != nil {
			return "", err
		}
		if err := file.SetCellStyle(sheetName, axis, axis, styleHeader); err != nil {
			return "", err
		}
		axis, _ = excelize.ColumnNumberToName(index)
		if err := file.SetColWidth(sheetName, axis, axis, float64(headerMap[index-1].With)); err != nil {
			return "", err
		}
	}

	//данные
	rowNo := 1
	for index := 0; index < len(s.Db); index++ {
		rowNo++
		// добавляем поля в строке
		values := reflect.ValueOf(s.Db[index])

		for i := 0; i < len(headerMap); i++ {
			axis, _ := excelize.CoordinatesToCellName(i+1, rowNo)
			f := values.FieldByName(strings.Title(headerMap[i].Name))
			if f.IsValid() {
				fieldValue := f.Interface()
				switch v := fieldValue.(type) {
				case float64:
					if err := file.SetCellFloat(sheetName, axis, v, 2, 64); err != nil {
						return "", err
					}
					if err := file.SetCellStyle(sheetName, axis, axis, styleFloat); err != nil {
						return "", err
					}
				case int:
					if err := file.SetCellInt(sheetName, axis, v); err != nil {
						return "", err
					}
				case string:
					if err := file.SetCellStr(sheetName, axis, v); err != nil {
						return "", err
					}
				case time.Time:
					if err := file.SetCellValue(sheetName, axis, v); err != nil {
						return "", err
					}
					if err := file.SetCellStyle(sheetName, axis, axis, styleDate); err != nil {
						return "", err
					}
				default:
					if err := file.SetCellValue(sheetName, axis, v); err != nil {
						return "", err
					}
				}
			}
		}
	}

	if err := file.SaveAs(fileName); err != nil {
		return "", err
	}
	return fileName, nil
}

//SaveToExcel2 lib Excelize используем жестко заданные поля
func SaveToExcel2(s *model.ListPayments, path, templateFile string) (string, error) {
	if templateFile == "" {
		templateFile = "file*.xlsx"
	}
	tmpfile, err := ioutil.TempFile(path, templateFile)
	if err != nil {
		return "", err
	}
	defer tmpfile.Close()
	fileName := tmpfile.Name()
	zap.S().Debugf("SaveToExcel2: %s", fileName)

	file := excelize.NewFile()

	sheetName := "Sheet1"
	indexSheet := file.NewSheet(sheetName)
	file.SetActiveSheet(indexSheet)

	expDate := "dd.MM.yyyy"
	styleDate, err := file.NewStyle(&excelize.Style{CustomNumFmt: &expDate})
	if err != nil {
		return "", err
	}
	styleHeader, err := file.NewStyle(`{"font":{"bold":true,"family":"Times New Roman","size":12}}`)
	if err != nil {
		return "", err
	}
	styleFloat, err := file.NewStyle(`{"number_format": 4}`)
	if err != nil {
		return "", err
	}
	CountCol := len(headerMap)
	//Зададим наименование колонок
	for index := 1; index <= CountCol; index++ {
		axis, _ := excelize.CoordinatesToCellName(index, 1)
		if err := file.SetCellValue(sheetName, axis, headerMap[index-1].Name); err != nil {
			return "", err
		}
		//err = file.SetCellStyle(sheetName, axis, axis, styleHeader)
	}
	axis, _ := excelize.CoordinatesToCellName(CountCol, 1)
	if err := file.SetCellStyle(sheetName, "A1", axis, styleHeader); err != nil {
		return "", err
	}
	//данные
	rowNo := 1
	for index := 0; index < len(s.Db); index++ {
		rowNo++
		// добавляем поля в строке

		for colNo := 1; colNo <= CountCol; colNo++ {
			switch headerMap[colNo-1].Name {
			case "Occ":
				axis, _ := excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellInt(sheetName, axis, s.Db[index].Occ); err != nil {
					return "", err
				}
			case "Address":
				axis, _ = excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellStr(sheetName, axis, s.Db[index].Address); err != nil {
					return "", err
				}
			case "Date":
				axis, _ = excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellValue(sheetName, axis, s.Db[index].Date); err != nil {
					return "", err
				}
				if err := file.SetCellStyle(sheetName, axis, axis, styleDate); err != nil {
					return "", err
				}
			case "Value":
				axis, _ = excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellFloat(sheetName, axis, s.Db[index].Value, 2, 64); err != nil {
					return "", err
				}
				if err := file.SetCellStyle(sheetName, axis, axis, styleFloat); err != nil {
					return "", err
				}
			case "Commission":
				axis, _ = excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellFloat(sheetName, axis, s.Db[index].Commission, 2, 64); err != nil {
					return "", err
				}
				if err := file.SetCellStyle(sheetName, axis, axis, styleFloat); err != nil {
					return "", err
				}
			case "Fio":
				axis, _ = excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellStr(sheetName, axis, s.Db[index].Fio); err != nil {
					return "", err
				}
			case "PaymentAccount":
				axis, _ = excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellStr(sheetName, axis, s.Db[index].PaymentAccount); err != nil {
					return "", err
				}
			case "№":
				axis, _ := excelize.CoordinatesToCellName(colNo, rowNo)
				if err := file.SetCellInt(sheetName, axis, index+1); err != nil {
					return "", err
				}
			}
		}
	}
	// устанавливаем ширину колонок
	for colNo := 1; colNo <= CountCol; colNo++ {
		startCol, _ := excelize.ColumnNumberToName(colNo)
		if err := file.SetColWidth(sheetName, startCol, startCol, float64(headerMap[colNo-1].With)); err != nil {
			return "", err
		}
	}

	if err := file.SaveAs(fileName); err != nil {
		return "", err
	}
	return fileName, nil
}

//SaveToExcelStream lib Excelize
func SaveToExcelStream(s *model.ListPayments, path, templateFile string) (string, error) {
	if templateFile == "" {
		templateFile = "file*.xlsx"
	}
	tmpfile, err := ioutil.TempFile(path, templateFile)
	if err != nil {
		return "", err
	}
	defer tmpfile.Close()
	fileName := tmpfile.Name()
	zap.S().Debugf("SaveToExcelStream: %s", fileName)
	file := excelize.NewFile()
	sheetName := "Sheet1"
	streamWriter, err := file.NewStreamWriter(sheetName)
	if err != nil {
		return "", err
	}

	expDate := "dd.MM.yyyy"
	styleDate, err := file.NewStyle(&excelize.Style{CustomNumFmt: &expDate})
	if err != nil {
		return "", err
	}
	expFloat := "#,##0.00"
	styleFloat, err := file.NewStyle(&excelize.Style{CustomNumFmt: &expFloat})
	if err != nil {
		return "", err
	}
	styleHeader, err := file.NewStyle(`{"font":{"bold":true,"family":"Times New Roman","size":12}}`)
	if err != nil {
		return "", err
	}

	CountCol := len(headerMap)
	rowHeader := make([]interface{}, CountCol)
	for i := 0; i < CountCol; i++ {
		rowHeader[i] = excelize.Cell{StyleID: styleHeader, Value: headerMap[i].Name}
	}
	if err := streamWriter.SetRow("A1", rowHeader); err != nil {
		return "", err
	}

	//данные
	rowNo := 1
	for index := 0; index < len(s.Db); index++ {
		row := make([]interface{}, CountCol)
		rowNo++

		row[0] = s.Db[index].Occ
		row[1] = s.Db[index].Address
		row[2] = excelize.Cell{StyleID: styleDate, Value: s.Db[index].Date}
		row[3] = excelize.Cell{StyleID: styleFloat, Value: s.Db[index].Value}
		row[4] = excelize.Cell{StyleID: styleFloat, Value: s.Db[index].Commission}
		row[5] = s.Db[index].Fio
		row[6] = s.Db[index].PaymentAccount

		cell, _ := excelize.CoordinatesToCellName(1, rowNo)
		if err := streamWriter.SetRow(cell, row); err != nil {
			return "", err
		}
	}

	if err := streamWriter.Flush(); err != nil {
		return "", err
	}

	// тормозит установка ширины колонок
	if err := file.SetColWidth(sheetName, "A", "G", 15); err != nil {
		return "", err
	}
	if err := file.SetColWidth(sheetName, "B", "B", 40); err != nil {
		return "", err
	}
	if err := file.SetColWidth(sheetName, "G", "G", 25); err != nil {
		return "", err
	}

	if err := file.SaveAs(fileName); err != nil {
		return "", err
	}
	return fileName, nil
}

// ReadFile читаем файл в структуру
func ReadFile(fileName string) (model.ListPayments, error) {
	res := model.ListPayments{}
	sheetName := "Sheet1"
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return res, err
	}
	// Получить все строки в Sheet1
	var p model.Payment
	rows, err := f.GetRows(sheetName)
	for _, row := range rows[1:] {
		p = model.Payment{}
		p.Occ, _ = strconv.Atoi(row[headerName["Occ"].Position])
		p.Address = row[headerName["Address"].Position]
		p.Date, _ = time.Parse("", row[headerName["Date"].Position])
		p.Value, _ = strconv.ParseFloat(row[headerName["Value"].Position], 64)
		p.Commission, _ = strconv.ParseFloat(row[headerName["Commission"].Position], 64)
		p.Fio = row[headerName["Fio"].Position]
		p.PaymentAccount = row[headerName["PaymentAccount"].Position]
		res.Db = append(res.Db, p)
	}
	return res, nil
}
