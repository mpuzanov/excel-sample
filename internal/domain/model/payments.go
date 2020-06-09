package model

import (
	"math/rand"
	"time"

	"golang.org/x/text/language"
	"golang.org/x/text/message"
)

// Payment Банковский платеж
type Payment struct {
	Occ            int       `json:"occ" db:"occ" xml:"occ"`
	Address        string    `json:"address,omitempty" db:"address" xml:"address,omitempty"`
	Date           time.Time `json:"date" db:"date" xml:"date"`
	Value          float64   `json:"value" db:"value" xml:"value"`
	Commission     float64   `json:"commission" db:"commission" xml:"commission"`
	Fio            string    `json:"fio,omitempty" db:"fio" xml:"fio,omitempty"`
	PaymentAccount string    `json:"payment_account,omitempty" db:"payment_account" xml:"payment_account,omitempty"`
}

// Payments структура для хранения платежей
type Payments struct {
	Db []Payment `json:"payments" xml:"payment"`
}

// ListPayments .
type ListPayments Payments

// StoragePayments .
type StoragePayments interface {
	SaveToExcel1(path, templateFile string) (string, error)
	SaveToExcel2(path, templateFile string) (string, error)
}

// CountTestData кол-во платежей для тестирования
var CountTestData = 10000

// String Вывод общей информации о слайсе платежей
func (p *ListPayments) String() string {
	val := 0.0
	com := 0.0
	for i := 0; i < len(p.Db); i++ {
		val += p.Db[i].Value
		com += p.Db[i].Commission
	}
	pr := message.NewPrinter(language.Russian)
	return pr.Sprintf("Кол-во платежей: %d на сумму: %.2f руб. с комиссией: %.2f руб.\n", len(p.Db), val, com)
	//return fmt.Sprintf("Кол-во платежей: %d на сумму: %.2f с комиссией: %.2f\n", len(p.Db), val, com)
}

// PrepareTestData создаём тестовый слайс платежей
func PrepareTestData(count int) *ListPayments {
	tp := ListPayments{}
	tp.Db = make([]Payment, count)
	for i := 0; i < count; i++ {
		tp.Db[i].Occ = rand.Intn(999999)
		tp.Db[i].Address = "Пушкинская, 240А, 50"
		tp.Db[i].Date = time.Date(2018, time.August, 30, 0, 0, 0, 0, time.UTC) //time.Now()
		tp.Db[i].Value = rand.Float64() + float64(rand.Intn(9999))
		tp.Db[i].Commission = rand.Float64() + float64(rand.Intn(99))
		tp.Db[i].Fio = "Иванов Иван Иванович"
		tp.Db[i].PaymentAccount = "12345678901234567890"
	}
	return &tp
}
