package excelize

import (
	"excel-sample/internal/domain/model"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
)

var testPayments *model.ListPayments

func TestMain(m *testing.M) {
	testPayments = model.PrepareTestData(model.CountTestData)
	os.Exit(m.Run())
}

func BenchmarkSaveToExcel1(b *testing.B) {
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		fileName, _ := SaveToExcel1(testPayments, ".", "file2*.xlsx")
		defer os.Remove(fileName)
	}
}

func TestSaveToExcel1(t *testing.T) {
	fileName, err := SaveToExcel1(testPayments, ".", "file2*.xlsx")
	defer os.Remove(fileName)
	assert.Empty(t, err)
	assert.FileExists(t, fileName)
}

func BenchmarkSaveToExcel2(b *testing.B) {
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		fileName, _ := SaveToExcel2(testPayments, ".", "file2*.xlsx")
		defer os.Remove(fileName)
	}
}

func TestSaveToExcel2(t *testing.T) {
	fileName, err := SaveToExcel2(testPayments, ".", "file2*.xlsx")
	defer os.Remove(fileName)
	assert.Empty(t, err)
	assert.FileExists(t, fileName)
}

func BenchmarkSaveToExcelStream(b *testing.B) {
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		fileName, _ := SaveToExcelStream(testPayments, ".", "fileStream*.xlsx")
		defer os.Remove(fileName)
	}
}

func TestSaveToExcelStream(t *testing.T) {
	fileName, err := SaveToExcelStream(testPayments, ".", "fileStream*.xlsx")
	assert.Empty(t, err)
	defer os.Remove(fileName)
	assert.FileExists(t, fileName)
}
