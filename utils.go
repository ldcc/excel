package excel

import (
	"reflect"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	DefSheet = "Sheet1"
	StartCol = "A"
	StartRow = "1"
	DefStyle = "general_style"
)

type (
	NameMap    map[string]string
	DateMapper map[string]int
	setAxis    func(v reflect.Value, col *int, row string, pn bool, fn setAxis)
	getAxis    func(v reflect.Value, col *int, row string, fn getAxis) (br bool)
)

func makeval(t reflect.Type) reflect.Value {
	switch t.Kind() {
	case reflect.Ptr:
		v0 := reflect.New(t).Elem()
		v1 := makeval(t.Elem())
		v0.Set(reflect.New(v1.Type()))
		v0.Elem().Set(v1)
		return v0
	default:
		return reflect.New(t).Elem()
	}
}

func indirect(v reflect.Value) reflect.Value {
	if v.Kind() != reflect.Ptr {
		return v
	}
	return indirect(v.Elem())
}

// 根据列数计算相应的 Excel 列名
func evalColumn(column int) string {
	if column == 0 {
		return ""
	}
	diff := column / 26
	if column%26 == 0 {
		diff--
	}
	return evalColumn(diff) + string(rune(column-1)%26+'A')
}

// 从 Excel 时间获取 Unix 时间戳
func timeFromExcelTime(cell string) time.Time {
	unix, err := strconv.ParseFloat(cell, 64)
	if err != nil {
		return time.Time{}
	}

	dt, err := excelize.ExcelDateToTime(unix, false)
	if err != nil {
		return time.Time{}
	}

	return dt
}
