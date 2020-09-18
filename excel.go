package excel

import (
	"fmt"
	"reflect"
	//_ "reflect"
	"strconv"

	"git.gdqlyt.com.cn/go/base/beego/bmodel"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

const (
	DefSheet = "Sheet1"
	StartCol = "A"
	StartRow = "1"
)

type (
	setAxis func(v reflect.Value, col *int, row string, pn bool, fn setAxis)
	getAxis func(v reflect.Value, col *int, row string, fn getAxis) (br bool)
)

// 导出 excel
func BuildExcel(_models interface{}, nMap map[string]string, _sheet ...string) (*excelize.File, error) {
	var sheet = DefSheet
	if len(_sheet) > 0 {
		sheet = _sheet[0]
	}

	var (
		col    = new(int)
		file   = excelize.NewFile()
		loop   = makeSetAxis(file, nMap, sheet)
		sliceV = reflect.Indirect(reflect.ValueOf(_models))
	)
	if sliceV.Kind() != reflect.Slice {
		return nil, fmt.Errorf("只能接收 Slice 类型数据")
	}

	sliceLen := sliceV.Len()
	for row := 0; row < sliceLen; row++ {
		refV := reflect.Indirect(sliceV.Index(row))
		if refV.Interface() == nil {
			continue
		}

		*col = 1
		// row 初始为 0，为字段名空 1 行，所以 +2
		loop(refV, col, strconv.Itoa(row+2), row <= 0, loop)
	}
	*col--

	return file, file.SetColWidth(sheet, StartCol, ComputeColumn(*col), 30)
}

func makeSetAxis(f *excelize.File, nameMap map[string]string, sheet string) setAxis {
	return func(v reflect.Value, col *int, row string, pn bool, fn setAxis) {
		var (
			ftype reflect.StructField
			field reflect.Value
		)

		for idx := 0; idx < v.NumField(); idx++ {
			ftype = v.Type().Field(idx)
			field = v.Field(idx)
			// 递归结构体
			if ftype.Type.Kind() == reflect.Struct &&
				ftype.Type.Name() != "LocalTime" &&
				ftype.Type.Name() != "DateTime" {
				fn(field, col, row, pn, fn)
				continue
			}

			// 写入单元格
			name, exist := nameMap[ftype.Name]
			if !exist {
				continue
			}
			strcol := ComputeColumn(*col)
			if pn {
				_ = f.SetCellValue(sheet, strcol+StartRow, name)
			}
			_ = f.SetCellValue(sheet, strcol+row, field.Interface())
			*col++
		}
	}
}

// 导入 excel
func LoadExcel(file *excelize.File, _nMap map[string]string, models interface{}, _sheet ...string) error {
	var sheet = file.GetSheetName(0)
	if len(_sheet) > 0 {
		sheet = _sheet[0]
	}

	var (
		col    = new(int)
		fMap   = make(map[string]string, len(_nMap))
		loop   = makeGetAxis(file, fMap, sheet)
		ptrV   = reflect.ValueOf(models)
		sliceV reflect.Value
		valueT reflect.Type
	)

	if ptrV.Kind() != reflect.Ptr {
		return fmt.Errorf("只能接收 Pointer 类型数据")
	}

	sliceV = ptrV.Elem()
	if sliceV.Kind() != reflect.Slice {
		return fmt.Errorf("只能接收指向 Slice 类型的 Pointer")
	}

	for k, v := range _nMap {
		fMap[v] = k
	}

	valueT = sliceV.Type().Elem()
	for row := 0; ; row++ {
		sref := makeval(valueT)
		vref := indirect(sref)

		*col = 1
		// row 初始为 0，为字段名空 1 行，所以 +2
		br := loop(vref, col, strconv.Itoa(row+2), loop)
		if br {
			break
		}

		sliceV.Set(reflect.Append(sliceV, sref))
	}

	return nil
}

func makeGetAxis(f *excelize.File, fieldMap map[string]string, sheet string) getAxis {
	return func(v reflect.Value, col *int, row string, fn getAxis) bool {
		for ; ; *col++ {
			strCol := ComputeColumn(*col)
			name, _ := f.GetCellValue(sheet, strCol+StartRow)
			if name == "" {
				if *col == 1 {
					return true
				}
				return false
			}

			fname, exist := fieldMap[name]
			if !exist {
				continue
			}

			// 读取单元格
			cell, _ := f.GetCellValue(sheet, strCol+row)
			if cell == "" {
				return true
			}

			fieldV := v.FieldByName(fname)
			switch fieldV.Kind() {
			case reflect.String:
				fieldV.SetString(cell)
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64, reflect.Float32, reflect.Float64:
				atoi, err := strconv.Atoi(cell)
				if err != nil {
					continue
				}
				fieldV.SetInt(int64(atoi))
			case reflect.Struct:
				switch fieldV.Type().Name() {
				case "BaseModel":
					fn(fieldV, col, row, fn)
				case "LocalTime":
					fieldV.Set(reflect.ValueOf(bmodel.NewLocalTime(cell)))
				case "DateTime":
					fieldV.Set(reflect.ValueOf(bmodel.NewDateTime(cell)))
				}
			}
		}
	}
}

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
func ComputeColumn(column int) string {
	if column == 0 {
		return ""
	}
	diff := column / 26
	if column%26 == 0 {
		diff--
	}
	return ComputeColumn(diff) + string(rune(column-1)%26+'A')
}
