package excel

import (
	"fmt"
	"reflect"
	"time"

	//_ "reflect"
	"strconv"

	"git.gdqlyt.com.cn/go/base/beego/bmodel"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
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

type Portal struct {
	nameMap    NameMap
	dateMapper DateMapper
}

func NewPortal(nameMap NameMap) *Portal {
	p := new(Portal)
	p.nameMap = nameMap
	p.dateMapper = make(DateMapper)
	p.dateMapper[DefStyle] = 0
	return p
}

func (p *Portal) SetNameMap(_nameMap NameMap) *Portal {
	p.nameMap = _nameMap
	return p
}

func (p *Portal) SetDateMapper(_datemapper DateMapper) *Portal {
	p.dateMapper = _datemapper
	p.dateMapper[DefStyle] = 0
	return p
}

// 导出 excel
func (p *Portal) BuildExcel(_models interface{}, _sheet ...string) (*excelize.File, error) {
	var sheet = DefSheet
	if len(_sheet) > 0 {
		sheet = _sheet[0]
	}

	var (
		col    = new(int)
		file   = excelize.NewFile()
		loop   = p.makeSetAxis(file, sheet)
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

	return file, file.SetColWidth(sheet, StartCol, computeColumn(*col), 20)
}

func (p *Portal) makeSetAxis(f *excelize.File, sheet string) setAxis {
	return func(v reflect.Value, col *int, row string, pn bool, fn setAxis) {
		var (
			ftype reflect.StructField
			field reflect.Value
		)

		for idx := 0; idx < v.NumField(); idx++ {
			ftype = v.Type().Field(idx)
			field = v.Field(idx)
			// 递归结构体
			if field.Kind() == reflect.Struct &&
				ftype.Type.Name() != "LocalTime" &&
				ftype.Type.Name() != "DateTime" {
				fn(field, col, row, pn, fn)
				continue
			}

			// 写入单元格
			name, exist := p.nameMap[ftype.Name]
			if !exist {
				continue
			}
			strcol := computeColumn(*col)
			if pn {
				_ = f.SetCellValue(sheet, strcol+StartRow, name)
			}

			var cell interface{}
			switch field.Kind() {
			case reflect.Struct:
				switch field.Type().Name() {
				case "LocalTime":
					cell = field.Interface().(bmodel.LocalTime).GetTime().UTC()
					p.formatDateTime(f, sheet, strcol+row, ftype.Name)
				case "DateTime":
					cell = field.Interface().(bmodel.DateTime).GetTime().UTC()
					p.formatDateTime(f, sheet, strcol+row, ftype.Name)
				}
			default:
				cell = field.Interface()
			}
			err := f.SetCellValue(sheet, strcol+row, cell)
			if err != nil {
				fmt.Println(err, cell)
				_ = f.SetCellValue(sheet, strcol+row, err.Error())
			}
			*col++
		}
	}
}

// 导入 excel
func (p *Portal) LoadExcel(file *excelize.File, models interface{}, _sheet ...string) error {
	var sheet = file.GetSheetName(0)
	if len(_sheet) > 0 {
		sheet = _sheet[0]
	}

	var (
		col    = new(int)
		fmap   = make(NameMap)
		loop   = p.makeGetAxis(file, fmap, sheet)
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

	for k, v := range p.nameMap {
		fmap[v] = k
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

func (p *Portal) makeGetAxis(f *excelize.File, fmap NameMap, sheet string) getAxis {
	fmaplen := len(p.nameMap)
	return func(v reflect.Value, col *int, row string, fn getAxis) bool {
		for ; ; *col++ {
			if *col >= fmaplen {
				return true
			}

			strCol := computeColumn(*col)
			name, _ := f.GetCellValue(sheet, strCol+StartRow)
			if name == "" {
				return *col == 1
			}

			fname, exist := fmap[name]
			if !exist {
				continue
			}

			// 读取单元格
			p.formatDateTime(f, sheet, strCol+row, DefStyle)
			cell, _ := f.GetCellValue(sheet, strCol+row)
			if cell == "" && *col == 1 {
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
					dt := timeFromExcelTime(cell)
					fieldV.Set(reflect.ValueOf(bmodel.LocalTime(dt)))
				case "DateTime":
					dt := timeFromExcelTime(cell)
					fieldV.Set(reflect.ValueOf(bmodel.DateTime(dt)))
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
func computeColumn(column int) string {
	if column == 0 {
		return ""
	}
	diff := column / 26
	if column%26 == 0 {
		diff--
	}
	return computeColumn(diff) + string(rune(column-1)%26+'A')
}

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

func (p *Portal) formatDateTime(f *excelize.File, sheet, axis, fname string) {
	styleid, exist := p.dateMapper[fname]
	if !exist {
		styleid = 22
	}

	style, _ := f.NewStyle(fmt.Sprintf(`{"number_format": %d, "lang": "zh-cn"}`, styleid))
	_ = f.SetCellStyle(sheet, axis, axis, style)
}
