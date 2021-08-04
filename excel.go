package excel

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	"git.gdqlyt.com.cn/go/base/beego/bmodel"
	"github.com/siddontang/go/num"
	"github.com/xuri/excelize/v2"
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
		runing int
		done   = make(chan int)
		file   = excelize.NewFile()
		loop   = p.makeSetAxis(file, sheet)
		sliceV = indirect(reflect.ValueOf(_models))
	)
	if sliceV.Kind() != reflect.Slice {
		return nil, fmt.Errorf("只能接收 Slice 类型数据")
	}

	sliceLen := sliceV.Len()
	if sliceLen == 0 {
		return file, nil
	}
	for row := 0; row < sliceLen; row++ {
		refV := indirect(sliceV.Index(row))
		if refV.Interface() == nil {
			continue
		}

		runing++
		go func(row int) {
			col := 1
			// row 初始为 0，为字段名空 1 行，所以 +2
			loop(refV, &col, strconv.Itoa(row+2), row <= 0, loop)
			done <- col
		}(row)
	}

	maxcol := 0
	for i := 0; i < runing; i++ {
		select {
		case col := <-done:
			maxcol = num.MaxInt(maxcol, col)
		}
	}

	return file, file.SetColWidth(sheet, StartCol, evalColumn(maxcol), 20)
}

func (p *Portal) makeSetAxis(f *excelize.File, sheet string) setAxis {
	var t001 = time.Time{}.UTC()
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

			// 写入列名
			name, exist := p.nameMap[ftype.Name]
			if !exist {
				continue
			}
			strcol := evalColumn(*col)
			if pn {
				_ = f.SetCellValue(sheet, strcol+StartRow, name)
			}

			// 写入单元格
			var cell interface{}
			switch field.Kind() {
			case reflect.Struct:
				switch field.Type().Name() {
				case "Time":
					cell = field.Interface().(time.Time).UTC()
				case "LocalTime":
					cell = field.Interface().(bmodel.LocalTime).GetTime().UTC()
				case "DateTime":
					cell = field.Interface().(bmodel.DateTime).GetTime().UTC()
				}
				if cell == t001 {
					*col++
					continue
				}
				cell = cell.(time.Time).Add(8 * time.Hour)
				p.formatDateTime(f, sheet, strcol+row, ftype.Name)
			default:
				cell = field.Interface()
			}

			err := f.SetCellValue(sheet, strcol+row, cell)
			if err != nil {
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
		loop   = p.makeGetAxis(file, sheet)
		ptrV   = reflect.ValueOf(models)
		sliceV reflect.Value
		valueT reflect.Type
	)

	if ptrV.Kind() != reflect.Ptr {
		return fmt.Errorf("只能接收 Pointer 类型数据")
	}

	sliceV = indirect(ptrV)
	if sliceV.Kind() != reflect.Slice {
		return fmt.Errorf("只能接收指向 Slice 类型的 Pointer")
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

func (p *Portal) makeGetAxis(f *excelize.File, sheet string) getAxis {
	maplen := len(p.nameMap)
	fmap := make(NameMap, maplen)
	for k, v := range p.nameMap {
		fmap[v] = k
	}
	return func(v reflect.Value, col *int, row string, fn getAxis) bool {
		emptycol := 1
		for ; ; *col++ {
			strcol := evalColumn(*col)
			name, _ := f.GetCellValue(sheet, strcol+StartRow)
			if name == "" {
				return emptycol == *col
			}

			// 读取单元格
			p.formatDateTime(f, sheet, strcol+row, DefStyle)
			cell, _ := f.GetCellValue(sheet, strcol+row)
			if cell == "" {
				emptycol++
				continue
			}

			fname, exist := fmap[name]
			if !exist {
				continue
			}

			fieldV := v.FieldByName(fname)
			switch fieldV.Kind() {
			case reflect.String:
				fieldV.SetString(cell)
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				atoi, err := strconv.Atoi(cell)
				if err != nil {
					continue
				}
				fieldV.SetInt(int64(atoi))
			case reflect.Float32, reflect.Float64:
				atoi, err := strconv.ParseFloat(cell, 64)
				if err != nil {
					continue
				}
				fieldV.SetFloat(atoi)
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

func (p *Portal) formatDateTime(f *excelize.File, sheet, axis, fname string) {
	styleid, exist := p.dateMapper[fname]
	if !exist {
		styleid = 22
	}

	style, _ := f.NewStyle(fmt.Sprintf(`{"number_format": %d, "lang": "zh-cn"}`, styleid))
	_ = f.SetCellStyle(sheet, axis, axis, style)
}

func (p *Portal) formatCenter(f *excelize.File, sheet, axis string) {
	style, _ := f.NewStyle(`{"alignment":{"horizontal":"center","vertical":"center"}}`)
	_ = f.SetCellStyle(sheet, axis, axis, style)
}

/**
 * 设置指定行一行的的数据
 */
func (p *Portal) AppendRow(file *excelize.File, _rowIndex int, rowSpan []string, _sheet ...string) *Portal {
	var sheet = file.GetSheetName(0)
	if len(_sheet) > 0 {
		sheet = _sheet[0]
	}

	row := strconv.Itoa(_rowIndex)
	for col, cell := range rowSpan {
		strcol := evalColumn(col + 1)
		err := file.SetCellStr(sheet, strcol+row, cell)
		if err != nil {
			_ = file.SetCellValue(sheet, strcol+row, err.Error())
		}
	}
	return p
}

/**
 * Grid :: [[String]]
 * Grid[i] 为列
 * "-" 为向左合并项
 * "^" 为向上合并项空项
 */
func (p *Portal) AppendGrid(file *excelize.File, _rowIndex int, grid [][]string, _sheet ...string) *Portal {
	var sheet = file.GetSheetName(0)
	if len(_sheet) > 0 {
		sheet = _sheet[0]
	}

	var err error
	for offset, rowSpan := range grid {
		row := strconv.Itoa(_rowIndex + offset)
		for col, cell := range rowSpan {
			prevcol := evalColumn(col)
			strcol := evalColumn(col + 1)
			switch cell {
			case "-":
				err = file.MergeCell(sheet, prevcol+row, strcol+row)
			case "^":
				prevrow := strconv.Itoa(_rowIndex + offset - 1)
				err = file.MergeCell(sheet, strcol+prevrow, strcol+row)
			default:
				err = file.SetCellStr(sheet, strcol+row, cell)
			}
			if err != nil {
				_ = file.SetCellValue(sheet, strcol+row, err.Error())
			}
			p.formatCenter(file, sheet, strcol+row)
		}
	}

	return p
}

func (p *Portal) FormatGrid(formats ...string) (grid [][]string) {
	for _, row := range formats {
		grid = append(grid, strings.Split(row, "|"))
	}
	return grid
}
