package excel

import (
	"fmt"
	"reflect"
	"strings"
	"testing"

	"git.gdqlyt.com.cn/go/base/beego/bmodel"
	"github.com/xuri/excelize/v2"
)

const (
	ExcelFile     = "out/test.xlsx"
	LoadTestFile  = "out/load.xlsx"
	BuildTestFile = "out/build.xlsx"
)

type Data struct {
	Ok string
}

type BaseModel struct {
	Data
	Id           string //内部编码
	Displayno    int    //显示序号
	Enablestatus string //是否停用 0:启用 1:停用
	Deleteflag   string //是否删除
	Hospitalcode string //机构编码
	Dataversion  int    //记录更新版本号，每次insert,update自动+1
}

type Test struct {
	BaseModel
	Patcode        string
	Patientname    string
	Createdate     bmodel.LocalTime
	Homedetailaddr string
	Socialnum      string
	Teamname       string
	Treatstartdate bmodel.LocalTime
	Treatenddate   bmodel.LocalTime
	Admissiondate  bmodel.LocalTime
	F1             float32
	F2             float64
}

func (t *Test) TableName() string {
	return "test"
}

var TestNameMap = NameMap{
	"Ok":             "整活",
	"Id":             "唯一索引ID",
	"Displayno":      "显示序号",
	"Enablestatus":   "是否停用",
	"Deleteflag":     "是否删除",
	"Hospitalcode":   "机构编码",
	"Dataversion":    "版本号",
	"Patcode":        "病人编号",
	"Createdate":     "创建时间",
	"Patientname":    "姓名",
	"Homedetailaddr": "户籍地（省、市、县）",
	"Socialnum":      "身份证号",
	"Teamname":       "审批机关",
	"Treatstartdate": "戒毒开始日期",
	"Treatenddate":   "戒毒结束日期",
	"Admissiondate":  "入所日期",
	"F1":             "F1",
	"F2":             "F2",
}

// DateFormatter 的配置请参考 README.md
var TestDateFormatter = DateMapper{
	//"Createdate": 27,
	"Stdt": 31,
}

func TestBuildExcel(t *testing.T) {
	// 测试 BuildExcel
	var list []Test
	list = append(list,
		Test{
			BaseModel: BaseModel{
				Data: Data{
					Ok: "ojbk",
				},
				Id:           "2153235",
				Displayno:    2,
				Enablestatus: "1",
				Deleteflag:   "1",
				Hospitalcode: "29",
				Dataversion:  2,
			},
			Patcode: "123123",
			F1:      3.1111,
			F2:      3.1111,
		},
		Test{
			BaseModel: BaseModel{
				Data: Data{
					Ok: "okkkkkkkkkkk",
				},
				Id:           "92929",
				Displayno:    454,
				Enablestatus: "0",
				Deleteflag:   "0",
				Hospitalcode: "91",
				Dataversion:  22,
			},
			Patcode:    "722",
			Createdate: bmodel.NewNowLocalTime(),
		})

	portal := NewPortal(TestNameMap).SetDateMapper(TestDateFormatter)
	file, err := portal.BuildExcel(list)
	if err != nil {
		t.Fatal(err)
	}

	// 测试 AppendGrid
	var grid = portal.FormatGrid(
		"||药品名称|规格|出库数量|-|-|-|厂家|价格|批号",
		"||^|^|大单位数量|药库单位|小单位数量|药房单位|^|^|^",
	)
	portal.AppendGrid(file, 5, grid)

	// 测试 AppendRow
	var rowSpan []string
	rowSpan = strings.Split("|合计|dsfa|1|das|123|3.000|63.30|asd|wqe|vxc", "|")
	portal.AppendRow(file, 7, rowSpan)
	_ = file.SaveAs(ExcelFile)
}

func TestLoadExcel(t *testing.T) {
	var list []Test
	file, err := excelize.OpenFile(ExcelFile)
	if err != nil {
		t.Fatal("读取 excel 文件失败，请检查文件是否存在")
	}

	portal := NewPortal(TestNameMap)
	err = portal.LoadExcel(file, &list)
	if err != nil {
		t.Fatal(err)
	}

	for _, m := range list {
		fmt.Println(m)
	}
}

func TestNewval(t *testing.T) {
	var test ******Test
	vref := reflect.TypeOf(test)
	value := makeval(vref)
	fmt.Println(
		value.Type(),
		value.Elem().Type(),
		value.Elem().Elem().Type(),
		value.Elem().Elem().Elem().Type(),
		value.Elem().Elem().Elem().Elem().Type(),
		value.Elem().Elem().Elem().Elem().Elem().Type(),
		value.Elem().Elem().Elem().Elem().Elem().Elem().Type(),
	)
}

func TestStructTag(t *testing.T) {
	//var data Test
	//v:= reflect.ValueOf(data)
	//
	//reflect.StructTag("").Lookup()
}
