package excel

import (
	"fmt"
	"git.gdqlyt.com.cn/go/base/beego/bmodel"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"testing"
)

const (
	ExcelFile     = "test.xlsx"
	LoadTestFile  = "load.xlsx"
	BuildTestFile = "build.xlsx"
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
}

var TestDateFormatter = DateMapper{
	"Createdate": 27,
	"Stdt":       31,
}

func TestBuildExcel(t *testing.T) {
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
			//Createdate: bmodel.NewNowLocalTime(),
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
	excelFile, err := portal.BuildExcel(list[:0])
	if err != nil {
		t.Fatal(err)
	}
	_ = excelFile.SaveAs(ExcelFile)
}

func TestLoadExcel(t *testing.T) {
	var list []Test
	file, err := excelize.OpenFile(LoadTestFile)
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
