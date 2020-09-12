package common

import (
	"fmt"
	"git.gdqlyt.com.cn/go/base/beego/bmodel"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"testing"
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
	Patcode    string //eg: 强戒号
	Createdate bmodel.LocalTime
}

func (t *Test) TableName() string {
	return "test"
}

var NameMap = map[string]string{
	"Ok":           "整活",
	"Id":           "唯一索引ID",
	"Displayno":    "显示序号",
	"Enablestatus": "是否停用",
	"Deleteflag":   "是否删除",
	"Hospitalcode": "机构编码",
	"Dataversion":  "版本号",
	"Patcode":      "病人编号",
	"Createdate":   "创建时间",
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
			Patcode:    "123123",
			Createdate: bmodel.NewNowLocalTime(),
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
			Patcode: "722",
			Createdate: bmodel.NewNowLocalTime(),
		})

	excelFile, err := BuildExcel(list, NameMap)
	if err != nil {
		t.Fatal(err)
	}
	_ = excelFile.SaveAs("test.xlsx")
}

func TestLoadExcel(t *testing.T) {
	var list []Test
	file, _ := excelize.OpenFile("test.xlsx")
	err := LoadExcel(file, NameMap, &list)
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
