# excel 导入数据

## 项目简介
1. 提供读取和快速解析 excel

## 使用示例程序
```go
package main

import (
	"os"
	"fmt"

	ed "github.com/harddies/excel"
)

type BaseInfo struct {
	OaContractNumber  ed.IntField    `ex:"电子签合同信息|基础信息|OA合同编号"`
	OaProcessNumber   ed.IntField    `ex:"电子签合同信息|基础信息|OA流程编号"`
	Biz               ed.StringField `ex:"电子签合同信息|基础信息|签约业务"`
	Entity            ed.StringField `ex:"电子签合同信息|基础信息|签约主体"`
	SignType          ed.StringField `ex:"电子签合同信息|基础信息|签约类型"`
	ContractType      ed.StringField `ex:"电子签合同信息|基础信息|合同类型"`
	SignWay           ed.StringField `ex:"电子签合同信息|基础信息|签约形式"`
	ContractBeginTime ed.TimeField   `ex:"电子签合同信息|基础信息|合同开始时间"`
	ContractEndTime   ed.TimeField   `ex:"电子签合同信息|基础信息|合同开始时间"`
	ContractFile      ed.StringField `ex:"电子签合同信息|基础信息|合同文件"`
}

type SignEntity struct {
	PartAType     ed.StringField `ex:"电子签合同信息|签约方信息|甲方-类型"`
	PartAName     ed.StringField `ex:"电子签合同信息|签约方信息|甲方-名称"`
	PartPayRate   ed.FloatField  `ex:"电子签合同信息|签约方信息|甲方-出资比例"`
	PartBType     ed.StringField `ex:"电子签合同信息|签约方信息|乙方-类型"`
	PartBName     ed.StringField `ex:"电子签合同信息|签约方信息|乙方-名称"`
	PartBId       ed.IntField    `ex:"电子签合同信息|签约方信息|乙方-标识"`
	PartBUid      ed.IntField    `ex:"电子签合同信息|签约方信息|乙方-UID"`
	OtherPartType ed.StringField `ex:"电子签合同信息|签约方信息|其他方一-类型"`
	OtherPartName ed.StringField `ex:"电子签合同信息|签约方信息|其他方一-名称"`
	OtherPartId   ed.IntField    `ex:"电子签合同信息|签约方信息|其他方一-标识"`
	OtherPartUid  ed.IntField    `ex:"电子签合同信息|签约方信息|其他方一-UID"`
}

type OaInfo struct {
	ContractName ed.StringField `ex:"电子签合同信息|OA信息|合同全称"`
	ContractType ed.StringField `ex:"电子签合同信息|OA信息|合同类别"`
	PayOrEarn    ed.StringField `ex:"电子签合同信息|OA信息|收款/付款"`
	CashWay      ed.StringField `ex:"电子签合同信息|OA信息|收付款方式"`
	CashType     ed.StringField `ex:"电子签合同信息|OA信息|币种"`
	Amount       ed.IntField    `ex:"电子签合同信息|OA信息|金额"`
	ExceedLimit  ed.BoolField   `ex:"电子签合同信息|OA信息|是否超预算"`
	BizRisk      ed.BoolField   `ex:"电子签合同信息|OA信息|业务风险"`
}

func main() {
	//importRow()
	importSubRow()
	//asyncScanRows()
}

// import the whole of a row
func importRow() {
	dir, _ := os.Getwd()
	f, _ := ed.NewExcelFromFile(dir + "/excel/example/demo.xlsx", ed.ActiveSheet("Sheet1"))

	rows, _ := f.GetRowsWithoutHeader()

	baseInfo, entityInfo, oaInfo := new(BaseInfo), new(SignEntity), new(OaInfo)

	if err := f.ScanExRow(rows[0], &baseInfo, &entityInfo, &oaInfo); err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println(baseInfo, entityInfo, oaInfo)
	fmt.Println(baseInfo.OaContractNumber.GetStdValue(), baseInfo.ContractType.GetStdValue())
}

// import the sub of a row by subImporter path
func importSubRow() {
	dir, _ := os.Getwd()
	f, _ := ed.NewExcelFromFile(dir + "/excel/example/demo.xlsx", ed.ActiveSheet("Sheet1"))

	rows, _ := f.GetRowsWithoutHeader()

	baseInfo := new(BaseInfo)
	baseInfoImporter := f.SubImporter("电子签合同信息|基础信息")
	if baseInfoImporter != nil {
		nodeColStartIdx, nodeColEndIdx := baseInfoImporter.GetColIndexPos()
		if err := f.ScanExRow(rows[0][nodeColStartIdx-1:nodeColEndIdx], baseInfo); err != nil {
			fmt.Println(err)
			return
		}
	}

	fmt.Println(baseInfo)
	fmt.Println(baseInfo.OaContractNumber.GetStdValue())
}

// async import rows
func asyncScanRows() {
	dir, _ := os.Getwd()
	f, _ := ed.NewExcelFromFile(dir + "/excel/example/demo.xlsx", ed.ActiveSheet("Sheet1"))

	rows, _ := f.GetRowsWithoutHeader()
	baseInfo, entityInfo, oaInfo := new(BaseInfo), new(SignEntity), new(OaInfo)
	for row := range f.AsyncScanExRows(rows, &baseInfo, &entityInfo, &oaInfo) {
		if row.Err != nil {
			fmt.Println(row.Err)
		} else {
			fmt.Println(row.Resps[0].(*BaseInfo), row.Resps[1].(*SignEntity), row.Resps[2].(*OaInfo))
		}
	}
}

```

## excel 导入模版要求
- 表头必须都是合并单元格。对应的正文则不能是合并单元格，只能调整对应单元格的列宽列高来适应内容