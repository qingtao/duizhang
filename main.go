package main

import (
	"encoding/json"
	"fmt"
	"image"
	"image/draw"
	"image/png"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/Luxurioust/excelize"
	"github.com/boombuler/barcode"
	"github.com/boombuler/barcode/code39"
	"github.com/golang/freetype"
)

//银行回执联和客户留存联对应excel单元格位置
type Statement struct {
	//客户账号
	Account string
	//账号类型
	AccountType string
	//币种
	Currency string
	//开户行
	DepositBank string
	//账户余额
	AccountBalance string
	//已质押余额
	Pledgedbalance string
	//原始序号
	ID string `json:",omitempty"`
	//银行代码
	BankCode string `json:",omitempty"`
}

type BankRecipts struct {
	MaxLine   int
	RowStart  int
	Statement Statement
}

type ClientCopy struct {
	ClientCode  string
	CompanyName string
	DueDate     string
	MaxLine     int
	RowStart    int
	Statement   Statement
}

//华夏银行对账单
type HuaxiaBankStatement struct {
	//邮政编码
	PostalCode string
	//客户号
	ClientCode string
	//单位名称
	CompanyName string
	//开户行
	DepositBank string
	//客户地址
	ClientAddress string
	//联系人
	Contact string
	//电话号码
	PhoneNumber string
	//余额截至日期
	DueDate string
	//条形码
	BarCode string
	//银行回执联
	BankRecipts BankRecipts
	//客户留存联
	ClientCopy ClientCopy
}

//存放对账单数据
type HuaxiaBank struct {
	//条码
	BarCode string
	//邮政编码
	PostalCode string
	//客户号
	ClientCode string
	//单位名称
	CompanyName string
	//开户行
	DepositBank string
	//客户地址
	ClientAddress string
	//联系人
	Contact string
	//电话号码
	PhoneNumber string
	//余额截至日期
	DueDate string
	//明细列表
	BankRecipts []*Statement
	//客户留存联列表
	//ClientCopy []*Statement
}

//生成条形码，并添加数字
func GenBarCode(x, y int, fontsize float64, src image.Image, s, dir, fontfile string) error {
	//创建空白背景图片，大小宽x，高y
	r := image.Rect(0, 0, x, y)
	dst := image.NewRGBA(r)
	//填充白色背景
	draw.Draw(dst, dst.Bounds(), image.White, image.ZP, draw.Over)
	//保存文件名是条码的数值+".png", 这个目录位置需要需改
	file := filepath.Join(dir, s+".png")
	dstf, err := os.Create(file)
	if err != nil {
		return err
	}
	defer dstf.Close()

	//绘制条码
	draw.Draw(dst, src.Bounds(), src, image.Pt(0, -10), draw.Over)

	//读取字体
	fb, err := ioutil.ReadFile(fontfile)
	if err != nil {
		return err
	}
	ft, err := freetype.ParseFont(fb)
	if err != nil {
		return err
	}
	//fg, bg := image.Black, image.White
	//rule := color.RGBA{0xdd, 0xdd, 0xdd, 0xff}

	//设置文本的格式
	c := freetype.NewContext()
	c.SetDPI(72)
	c.SetFont(ft)
	c.SetFontSize(fontsize)
	c.SetClip(dst.Bounds())
	c.SetDst(dst)
	c.SetSrc(image.Black)
	//文本开始位置
	pt := freetype.Pt(100, 25+int(c.PointToFixed(12)>>6))
	//绘制条码数字
	s = strings.Join(strings.Split(s, ""), " ")
	s = "* " + s + " *"
	_, err = c.DrawString(s, pt)
	if err != nil {
		return err
	}

	//保存png图片
	err = png.Encode(dstf, dst)
	if err != nil {
		return err
	}
	return nil
}

//打开excel模版文件
func OpenExcel(file string) (*excelize.File, error) {
	f, err := excelize.OpenFile(file)
	return f, err
}

//设置对账单列表部分
func SetStatement(xlsx *excelize.File, sheet string, rowStart, maxLine int, fst Statement, dst []*Statement) {
	for i := 0; i < maxLine; i++ {
		row := rowStart + i
		account := fmt.Sprintf("%s%d", fst.Account, row)
		if i < len(dst) {
			xlsx.SetCellStr(sheet, account, dst[i].Account)
			//如果数据行数小于maxline，清空多余的内容
		} else {
			xlsx.SetCellStr(sheet, account, "")
		}

		accountType := fmt.Sprintf("%s%d", fst.AccountType, row)
		if i < len(dst) {
			xlsx.SetCellStr(sheet, accountType, dst[i].AccountType)
		} else {
			xlsx.SetCellStr(sheet, accountType, "")
		}

		currency := fmt.Sprintf("%s%d", fst.Currency, row)
		if i < len(dst) {
			xlsx.SetCellStr(sheet, currency, dst[i].Currency)
		} else {
			xlsx.SetCellStr(sheet, currency, "")
		}

		depositBank := fmt.Sprintf("%s%d", fst.DepositBank, row)
		if i < len(dst) {
			xlsx.SetCellStr(sheet, depositBank, dst[i].DepositBank)
		} else {
			xlsx.SetCellStr(sheet, depositBank, "")
		}

		accountBalance := fmt.Sprintf("%s%d", fst.AccountBalance, row)
		if i < len(dst) {
			xlsx.SetCellStr(sheet, accountBalance, dst[i].AccountBalance)
		} else {
			xlsx.SetCellStr(sheet, accountBalance, "")
		}

		pledgedbalance := fmt.Sprintf("%s%d", fst.Pledgedbalance, row)
		if i < len(dst) {
			xlsx.SetCellStr(sheet, pledgedbalance, dst[i].Pledgedbalance)
		} else {
			xlsx.SetCellStr(sheet, pledgedbalance, "")
		}
	}
}

//使用hxst的excel单元格信息和data的对账单数据填充模版, sheet值固定是"sheet1"
func SetValues(xlsx *excelize.File, sheet string, hxst HuaxiaBankStatement, data *HuaxiaBank) {
	xlsx.SetCellStr(sheet, hxst.PostalCode, data.PostalCode)
	xlsx.SetCellStr(sheet, hxst.ClientCode, data.ClientCode)
	xlsx.SetCellStr(sheet, hxst.CompanyName, data.CompanyName)
	xlsx.SetCellStr(sheet, hxst.DepositBank, data.DepositBank)
	xlsx.SetCellStr(sheet, hxst.ClientAddress, data.ClientAddress)
	xlsx.SetCellStr(sheet, hxst.Contact, data.Contact)
	xlsx.SetCellStr(sheet, hxst.PhoneNumber, data.PhoneNumber)
	xlsx.SetCellStr(sheet, hxst.DueDate, data.DueDate)

	//填充银行回执单
	brs, dbrs := hxst.BankRecipts, data.BankRecipts
	SetStatement(xlsx, sheet, brs.RowStart, brs.MaxLine, brs.Statement, dbrs)

	//填充客户留存
	cp := hxst.ClientCopy
	xlsx.SetCellStr(sheet, cp.ClientCode, data.ClientCode)
	xlsx.SetCellStr(sheet, cp.CompanyName, data.CompanyName)
	xlsx.SetCellStr(sheet, cp.DueDate, data.DueDate)
	SetStatement(xlsx, sheet, cp.RowStart, cp.MaxLine, cp.Statement, dbrs)
}

//excel单元格位置
var stm = HuaxiaBankStatement{
	PostalCode:    "C6",
	ClientCode:    "K6",
	CompanyName:   "C7",
	DepositBank:   "K7",
	ClientAddress: "C8",
	Contact:       "C9",
	PhoneNumber:   "C10",
	DueDate:       "L10",
	BarCode:       "B11",
	BankRecipts: BankRecipts{
		MaxLine:  5,
		RowStart: 14,
		Statement: Statement{
			Account:        "B",
			AccountType:    "D",
			Currency:       "E",
			DepositBank:    "F",
			AccountBalance: "I",
			Pledgedbalance: "J",
		},
	},
	ClientCopy: ClientCopy{
		ClientCode:  "M32",
		CompanyName: "C33",
		DueDate:     "M33",
		MaxLine:     5,
		RowStart:    35,
		Statement: Statement{
			Account:        "B",
			AccountType:    "D",
			Currency:       "F",
			DepositBank:    "H",
			AccountBalance: "K",
			Pledgedbalance: "N",
		},
	},
}

func main() {
	b, err := json.MarshalIndent(&stm, "", "  ")
	if err != nil {
		log.Fatalln(err)
	}
	fmt.Printf("%s\n", b)
	f, err := OpenExcel(filepath.Join("static", "华夏银企对账单 顺丰套打 冀州.xlsx"))
	if err != nil {
		log.Fatalln(err)
	}
	index := "sheet1"
	fmt.Println(index)
	fmt.Println("------------------------------")
	pcode := f.GetCellValue(index, "F3")
	fmt.Println(pcode)
	fmt.Println("------------------------------")
	//读取数据文本，测试中使用本地文件
	b, err = ioutil.ReadFile("tmp/华夏银企对账单 顺丰套打 冀州.csv")
	if err != nil {
		log.Fatalln(err)
	}
	hx, err := ParseCsv(string(b), 16)
	if err != nil {
		log.Fatalln(err)
	}

	for _, data := range hx {
		fmt.Printf("%s, %d\n", data.BarCode, len(data.BankRecipts))
		SetValues(f, index, stm, data)

		//不通编码需要替换此函数
		bc, err := code39.Encode(data.BarCode, false, false)
		//bc, err := code128.Encode(data.BarCode)
		if err != nil {
			log.Printf("%s\n", err)
			continue
		}
		//设定图片大小
		bar, _ := barcode.Scale(bc, 400, 25)
		//添加数值显示
		GenBarCode(400, 40, 12, bar, data.BarCode, "barcode", "static/arial.ttf")
		barcodefile := filepath.Join("barcode", data.BarCode+".png")
		err = f.AddPicture(index, "B11", barcodefile, `{"print_obj": true, "locked": false, lock_aspect_ratio": false, "x_offset": 540, "y_offset", 1500, "x_scale":1.5,"y_scale":1.5}`)
		if err != nil {
			log.Printf("%s\n", err)
			continue
		}
		hxfile := filepath.Join("wait_print", data.BarCode+".xlsx")
		err = f.SaveAs(hxfile)
		if err != nil {
			log.Printf("%s\n", err)
		}
	}
}
