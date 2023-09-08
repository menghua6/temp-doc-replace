package main

import (
	"flag"
	"fmt"
	"net/http"
	"strings"

	"strconv"

	"baliance.com/gooxml/document"
	"github.com/gin-gonic/gin"
	"github.com/shakinm/xlsReader/xls"
)

func main() {
	// referDate := time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local) 日期数字代表与该日期相差了多少天
	api := false
	port := ""
	flag.BoolVar(&api, "api", false, "open api")
	flag.StringVar(&port, "port", "", "port")
	flag.Parse()

	if api {
		r := gin.Default()
		r.GET("/", func(ctx *gin.Context) {
			respose := "" +
			"文档模板替换 \n" +
			"\n" +
			"获取相关文件请求路径 \n" +
			"/example \n"
			ctx.String(200,respose)
		})

		r.StaticFS("/example", http.Dir("example"))
		r.Run(port)
	}

	list, err := xls.OpenFile("list.xls")
	if err != nil {
		fmt.Println("read xls file error:", err.Error())
		return
	}

	sheet, err := list.GetSheet(0)
	if err != nil {
		fmt.Println("get first sheet error:", err.Error())
		return
	}

	dataMap := make(map[string][]string)

	rows := sheet.GetRows()

	types := make([]string, 0)
	typesStruct := rows[0].GetCols()
	for i := 0; i < len(typesStruct); i++ {
		types = append(types, typesStruct[i].GetString())
	}

	for i := 1; i < len(rows); i++ {
		cols := rows[i].GetCols()
		for j := 0; j < len(cols); j++ {
			dataMap[types[j]] = append(dataMap[types[j]], cols[j].GetString())
		}
	}

	num := sheet.GetNumberRows() - 1
	for i := 0; i < num; i++ {
		doc, err := document.Open("template.docx")
		if err != nil {
			fmt.Printf("error opening document: %s", err)
			return
		}
		for _, para := range doc.Paragraphs() {
			for _, run := range para.Runs() {
				text := run.Text()
				for k, v := range dataMap {
					if strings.Contains(text, k) {
						text = strings.ReplaceAll(text, k, v[i])
					}
				}
				run.Clear()
				run.AddText(text)
			}
		}
		doc.SaveToFile("files/" + strconv.Itoa(i+1) + ".docx")
	}
}
