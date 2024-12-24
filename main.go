package main

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

func processExcel(filePath string) error {
	// 打开工作簿
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("无法打开文件: %v", err)
	}
	defer f.Close()

	// 获取所有工作表名称
	sheetNames := f.GetSheetList()
	for _, sheetName := range sheetNames {
		// 替换工作表名称
		if sheetName == "甲烷非甲烷分析仪" {
			err := f.SetSheetName(sheetName, "在线NMHC监测仪")
			if err != nil {
				return fmt.Errorf("无法重命名工作表: %v", err)
			}
			fmt.Printf("工作表名称已从 '%s' 替换为 '在线NMHC监测仪'\n", sheetName)
		}
	}

	// 获取活动工作表名称
	activeSheetName := f.GetSheetName(f.GetActiveSheetIndex())

	// 定义红色填充颜色
	redFill, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FF0000"},
			Pattern: 1,
		},
	})
	if err != nil {
		return fmt.Errorf("无法定义红色填充样式: %v", err)
	}

	// 获取最大行数和列数
	rows, err := f.GetRows(activeSheetName)
	if err != nil {
		return fmt.Errorf("无法获取行: %v", err)
	}
	maxRow := len(rows)
	maxColumn := 0
	for _, row := range rows {
		if len(row) > maxColumn {
			maxColumn = len(row)
		}
	}

	// 正则表达式，用于删除括号及其中的内容
	re := regexp.MustCompile(`\([^)]*\)`)

	// 遍历整个表格处理替换
	for row := 1; row <= maxRow; row++ {
		for col := 1; col <= maxColumn; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return fmt.Errorf("无法获取单元格名称: %v", err)
			}

			value, err := f.GetCellValue(activeSheetName, cell)
			if err != nil {
				return fmt.Errorf("无法获取单元格值: %v", err)
			}

			originalValue := value // 保留原始值，用于判断是否需要标记红色

			// 替换指定字符串，不设置红色背景
			if strings.Contains(value, "甲烷非甲烷分析仪") {
				value = strings.ReplaceAll(value, "甲烷非甲烷分析仪", "在线NMHC监测仪")
			}
			if strings.Contains(value, "总烃(ppbv)") {
				value = strings.ReplaceAll(value, "总烃(ppbv)", "总烃(ppbC)")
			}

			// 替换“间、对-二甲苯”为“间/对二甲苯”，并设置红色背景
			if strings.Contains(value, "间、对-二甲苯") {
				value = strings.ReplaceAll(value, "间、对-二甲苯", "间/对二甲苯")
			}

			// 如果是第3行及之后，删除括号及其中的内容，并设置红色背景
			if row >= 3 && re.MatchString(value) {
				value = re.ReplaceAllString(value, "")
				err = f.SetCellStyle(activeSheetName, cell, cell, redFill)
				if err != nil {
					return fmt.Errorf("无法设置单元格样式: %v", err)
				}
			}

			// 如果单元格值被修改，更新值
			if value != originalValue {
				err = f.SetCellValue(activeSheetName, cell, strings.TrimSpace(value))
				if err != nil {
					return fmt.Errorf("无法更新单元格值: %v", err)
				}
			}
		}
	}

	// 获取文件的基本名称并生成输出路径
	baseName := filepath.Base(filePath)
	outputPath := "processed_" + baseName

	// 保存修改后的工作簿
	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("无法保存文件: %v", err)
	}

	fmt.Printf("文件已处理并保存为: %s\n", outputPath)
	return nil
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("请提供文件名作为参数，例如：./program 45vocs2.xlsx")
		return
	}
	filePath := os.Args[1]
	if err := processExcel(filePath); err != nil {
		fmt.Println("处理Excel文件时出错:", err)
	}
}
