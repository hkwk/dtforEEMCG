/*
Created on Tue Oct 22 18:26:13 2024
@author: Kunlun HUANG
github.com/kunlunh
*/
package main

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"

	"github.com/xuri/excelize/v2"
)

func processExcel(filePath string) error {
	// 打开工作簿
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("无法打开文件: %v", err)
	}
	defer f.Close()

	// 获取活动工作表的名称
	sheetName := f.GetSheetName(f.GetActiveSheetIndex())

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
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("无法获取行: %v", err)
	}
	maxRow := len(rows)
	maxColumn := len(rows[0])

	// 正则表达式，用于删除括号及其中的内容
	re := regexp.MustCompile(`\([^)]*\)`)

	// 从第三行开始遍历所有单元格
	for row := 3; row <= maxRow; row++ {
		for col := 1; col <= maxColumn; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return fmt.Errorf("无法获取单元格名称: %v", err)
			}

			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return fmt.Errorf("无法获取单元格值: %v", err)
			}

			// 检查并删除括号及其中的内容
			if re.MatchString(value) {
				newValue := re.ReplaceAllString(value, "")
				err = f.SetCellValue(sheetName, cell, newValue)
				if err != nil {
					return fmt.Errorf("无法更新单元格值: %v", err)
				}

				// 设置单元格背景为红色
				err = f.SetCellStyle(sheetName, cell, cell, redFill)
				if err != nil {
					return fmt.Errorf("无法设置单元格样式: %v", err)
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
