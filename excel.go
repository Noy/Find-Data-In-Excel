package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"bufio"
	"os"
	"strings"
	"path/filepath"
	"time"
)

func main() {
	findData()
}

func findData() {
	reader := bufio.NewReader(os.Stdin)
	fmt.Println("Enter the Data:")
	nextText, _ := reader.ReadString('\n')
	searchDir := "."
	fileList := []string{}
	err := filepath.Walk(searchDir, func(path string, f os.FileInfo, err error) error {
		fileList = append(fileList, path)
		return nil
	})
	var cellTextString string
	var fileTextString string
	for _, file := range fileList {
		if !strings.HasSuffix(file, "xlsx") {continue}
		if strings.HasPrefix(file, "~") {continue}
		xlFile, err := xlsx.OpenFile(file)
		if err != nil { handleError(err) }
		for _, sheet := range xlFile.Sheets {
			for _, row := range sheet.Rows {
				for _, cell := range row.Cells {
					cellText, _ := cell.String()
					if strings.Contains(nextText, cellText) {
						fileTextString = file
						cellTextString = "*** " + cellText + " ***"
					}
				}
				time.Sleep(2)
			}
		}
		fmt.Println("Searching: " + fileTextString + " for cell: " + cellTextString)
	}
	if err != nil {
		handleError(err)
	}
}

func handleError(err error) {
	fmt.Println("Something went wrong! Check below:")
	fmt.Println(err)
}

func searchInstead() {
	//nextReader := bufio.NewReader(os.Stdin)
	//text, _ := nextReader.ReadString('\n')
	//excelFilename =  text + ".xlsx"
	//fmt.Println("Enter the website:")
	//xlFile, err := xlsx.OpenFile(strings.Replace(excelFilename, "\n", "", -1))
}