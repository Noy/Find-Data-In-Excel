package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"bufio"
	"os"
	"strings"
	"io/ioutil"
)

func main() {
	reader := bufio.NewReader(os.Stdin)
	fmt.Println("Enter the Data string:")
	nextText, _ := reader.ReadString('\n')
	files, err := ioutil.ReadDir(".")
	if err != nil { handleError(err) }
	for _, file := range files {
		if strings.HasPrefix(file.Name(), "~") {continue}
		if !strings.HasSuffix(file.Name(), ".xlsx") {continue}
		xlFile, _ := xlsx.OpenFile(file.Name())
		for _, sheet := range xlFile.Sheets {
			for _, row := range sheet.Rows {
				for _, cell := range row.Cells {
					if cell.Value == strings.ToLower(strings.Replace(nextText, "\n", "", -1)) {
						fmt.Println("That pattern occured in: " + file.Name())
					}
				}
			}
		}
	}
}

func handleError(err error) {
	fmt.Println("Something went wrong! Check below:")
	fmt.Println(err)
}

//// OLD STUFF

func holdOn() {
	//xlFile, _ := xlsx.OpenFile(strings.Replace(file.Name(), "\n", "", -1))
	//for _, sheet := range xlFile.Sheets {
	//	for _, row := range sheet.Rows {
	//		for _, cell := range row.Cells {
	//			if !strings.HasSuffix(file.Name(), "xlsx") {continue}
	//			if strings.HasPrefix(file.Name(), "~") {continue}
	//			cellText, _ := cell.String()
	//			if strings.Contains(s, cellText) {
	//				fmt.Println(s + " is there")
	//			}
	//		}
	//	}
	//}
}

func findData() {
	reader := bufio.NewReader(os.Stdin)
	fmt.Println("Enter the Data string:")
	nextText, _ := reader.ReadString('\n')
	fmt.Println("Enter the File name:")
	text, _ := reader.ReadString('\n')
	files, err := ioutil.ReadDir(".")
	for _, file := range files {
		fmt.Println(file)
	}
	excelFilename := text
	xlFile, err := xlsx.OpenFile(strings.Replace(excelFilename, "\n", "", -1))
	if err != nil { handleError(err) }
	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if !strings.HasSuffix(excelFilename, "xlsx") {continue}
				if strings.HasPrefix(excelFilename, "~") {continue}
				cellText, _ := cell.String()
				if strings.Contains(nextText, cellText) {
					fmt.Println(nextText + " is there")
				}
			}
		}
	}
}

func searchInstead() {
	//nextReader := bufio.NewReader(os.Stdin)
	//text, _ := nextReader.ReadString('\n')
	//excelFilename =  text + ".xlsx"
	//fmt.Println("Enter the website:")
	//xlFile, err := xlsx.OpenFile(strings.Replace(excelFilename, "\n", "", -1))
}