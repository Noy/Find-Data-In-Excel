package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"bufio"
	"os"
	"strings"
	"io/ioutil"
)

// Commented for tutorial purposes, made by Noy Hillel @ TrafficLabel.com

func main() {
	// Get the reader for user input
	reader := bufio.NewReader(os.Stdin)
	// Print something on the screen, knowing that the program has started
	fmt.Println("Enter the data string:")
	// Get the text value from the user, it's called 'nextText' because I didn't rename it after testing lol
	nextText, _ := reader.ReadString('\n')
	// Get the files which we need to read. 'ReadDir' returns an error, so we need to give it one
	files, err := ioutil.ReadDir(".") // You could also get an input for the directory, but it's not necessary.
	// Check if the error occurs, if it does, execute our simple handle error function
	if err != nil { handleError(err) }
	// This is in order to print if data was found or not, I'll explain as I go down..
	count := 0
	// Iterate through the files in the directory
	for _, file := range files {
		// This check was throwing an invalid memory address or nil pointer dereference error, hidden/cached files cannot be checked
		if strings.HasPrefix(file.Name(), "~") {continue}
		// Check if our file is an excel document, if not, ignore it (e.g. our main.go file)
		if !strings.HasSuffix(file.Name(), ".xlsx") {continue}
		// Get the excel file and open it with the name of the file within the loop
		xlFile, _ := xlsx.OpenFile(file.Name())
		// Iterate through the sheet
		for _, sheet := range xlFile.Sheets {
			// Iterate through the Rows
			for _, row := range sheet.Rows {
				// Iterate through the cells
				for _, cell := range row.Cells {
					// In our case, we are checking if the data we entered matches any cell in our sheets
					/*
					  In my case I was searching upper case data to lower cased data, that's why I did strings.ToLower.
					  Change it to: if cell.Value == strings.Replace(nextText, "\n", "", -1) { If you just want to do the check
					 */
					if cell.Value == strings.ToLower(strings.Replace(nextText, "\n", "", -1)) {
						// Increment our count
						count++
						// Let the user know where the pattern occurred
						fmt.Println("That pattern occured in: " + file.Name())
					}
				}
			}
		}
	}
	/*
	   If the count is 0 it would mean that no data was passed, therefore we safely know it didn't match our string
	   which we entered to match any string in the excel document. Thought this was a smart way to do the check lol
	 */
	if count == 0 { fmt.Println("Could not find that pattern") } // Wanted a confirmation message
}

// Handle Error function
func handleError(err error) {
	// Simply print this string..
	fmt.Println("Something went wrong! Check below:")
	// ..and the error
	fmt.Println(err)
}