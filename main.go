package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type User struct {
	FirstName string `json:"firstname"`
	LastName  string `json:"lastName"`
	Email     string `json:"email"`
}

func exportCSV(user []User) {
	file, err := os.Create("user.csv")
	if err != nil {
		fmt.Println("Error creating CSV file:", err)
		return
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	header := []string{"firstname", "lastName", "email"}
	if err := writer.Write(header); err != nil {
		fmt.Println("Error writing CSV header:", err)
		return
	}

	for _, person := range user {
		row := []string{person.FirstName, person.LastName, person.Email}
		if err := writer.Write(row); err != nil {
			fmt.Println("Error writing CSV data:", err)
			return
		}
	}

	fmt.Println("CSV file 'user.csv' has been created in the current directory.")
}

func exportXLSX(user []User) {
	xlsx := excelize.NewFile()
	sheetName := "Sheet1"
	xlsx.SetSheetName(xlsx.GetSheetName(1), sheetName)

	xlsx.SetCellValue(sheetName, "A1", "firstname")
	xlsx.SetCellValue(sheetName, "B1", "lastName")
	xlsx.SetCellValue(sheetName, "C1", "email")

	for i, person := range user {
		rowIndex := i + 2
		xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), person.FirstName)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), person.LastName)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIndex), person.Email)
	}

	if err := xlsx.SaveAs("user.xlsx"); err != nil {
		fmt.Println("Error saving XLSX file:", err)
		return
	}

	fmt.Println("XLSX file 'user.xlsx' has been created in the current directory.")
}

func exportXML(user []User) {
	xmlData := `<user>`
	for _, person := range user {
		xmlData += fmt.Sprintf(
			`<person><firstname>%s</firstname><lastName>%s</lastName><email>%s</email></person>`,
			person.FirstName, person.LastName, person.Email,
		)
	}
	xmlData += `</user>`

	file, err := os.Create("user.xml")
	if err != nil {
		fmt.Println("Error creating XML file:", err)
		return
	}
	defer file.Close()

	_, err = file.WriteString(xmlData)
	if err != nil {
		fmt.Println("Error writing XML data:", err)
		return
	}

	fmt.Println("XML file 'user.xml' has been created in the current directory.")
}

func exportHTML(user []User) {
	htmlData := `<html><head><title>People Data</title></head><body><table><tr><th>First Name</th><th>Last Name</th><th>Email</th></tr>`
	for _, person := range user {
		htmlData += fmt.Sprintf(
			`<tr><td>%s</td><td>%s</td><td>%s</td></tr>`,
			person.FirstName, person.LastName, person.Email,
		)
	}
	htmlData += `</table></body></html>`

	file, err := os.Create("user.html")
	if err != nil {
		fmt.Println("Error creating HTML file:", err)
		return
	}
	defer file.Close()

	_, err = file.WriteString(htmlData)
	if err != nil {
		fmt.Println("Error writing HTML data:", err)
		return
	}

	fmt.Println("HTML file 'user.html' has been created in the current directory.")
}

func exportToFile(input, exportType string) {
	var user []User
	if err := json.Unmarshal([]byte(input), &user); err != nil {
		fmt.Println("Error unmarshaling JSON:", err)
		return
	}
	exportType = strings.ToLower(exportType)
	switch exportType {
	case "csv":
		exportCSV(user)
	case "xlsx":
		exportXLSX(user)
	case "xml":
		exportXML(user)
	case "html":
		exportHTML(user)
	}
}

func main() {
	jsonData := `[{"firstname" : "Alex", "lastName" : "Mat", "email" : "alex.matt@gmail.com"}]`
	exportToFile(jsonData, "csv")
}
