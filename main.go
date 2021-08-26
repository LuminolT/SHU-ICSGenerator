package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type course struct {
	courseName string
	courseTime string
	courseRoom string
}

func main() {
	courseList, err := readTable("course_table.xlsx", "Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(courseList)
}

func (this *course) timeSet() error {

}

func readTable(fileName, sheetName string) ([]course, error) {
	//ReadCourse
	courseList := make([]course, 0)
	xlFile, err := excelize.OpenFile(fileName)
	if err != nil {
		return courseList, err
	}
	rows, err := xlFile.GetRows(sheetName)
	if err != nil {
		return courseList, err
	}
	for rowIdx, row := range rows {
		tempRunes := []rune(row[0]) //Row Front Slice
		if rowIdx == 0 {
			continue
		}
		if tempRunes[0] > 'Z' {
			break
		}
		tempCourse := course{row[2], row[6], row[7]}
		courseList = append(courseList, tempCourse)
	}
	return courseList, err
}
