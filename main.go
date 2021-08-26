package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type time struct {
	start int
	end   int
	day   int      // 1 for Monday etc
	week  [10]bool // 12345678910
}

type course struct {
	name string
	time []time
	room string
}

func courseSet(nameInfo, timeInfo, roomInfo string) ([]course, error) {
	var temp course
	courseList := make([]course, 0)
	temp.name = nameInfo
	temp.room = roomInfo

	return courseList
}

func main() {
	courseList, err := readTable("course_table.xlsx", "Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(courseList)
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
		//每次读到时间，要先进行切片处理
		tempCourseList, err := courseSet(row[2], row[6], row[7])
		if err != nil {
			return courseList, err
		}
		for course := range tempCourseList
		courseList = append(courseList, tempCourse)
	}
	return courseList, err
}

// Todo:
// 正确流程：先确定时间，获得一个TimeList，然后time := range TimeList进行遍历，把CourseName和CourseRoom传进去，一个个加到courseList