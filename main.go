package main

import (
	"fmt"
	"strings"

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
	room string
	time
}

// func courseSet(nameInfo, timeInfo, roomInfo string) ([]course, error) {
// 	var temp course
// 	courseList := make([]course, 0)
// 	temp.name = nameInfo
// 	temp.room = roomInfo

// 	return courseList
// }

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
		// tempCourseList, err := courseSet(row[2], row[6], row[7])
		// if err != nil {
		// 	return courseList, err
		// }
		// row[6] for timeInfo
		var tempCourse course
		tempCourse.name = row[2]
		tempCourse.room = row[7]
		tempTimeList := timeHandle(row[6])
		for _, tempTime := range tempTimeList {
			tempCourse.time = tempTime
			courseList = append(courseList, tempCourse)
		}
	}
	return courseList, err
}

// Todo:
// 正确流程：先确定时间，获得一个TimeList，然后time := range TimeList进行遍历，把CourseName和CourseRoom传进去，一个个加到courseList

func timeHandle(timeInfo string) []time, error {
	timeList := make([]time, 0)
	//第一遍，筛里面有没有周这个字，两种情况，1-5周，6-10周这种，1,6周，2,7周这种
	//第二遍，做切分，切分的时候根据里面有没有单双周进行一个判断
	timeInfoSlice, err := strings.Split(timeInfo, " ")
	if err != nil {
		return timeList, err
	}
	for _, timePiece := range timeInfoSlice {
		if 这里有单双周() {
			if 判断单双周() {
				单周设为true()
			} else {
				双周设为true()
			}
			把单双周去掉()
		} else if 这里有周() {
			if 正则判断新研还是行政() {
				把前5周后5周设为true()
			} else {
				把对应周设为true()
			}
		} else {
			全部设为true()
		}
	}
	return timeList, err
}
