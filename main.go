package main

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type time struct {
	start int
	end   int
	day   int      // 1 for Monday etc
	week  [11]bool // 12345678910
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
		tempTimeList, _ := timeHandle(row[6])
		for _, tempTime := range tempTimeList {
			tempCourse.time = tempTime
			courseList = append(courseList, tempCourse)
		}
	}
	return courseList, err
}

// Todo:
// 正确流程：先确定时间，获得一个TimeList，然后time := range TimeList进行遍历，把CourseName和CourseRoom传进去，一个个加到courseList

func timeHandle(timeInfo string) ([]time, error) {
	var tempTime time
	timeList := make([]time, 0)
	//第一遍，筛里面有没有周这个字，两种情况，1-5周，6-10周这种，1,6周，2,7周这种
	//第二遍，做切分，切分的时候根据里面有没有单双周进行一个判断
	splitFunc := func(r rune) bool { return r == ' ' || r == '(' || r == ')' }
	timeInfoSlice := strings.FieldsFunc(timeInfo, splitFunc)
	//Check
	fmt.Println(timeInfoSlice)
	//新生研讨课
	judgeXy, err := regexp.MatchString("[0-9]-[0-9]周", timeInfoSlice[len(timeInfoSlice)-1])
	if err != nil {
		return timeList, err
	}
	if judgeXy {
		if timeInfoSlice[len(timeInfoSlice)-1][0] == '1' {
			for i := 1; i <= 5; i++ {
				tempTime.week[i] = true
			}
		} else {
			for i := 6; i <= 10; i++ {
				tempTime.week[i] = true
			}
		}
	}
	//形式政策课
	judgeXszc, err := regexp.MatchString("[0-9]周,[0-9]周", timeInfoSlice[len(timeInfoSlice)-1])
	if err != nil {
		return timeList, err
	}
	if judgeXszc {
		temp, err := strconv.Atoi(timeInfoSlice[len(timeInfoSlice)-1][0:1])
		if err != nil {
			return timeList, err
		}
		tempTime.week[temp] = true
		tempTime.week[temp+5] = true
	}
	//下面进行单双周判定
	splitFunc = func(r rune) bool {
		return r == '一' || r == '二' || r == '三' || r == '四' || r == '五' || r == '单' || r == '双' || r == '-'
	}
	for _, timePiece := range timeInfoSlice {
		if 单() {
			for i := 1; i <= 10; i += 2 {
				tempTime.week[i] = true
			}
		}
		if 双() {
			for i := 2; i <= 10; i += 2 {
				tempTime.week[i] = true
			}
		}
		switch 12345 {

		}
		timePieceSlice := strings.FieldsFunc(timePiece, splitFunc)
		startTime, err := strconv.Atoi(timePieceSlice[0])
		if err != nil {
			return timeList, err
		}
		endTime, err := strconv.Atoi(timePieceSlice[1])
		if err != nil {
			return timeList, err
		}
		tempTime.start = setTime(startTime, 1)
		tempTime.end = setTime(endTime, 2)
	}
	return timeList, err
}

func setTime(timeIdx, timeType int) int {
	// timeIdx := 1~12
	// timeType := 1,2 1 for StartTime, 2 for EndTime
	if timeType == 1 {
		switch timeIdx {
		case 1:
			return 800
		case 2:
			return 855
		case 3:
			return 1000
		case 4:
			return 1055
		case 5:
			return 1300
		case 6:
			return 1355
		case 7:
			return 1500
		case 8:
			return 1555
		case 9:
			return 1800
		case 10:
			return 1855
		case 11:
			return 2000
		case 12:
			return 2055
		}
	} else {
		switch timeIdx {
		case 1:
			return 845
		case 2:
			return 940
		case 3:
			return 1045
		case 4:
			return 1140
		case 5:
			return 1345
		case 6:
			return 1440
		case 7:
			return 1545
		case 8:
			return 1640
		case 9:
			return 1845
		case 10:
			return 1940
		case 11:
			return 2045
		case 12:
			return 2140
		}
	}
}
