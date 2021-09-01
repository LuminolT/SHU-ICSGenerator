package main

import (
	"crypto/sha256"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"time"

	ics "github.com/arran4/golang-ical"
	"github.com/xuri/excelize/v2"
)

type courseTime struct {
	start int
	end   int
	day   int      // 1 for Monday etc
	week  [11]bool // 12345678910
}

type course struct {
	name string
	room string
	courseTime
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
	// fmt.Println(courseList)
	cal := ics.NewCalendar()
	cal.SetMethod(ics.MethodRequest)
	for _, coursePiece := range courseList {
		// fmt.Println(coursePiece) // Test
		h := sha256.New()
		plaintext := fmt.Sprintf("%s%d%d", coursePiece.name, coursePiece.courseTime.day, coursePiece.courseTime.start)
		fmt.Println(plaintext)
		h.Write([]byte(plaintext))
		id := fmt.Sprintf("%x@%s", h.Sum(nil), "ical-relay") // get HashValue in SHA256, used as EVENTID
		event := cal.AddEvent(id)
		event.SetStartAt(time.Now())
		event.SetEndAt(time.Now())
		event.SetSummary(coursePiece.name)
		event.SetLocation(coursePiece.room)
		event.AddRrule(fmt.Sprintf("FREQ=WEEKLY;INTERVAL=%d;COUNT=%d", 1, 10))
	}
	// fmt.Println(cal.Serialize())
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
			tempCourse.courseTime = tempTime
			courseList = append(courseList, tempCourse)
		}
	}
	return courseList, err
}

// Todo:
// 正确流程：先确定时间，获得一个TimeList，然后time := range TimeList进行遍历，把CourseName和CourseRoom传进去，一个个加到courseList

func timeHandle(timeInfo string) ([]courseTime, error) {
	var tempTime courseTime
	timeList := make([]courseTime, 0)
	//第一遍，筛里面有没有周这个字，两种情况，1-5周，6-10周这种，1,6周，2,7周这种
	//第二遍，做切分，切分的时候根据里面有没有单双周进行一个判断
	splitFunc := func(r rune) bool { return r == ' ' || r == '(' || r == ')' }
	timeInfoSlice := strings.FieldsFunc(timeInfo, splitFunc)
	//Check
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
		tempTime.week[0] = true
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
		tempTime.week[0] = true
	}
	//下面进行单双周判定
	splitFunc = func(r rune) bool {
		return r == '一' || r == '二' || r == '三' || r == '四' || r == '五' || r == '单' || r == '双' || r == '-'
	}
	// 1st Slice [一1-2单]
	for _, timePiece := range timeInfoSlice {
		if strings.Contains(timePiece, "单") {
			for i := 1; i <= 10; i += 2 {
				tempTime.week[i] = true
			}
			tempTime.week[0] = true
		}
		if strings.Contains(timePiece, "双") {
			for i := 2; i <= 10; i += 2 {
				tempTime.week[i] = true
			}
			tempTime.week[0] = true
		}
		if !tempTime.week[0] {
			for i := 1; i <= 10; i++ {
				tempTime.week[i] = true
			}
			tempTime.week[0] = true
		}
		switch timePiece[0:3] {
		case "一":
			tempTime.day = 1
		case "二":
			tempTime.day = 2
		case "三":
			tempTime.day = 3
		case "四":
			tempTime.day = 4
		case "五":
			tempTime.day = 5
		}
		// 2nd Slice [1 2]
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
		timeList = append(timeList, tempTime)
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
	return 0
}
