package main

import (
	"crypto/sha256"
	"fmt"
	"io"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	ics "github.com/arran4/golang-ical"
	"github.com/xuri/excelize/v2"
)

type courseTime struct {
	start     [2]int
	end       [2]int
	day       int     // 1 for Monday etc
	week      [11]int // 12345678910 week[0]: 1-新研, 2-行政, 3-单周, 4-双周, 5-正常
	startWeek int
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
	// Initiate Time
	var SYEAR, SDAY int
	var SMONTH time.Month
	var GapWeek int
	var EndWeek int
	fmt.Println("请输入本学期第一周周一的年月日（如2021-9-6）：")
	fmt.Scanf("%d-%d-%d\n", &SYEAR, &SMONTH, &SDAY)
	if SMONTH > time.September {
		fmt.Println("检测到本学期为冬季学期，请输入寒假前最后一周的周数和寒假时长（如8-4表示第八周结束开始放假，放4周）")
		fmt.Scanf("%d-%d\n", &EndWeek, &GapWeek)
	}
	// SYEAR = 2021
	// SMONTH = 9
	// SDAY = 6
	// Read Table
	courseList, err := readTable("course_table.xlsx", "Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	TIME_LOCATION, err := time.LoadLocation("Asia/Shanghai")
	if err != nil {
		TIME_LOCATION = time.FixedZone("CST", 8*3600)
		fmt.Println(err)
	}
	// fmt.Println(courseList)
	cal := ics.NewCalendar()
	cal.SetMethod(ics.MethodRequest)
	for _, coursePiece := range courseList {
		// fmt.Println(coursePiece) // Test

		// 🤬🤬🤬🤬🤬
		tempStartTime := time.Date(SYEAR, SMONTH, SDAY, coursePiece.start[0], coursePiece.start[1], 0, 0, TIME_LOCATION)
		tempEndTime := time.Date(SYEAR, SMONTH, SDAY, coursePiece.end[0], coursePiece.end[1], 0, 0, TIME_LOCATION)
		// fmt.Println(coursePiece.day)
		tempStartTime = tempStartTime.AddDate(0, 0, coursePiece.day-1)
		// tempStartTime = tempStartTime.AddDate(0, 0, 7*(coursePiece.startWeek-1))
		tempEndTime = tempEndTime.AddDate(0, 0, coursePiece.day-1)
		// tempEndTime = tempEndTime.AddDate(0, 0, 7*(coursePiece.startWeek-1))
		// 由于要加入冬季学期寒假的判断，下面就不用重复的了。
		/*
			switch coursePiece.week[0] {
			case 1: //新生研讨课
				event.AddRrule(fmt.Sprintf("FREQ=WEEKLY;INTERVAL=%d;COUNT=%d", 1, 5))
			case 2: //形势政策课
				event.AddRrule(fmt.Sprintf("FREQ=WEEKLY;INTERVAL=%d;COUNT=%d", 5, 2))
			case 3: //单周
				event.AddRrule(fmt.Sprintf("FREQ=WEEKLY;INTERVAL=%d;COUNT=%d", 2, 5))
			case 4: //双周
				event.AddRrule(fmt.Sprintf("FREQ=WEEKLY;INTERVAL=%d;COUNT=%d", 2, 5))
			case 5: //正常
				event.AddRrule(fmt.Sprintf("FREQ=WEEKLY;INTERVAL=%d;COUNT=%d", 1, 10))
			}
		*/
		// fmt.Println(tempStartTime, "\n", tempEndTime)
		// fmt.Println(coursePiece.name, coursePiece.week)
		for i := 1; i <= 10; i++ {
			if coursePiece.week[i] == 1 {
				//Hash ID Check
				h := sha256.New()
				plaintext := fmt.Sprintf("%s%d%d", coursePiece.name, coursePiece.courseTime.day, coursePiece.courseTime.start)
				// fmt.Println(plaintext)
				h.Write([]byte(plaintext))
				id := fmt.Sprintf("%x@%s", h.Sum(nil), "ical") // get HashValue in SHA256, used as EVENTID
				// new a pointer of cal.EEvent
				event := cal.AddEvent(id)
				// Check Real Week
				finalStartTime := tempStartTime.AddDate(0, 0, 7*(i-1))
				finalEndTime := tempEndTime.AddDate(0, 0, 7*(i-1))
				if i > EndWeek {
					finalStartTime = finalStartTime.AddDate(0, 0, 7*GapWeek)
					finalEndTime = finalEndTime.AddDate(0, 0, 7*GapWeek)
				}
				event.SetStartAt(finalStartTime)
				event.SetEndAt(finalEndTime)
				event.SetSummary(coursePiece.name)
				event.SetLocation(coursePiece.room)
				alarm := event.AddAlarm()
				// advancedTime := fmt.Sprintf("-PT%dM", ADVANCEDTIME)
				alarm.SetTrigger("-PT10M")
			}
		}
		// alarm.SetAction()
	}
	// fmt.Println(cal.Serialize())
	err2 := WriteFile("./output.ics", []byte(cal.Serialize()), 0666)
	if err2 != nil {
		fmt.Println(err)
	} else {
		fmt.Println("成功写入")
	}
	fmt.Println("按任意键退出")
	b := make([]byte, 1)
	os.Stdin.Read(b)
	os.Stdin.Read(b)
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

func timeHandle(timeInfo string) ([]courseTime, error) {
	// fmt.Println(timeInfo)
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
			tempTime.startWeek = 1
			for i := 1; i <= 5; i++ {
				tempTime.week[i] = 1
			}
		} else {
			tempTime.startWeek = 6
			for i := 6; i <= 10; i++ {
				tempTime.week[i] = 1
			}
		}
		tempTime.week[0] = 1
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
		tempTime.startWeek = temp
		tempTime.week[temp] = 1
		tempTime.week[temp+5] = 1
		tempTime.week[0] = 2
	}
	//下面进行单双周判定
	splitFunc = func(r rune) bool {
		return r == '一' || r == '二' || r == '三' || r == '四' || r == '五' || r == '单' || r == '双' || r == '-'
	}
	// 1st Slice [一1-2单]
	for _, timePiece := range timeInfoSlice {

		// if timePiece == "上机" {
		// 	continue
		// }
		if tempTime.week[0] == 3 || tempTime.week[0] == 4 {
			tempTime.week[0] = 0
		}
		if strings.Contains(timePiece, "单") {
			// fmt.Println("Checked")
			for i := 1; i <= 10; i += 2 {
				tempTime.week[i] = 1
				tempTime.week[i+1] = 0
			}
			tempTime.week[0] = 3
			tempTime.startWeek = 1
			// fmt.Println(tempTime)
		}
		// fmt.Println(tempTime)
		if strings.Contains(timePiece, "双") {
			for i := 1; i <= 10; i += 2 {
				tempTime.week[i] = 0
				tempTime.week[i+1] = 1
			}
			tempTime.week[0] = 4
			tempTime.startWeek = 2
		}
		if tempTime.week[0] == 0 {
			for i := 1; i <= 10; i++ {
				tempTime.week[i] = 1
			}
			tempTime.week[0] = 5
			tempTime.startWeek = 1
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
		default:
			continue // ╰(*°▽°*)╯防止“上机”“学院机房上机”等情况
		}
		// fmt.Println(tempTime)
		// fmt.Println(timePiece)
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
		tempTime.start[0], tempTime.start[1] = setTime(startTime, 1)
		tempTime.end[0], tempTime.end[1] = setTime(endTime, 2)
		timeList = append(timeList, tempTime)
		// fmt.Println(tempTime)
	}
	return timeList, err
}

func setTime(timeIdx, timeType int) (int, int) {
	// timeIdx := 1~12
	// timeType := 1,2 1 for StartTime, 2 for EndTime
	if timeType == 1 {
		switch timeIdx {
		case 1:
			return 8, 00
		case 2:
			return 8, 55
		case 3:
			return 10, 00
		case 4:
			return 10, 55
		case 5:
			return 13, 00
		case 6:
			return 13, 55
		case 7:
			return 15, 00
		case 8:
			return 15, 55
		case 9:
			return 18, 00
		case 10:
			return 18, 55
		case 11:
			return 20, 00
		case 12:
			return 20, 55
		}
	} else {
		switch timeIdx {
		case 1:
			return 8, 45
		case 2:
			return 9, 40
		case 3:
			return 10, 45
		case 4:
			return 11, 40
		case 5:
			return 13, 45
		case 6:
			return 14, 40
		case 7:
			return 15, 45
		case 8:
			return 16, 40
		case 9:
			return 18, 45
		case 10:
			return 19, 40
		case 11:
			return 20, 45
		case 12:
			return 21, 40
		}
	}
	return 0, 0
}

func WriteFile(filename string, data []byte, perm os.FileMode) error {
	f, err := os.OpenFile(filename, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, perm)
	if err != nil {
		return err
	}
	n, err := f.Write(data)
	if err == nil && n < len(data) {
		err = io.ErrShortWrite
	}
	if err1 := f.Close(); err == nil {
		err = err1
	}
	return err
}
