package main

import (
	"flag"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

var m = make(map[string]bool)
var a = []string{}
var waitChan = make(chan struct{}, 100)
var imgStartIndex int
var imgEndIndex int
var sheetNames string
var excelName string

func init() {
	flag.StringVar(&excelName, "excelName", "2024圖片下載測試.xlsx", "Excel 名稱")
	flag.IntVar(&imgStartIndex, "start", 1, "抓照片的起始位置，一般為 1 開始")
	flag.IntVar(&imgEndIndex, "end", 10, "抓照片的結束位置")
	flag.StringVar(&sheetNames, "sheetNames", "1130", "sheetA,sheetB,sheetC")
}

func add(s string) {
	if m[s] {
		return // Already in the map
	}
	a = append(a, s)
	m[s] = true
}

func main() {
	flag.Parse()

	for _, sheetName := range strings.Split(sheetNames, ",") {
		err := os.RemoveAll(sheetName)
		if err != nil {
			log.Fatal(err)
		}
	}

	// if _, err := os.Stat("img"); os.IsNotExist(err) {
	// 	os.Mkdir("img", 0755)
	// }

	fmt.Println("讀取 91 xlsx")
	Excel91, err := excelize.OpenFile(excelName)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := Excel91.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	allSheets := strings.Split(sheetNames, ",")
	for _, sheet := range allSheets {
		if _, err := os.Stat(sheet); os.IsNotExist(err) {
			os.Mkdir(sheet, 0755)
		}
		row91s, err := Excel91.GetRows(sheet)
		if err != nil {
			fmt.Println(err)
			return
		}

		for rowI, row := range row91s {
			if rowI > 0 {
				for i, colCell := range row {
					if i == 1 {

						// for i := 2; i < 6; i++ {

						// 	add(colCell[0:9] + "_0" + strconv.Itoa(i))
						// }

						// for i := 6; i < 11; i++ {

						// 	if i == 10 {
						// 		add(colCell[0:9] + "_" + strconv.Itoa(i))
						// 	} else {
						// 		add(colCell[0:9] + "_0" + strconv.Itoa(i))
						// 	}
						// }

						for i := imgStartIndex; i <= imgEndIndex; i++ {

							if i >= 10 {
								// add(colCell + "_" + strconv.Itoa(i) + "@" + sheet)
								add(colCell + "_01.jpg" + "@" + sheet)
							} else {
								// add(colCell + "_0" + strconv.Itoa(i) + "@" + sheet)
								add(colCell + "_01.jpg" + "@" + sheet)
							}
						}
					}
				}
			}
		}
	}

	wg := new(sync.WaitGroup)
	num := len(a)
	wg.Add(num)
	fmt.Println(a)

	for _, s := range a {
		waitChan <- struct{}{}
		go downloadIMG(s, wg)
	}
	wg.Wait()
	fmt.Println(len(a), "張圖片全部下載完成")
}

func downloadIMG(s string, wg *sync.WaitGroup) {
	defer wg.Done()

	loop := true
	for loop {
		// https://www.peachjohn.co.jp/img/goods/L/102908009_01.jpg
		url := fmt.Sprintf(`https://www.peachjohn.co.jp/img/goods/L/%s`, strings.Split(s, "@")[0])
		fmt.Println(url)
		// don't worry about errors
		response, e := http.Get(url)
		if e != nil {
			fmt.Println(e.Error())
			log.Fatal(e)
		}
		// defer response.Body.Close()

		//open a file for writing
		strings.Split(s, "@")

		// tmpImgName := strings.Split(strings.Split(s, "@")[0], "_")
		// newImgName := tmpImgName[0] + tmpImgName[1]
		newImgName := strings.Split(s, "@")[0]
		// file, err := os.Create(fmt.Sprintf(`%s\%s.png`, strings.Split(s, "@")[1], newImgName))
		file, err := os.Create(fmt.Sprintf(`%s/%s`, strings.Split(s, "@")[1], newImgName))
		if err != nil {
			fmt.Println(err.Error())
			log.Fatal(err)
		}
		// defer file.Close()
		// Use io.Copy to just dump the response body to the file. This supports huge files
		_, err = io.Copy(file, response.Body)
		if err != nil {
			fmt.Println(err.Error())
			log.Fatal(err)
		}
		g, _ := file.Stat()
		if g.Size() > 321 {
			loop = false
		} else {

			// e := os.Remove(fmt.Sprintf(`%s/%s.png`, strings.Split(s, "@")[1], strings.Split(s, "@")[0]))
			tmpImgName := strings.Split(strings.Split(s, "@")[0], "_")
			newImgName := tmpImgName[0] + tmpImgName[1]
			e := os.Remove(fmt.Sprintf(`%s\%s.png`, strings.Split(s, "@")[1], newImgName))
			if e != nil {
				log.Fatal(e)
			}
			time.Sleep(5 * time.Second)
		}
		response.Body.Close()
		file.Close()
	}
	<-waitChan
}
