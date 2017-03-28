package main 
import (
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"github.com/tealeg/xlsx"
	"os"
)

var fileConf, err = os.Open("config.json")
var decoder = json.NewDecoder(fileConf)
var config = Config{}

type json_struct struct {
	// Описываем структуру JSON'a
	Data struct {
		Header [] string 
		Page string
		Data [][]string
		Filename string
	}
}

type Config struct {
    FolderPath    	string
	LogFile 		string
	HostPort 		string
}




func main() {
	if err != nil {
		log.Fatal("No config file (e.g. config.json at same dir)")
	}
	
	err = decoder.Decode(&config)
	
	if err != nil {
	  fmt.Println("error:", err)
	}
	
	// open a file
	f, err := os.OpenFile(config.LogFile, os.O_APPEND | os.O_CREATE | os.O_RDWR, 0666)
	//
	if err != nil {
	  log.Fatal("error opening file: " + config.LogFile)
	}
	
	log.SetOutput(f)
	log.SetFlags( log.LstdFlags | log.Lmicroseconds)
	
	log.Output(1, "Opening log file... " + config.LogFile + " ... OK")
	

	// don't forget to close it
	defer f.Close()

	// assign it to the standard logger
	log.Output(1, "Starting excelExporter service, http point: " + config.HostPort + "/json2excel")
		
	//ставим обработчик 
	http.HandleFunc("/json2excel", httpExcelHandler)
	log.Fatal(http.ListenAndServe(config.HostPort, nil))
}





func addRows(s [][]string, cell *xlsx.Cell, row *xlsx.Row, sheet *xlsx.Sheet) {
	for k := range s {
    	for v := range s[k] {
    		cell = row.AddCell()
			cell.Value = s[k][v]
    	}
    	row = sheet.AddRow()
    }
}

func httpExcelHandler(w http.ResponseWriter, r *http.Request) {

	if r.Method != "POST" {
		fmt.Printf("Error http method, only POST allowed")
		http.Error(w, "wrong method", 500)
	}
	
	


    var j json_struct   
    
    // Парсим JSON, если ошибка или другой формат данных - 400 ошибка
    err = json.NewDecoder(r.Body).Decode(&j)
        if err != nil {
            http.Error(w, err.Error(), 400)
            return
        }
    defer r.Body.Close()

    log.Output(1, "Закончили парсить джейсон")
    // Инициализируем элементы xlsx файла для записи
    var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell

	// Создаем файл
	file = xlsx.NewFile()

	// Добавляем страницу
	sheet, err = file.AddSheet(j.Data.Page)
	if err != nil {
	    fmt.Printf(err.Error())
	}

	// Добавляем строку
	row = sheet.AddRow()

	// Записываем заголовки
    for i := range j.Data.Header {
    	cell = row.AddCell()
		cell.Value = j.Data.Header[i]
    }

    // Переходим на новую строку
    row = sheet.AddRow()


    // Записываем основние данные и переходим в конце на новую строку
    go addRows(j.Data.Data, cell, row, sheet)
    
    // Сохраняем файл
    os.MkdirAll(config.FolderPath, 0777);
    err = file.Save(config.FolderPath + "/" + j.Data.Filename)
    log.Output(1, "Сохранили файл")
	if err != nil {
	    fmt.Printf(err.Error())
	}


}


