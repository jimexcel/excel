package excel

import (
	"errors"
	"io/ioutil"

	"path/filepath"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//获取指定目录下所有指定ext后缀的文件（如.xlsx）的文件
func GetDirAllFiles(dirPath, ext string) ([]string, error) {
	filenames := []string{}
	rd, err := ioutil.ReadDir(dirPath)
	if err != nil {
		return []string{}, err
	}
	for _, fi := range rd {
		if fi.IsDir() {
			s, err := GetDirAllFiles(dirPath+fi.Name()+"/", ext)
			if err != nil {
				return []string{}, err
			}
			filenames = append(filenames, s...)
		} else {
			if filepath.Ext(dirPath+"/"+fi.Name()) == ext {
				filenames = append(filenames, dirPath+"/"+fi.Name())
			}
		}
	}
	return filenames, nil
}

//将excel读取到二位数组，读取默认表格
func ReadToArray(path string) ([][]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return [][]string{}, err
	}
	//获取默认工作表
	index := f.GetActiveSheetIndex()
	if index == 0 {
		return [][]string{}, errors.New("Active sheet not found")
	}
	rows := f.GetRows(f.GetSheetName(index))
	return rows, nil
}

//将excel写入二维数组，默认表格名Sheet1
func WriteArray(datas [][]string, path string) error {
	f := excelize.NewFile()
	sheetName := "Sheet1"
	index := f.NewSheet(sheetName)
	for i := 0; i < len(datas); i++ {
		for j := 0; j < len(datas[i]); j++ {
			cellIndex := string(rune(65+j)) + strconv.Itoa(i+1) //start from A1
			f.SetCellStr(sheetName, cellIndex, datas[i][j])
		}
	}
	f.SetActiveSheet(index)
	err := f.SaveAs(path)
	if err != nil {
		return err
	}
	return nil
}
