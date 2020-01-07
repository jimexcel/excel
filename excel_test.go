package excel

import (
	"testing"
)

func TestRead(t *testing.T) {
	datas, err := ReadToArray("datas/testread.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	t.Log(datas)
}

func TestWrite(t *testing.T) {
	datas, err := ReadToArray("datas/testread.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	err = WriteArray(datas, "datas/dst.xlsx")
	if err != nil {
		t.Fatal(err)
	}
}

func TestReadDir(t *testing.T) {
	datas, err := GetDirAllFiles("datas", ".xlsx")
	if err != nil {
		t.Fatal(err)
	}
	t.Log(datas)
}
