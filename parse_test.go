package main

import (
	"fmt"
	"io/ioutil"
	"testing"
)

func TestParseCsv(t *testing.T) {
	b, err := ioutil.ReadFile("tmp/华夏银企对账单 顺丰套打 冀州.csv")
	if err != nil {
		t.Fatalf("%s\n", err)
	}
	hx, err := ParseCsv(string(b), 16)
	if err != nil {
		t.Fatalf("%s\n", err)
	}
	for _, v := range hx {
		fmt.Printf("%s %d\n", v.BarCode, len(v.BankRecipts))
	}
}
