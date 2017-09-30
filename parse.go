package main

import (
	"errors"
	"fmt"
	"sort"
	"strings"
)

var (
	textErr  = "the text is invalid"
	splitErr = "split text failed, length of slice not 16"
)

func ParseCsv(s string, n int) (map[string]*HuaxiaBank, error) {
	lines := strings.Split(s, "\n")
	if len(lines) < 1 {
		return nil, errors.New("parse string " + textErr)
	}

	ss := make([][]string, 0)
	for i, line := range lines {
		index := strings.Index(line, "|")
		if index == -1 {
			continue
		}
		st := strings.Split(line, "|")
		if len(st) != n {
			return nil, errors.New(fmt.Sprintf("line: %d, %s", i, splitErr))
		}
		ss = append(ss, st)
	}
	sort.Slice(ss, func(i, j int) bool { return ss[i][1] < ss[j][1] })
	hx := make(map[string]*HuaxiaBank)
	for i := 0; i < len(ss); i++ {
		barcode := ss[i][1]
		t := ss[i]
		stm := &Statement{
			ID:             t[0],
			BankCode:       t[3],
			Account:        t[7],
			AccountType:    t[11],
			Currency:       t[6],
			DepositBank:    t[4],
			AccountBalance: t[8],
			Pledgedbalance: t[9],
		}
		if bank, ok := hx[barcode]; ok {
			bank.BankRecipts = append(bank.BankRecipts, stm)
			continue
		}
		h := &HuaxiaBank{
			BarCode:       t[1],
			PostalCode:    t[13],
			ClientCode:    t[5],
			CompanyName:   t[10],
			DepositBank:   t[4],
			ClientAddress: t[12],
			Contact:       t[15],
			PhoneNumber:   t[14],
			DueDate:       t[2],
			BankRecipts:   []*Statement{stm},
		}
		hx[barcode] = h
	}
	return hx, nil
}
