// Copyright 2017 FoxyUtils ehf. All rights reserved.
package main

/*
#include <stdint.h>
typedef uint64_t Handle;

typedef enum {
  ss_ok,
  ss_worksheet_error,
  ss_save_failed
} ss_status;
*/
import "C"

import (
	"strconv"
	"unsafe"

	"github.com/unidoc/unioffice/spreadsheet"
	"github.com/unidoc/unioffice/spreadsheet/convert"
)

//export ss_new
func ss_new() C.Handle {
	ss := spreadsheet.New()
	handle := C.Handle(uintptr(unsafe.Pointer(&ss)))
	workbooks[handle] = ss
	return handle
}

//export ss_open
func ss_open(filepath *C.char) C.Handle {
	file := C.GoString(filepath)
	ss, err := spreadsheet.Open(file)
	if err != nil {
		return 0
	}
	handle := C.Handle(uintptr(unsafe.Pointer(&ss)))
	workbooks[handle] = ss
	return handle
}

//export ss_add_sheet
func ss_add_sheet(h C.Handle) *C.char {
	ss := workbooks[h]
	sh := ss.AddSheet()
	return C.CString(sh.Name())
}

//export ss_add_row
func ss_add_row(h C.Handle, sheet *C.char, row *C.uint32_t) C.ss_status {
	sh, err := get_sheet(h, sheet)
	if err != nil {
		return C.ss_worksheet_error
	}
	r := sh.AddRow()
	*row = C.uint32_t(r.RowNumber())

	return C.ss_ok
}

//export ss_add_rows
func ss_add_rows(h C.Handle, sheet *C.char, count C.int32_t) C.int32_t {
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	sh, _ := ss.GetSheet(sheet_name)
	for i := 1; i < int(count); i++ {
		sh.AddRow()
	}
	return count
}

//export ss_insert_rows
func ss_insert_rows(h C.Handle, sheetName *C.char, rowNum, rows C.int32_t) C.ss_status {
	sheet, err := get_sheet(h, sheetName)
	if err != nil {
		return C.ss_worksheet_error
	}
	sheet.InsertRows(int(rowNum), uint32(rows))
	return C.ss_ok
}

//export ss_copy_rows
func ss_copy_rows(h C.Handle, sheetName *C.char, source, dest, rows C.int32_t) C.ss_status {
	sheet, err := get_sheet(h, sheetName)
	if err != nil {
		return C.ss_worksheet_error
	}
	sheet.CopyRows(uint32(source), uint32(dest), int(rows))
	return C.ss_ok
}

//export ss_add_cell
func ss_add_cell(h C.Handle, sheet *C.char, row C.uint32_t) *C.char {
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	sh, _ := ss.GetSheet(sheet_name)
	r := sh.Row(uint32(row))
	cell := r.AddCell()
	return C.CString(cell.Reference())
}

//export ss_close
func ss_close(h C.Handle) {
	ss := workbooks[h]
	ss.Close()
	delete(workbooks, h)
}

//export ss_save
func ss_save(ws C.Handle, filepath *C.char) C.ss_status {
	wb := workbooks[ws]
	defer wb.Close()
	defer delete(workbooks, ws)
	err := wb.SaveToFile(C.GoString(filepath))
	if err != nil {
		return C.ss_save_failed
	}
	return C.ss_ok
}

//export ss_save_pdf
func ss_save_pdf(ws C.Handle, sheet, dest *C.char) C.ss_status {
	wb := workbooks[ws]
	defer wb.Close()
	defer delete(workbooks, ws)
	sheet_name := C.GoString(sheet)
	sh, err := wb.GetSheet(sheet_name)
	if err != nil {
		return C.ss_worksheet_error
	}
	pdf := convert.ConvertToPdf(&sh)
	err = pdf.WriteToFile(C.GoString(dest))
	if err != nil {
		return C.ss_save_failed
	}
	return C.ss_ok
}

//export ss_check_sheet
func ss_check_sheet(h C.Handle, sheet *C.char) C.ss_status {
	_, err := get_sheet(h, sheet)
	if err != nil {
		return C.ss_worksheet_error
	}
	return C.ss_ok
}

func get_sheet(h C.Handle, sheet *C.char) (spreadsheet.Sheet, error) {
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	return ss.GetSheet(sheet_name)
}

func get_cell(h C.Handle, sheet *C.char, cell *C.char) (spreadsheet.Cell, error) {
	sh, err := get_sheet(h, sheet)
	var c spreadsheet.Cell
	if err == nil {
		c = sh.Cell(C.GoString(cell))
	}
	return c, err
}

//export ss_set_cell_string
func ss_set_cell_string(h C.Handle, sheet *C.char, cell, value *C.char) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetString(C.GoString(value))
	return 0
}

//export ss_set_cell_bool
func ss_set_cell_bool(h C.Handle, sheet *C.char, cell *C.char, value C.uint8_t) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetBool(value == 1)
	return 0
}

// func (c Cell) SetBool(v bool)
// func (c Cell) SetDate(d time.Time)
// func (c Cell) SetDateWithStyle(d time.Time)
// func (c Cell) SetFormulaArray(s string)
// func (c Cell) SetFormulaRaw(s string)
// func (c Cell) SetFormulaShared(formulaStr string, rows, cols uint32) error
// func (c Cell) SetHyperlink(hl common.Hyperlink)
// func (c Cell) SetNumber(v float64)
// func (c Cell) SetRichTextString() RichText

func ss_get_cell_string() {}

//export test_write_multi
func test_write_multi(h C.Handle, sheet *C.char, count C.int32_t, value *C.char) C.int32_t {
	c := int(count)
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	sh, err := ss.GetSheet(sheet_name)
	if err != nil {
		panic(err)
	}
	for i := 0; i < c; i++ {
		sh.Cell("A" + strconv.Itoa(i+1)).SetString(C.GoString(value))
	}
	return C.int32_t(0)
}

var workbooks = make(map[C.Handle]*spreadsheet.Workbook)

func main() {
}
