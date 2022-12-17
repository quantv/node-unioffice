/* Code generated by cmd/cgo; DO NOT EDIT. */

/* package command-line-arguments */


#line 1 "cgo-builtin-export-prolog"

#include <stddef.h>

#ifndef GO_CGO_EXPORT_PROLOGUE_H
#define GO_CGO_EXPORT_PROLOGUE_H

#ifndef GO_CGO_GOSTRING_TYPEDEF
typedef struct { const char *p; ptrdiff_t n; } _GoString_;
#endif

#endif

/* Start of preamble from import "C" comments.  */


#line 3 "lib.go"

#include <stdint.h>
#include <errno.h>
typedef uint64_t Handle;

typedef enum {
  ss_ok,
  ss_worksheet_error,
  ss_save_failed,
  ss_not_a_number,
} ss_status;

typedef struct {
	char* v;
	uint8_t t;
} cellValue;

#line 1 "cgo-generated-wrapper"


/* End of preamble from import "C" comments.  */


/* Start of boilerplate cgo prologue.  */
#line 1 "cgo-gcc-export-header-prolog"

#ifndef GO_CGO_PROLOGUE_H
#define GO_CGO_PROLOGUE_H

typedef signed char GoInt8;
typedef unsigned char GoUint8;
typedef short GoInt16;
typedef unsigned short GoUint16;
typedef int GoInt32;
typedef unsigned int GoUint32;
typedef long long GoInt64;
typedef unsigned long long GoUint64;
typedef GoInt64 GoInt;
typedef GoUint64 GoUint;
typedef size_t GoUintptr;
typedef float GoFloat32;
typedef double GoFloat64;
#ifdef _MSC_VER
#include <complex.h>
typedef _Fcomplex GoComplex64;
typedef _Dcomplex GoComplex128;
#else
typedef float _Complex GoComplex64;
typedef double _Complex GoComplex128;
#endif

/*
  static assertion to make sure the file is being used on architecture
  at least with matching size of GoInt.
*/
typedef char _check_for_64_bit_pointer_matching_GoInt[sizeof(void*)==64/8 ? 1:-1];

#ifndef GO_CGO_GOSTRING_TYPEDEF
typedef _GoString_ GoString;
#endif
typedef void *GoMap;
typedef void *GoChan;
typedef struct { void *t; void *v; } GoInterface;
typedef struct { void *data; GoInt len; GoInt cap; } GoSlice;

#endif

/* End of boilerplate cgo prologue.  */

#ifdef __cplusplus
extern "C" {
#endif

extern Handle ss_new();
extern Handle ss_open(char* filepath);
extern char* ss_add_sheet(Handle h);
extern ss_status ss_add_row(Handle h, char* sheet, uint32_t* row);
extern int32_t ss_add_rows(Handle h, char* sheet, int32_t count);
extern ss_status ss_insert_rows(Handle h, char* sheetName, int32_t rowNum, int32_t rows);
extern ss_status ss_copy_rows(Handle h, char* sheetName, int32_t source, int32_t dest, int32_t rows);
extern char* ss_add_cell(Handle h, char* sheet, uint32_t row);
extern void ss_close(Handle h);
extern ss_status ss_save(Handle ws, char* filepath);
extern ss_status ss_save_pdf(Handle ws, char* sheet, char* dest);
extern ss_status ss_check_sheet(Handle h, char* sheet);
extern int32_t ss_set_cell_string(Handle h, char* sheet, char* cell, char* value);
extern int32_t ss_set_cell_bool(Handle h, char* sheet, char* cell, uint8_t value);
extern int32_t ss_set_cell_date(Handle h, char* sheet, char* cell, double value);
extern int32_t ss_set_cell_date_with_style(Handle h, char* sheet, char* cell, double value);
extern int32_t ss_set_cell_formula_array(Handle h, char* sheet, char* cell, char* value);
extern int32_t ss_set_cell_formula_raw(Handle h, char* sheet, char* cell, char* value);
extern int32_t ss_set_cell_formula_shared(Handle h, char* sheet, char* cell, char* value, uint32_t rows, uint32_t cols);
extern int32_t ss_set_cell_number(Handle h, char* sheet, char* cell, double value);
extern cellValue ss_cell_get_value(Handle h, char* sheet, char* cell);
extern char* ss_cell_get_as_string(Handle h, char* sheet, char* cell);
extern double ss_cell_get_as_number(Handle h, char* sheet, char* cell);
extern uint8_t ss_cell_get_bool(Handle h, char* sheet, char* cell);
extern int64_t ss_cell_get_date(Handle h, char* sheet, char* cell);
extern void ss_recalculate_formulas(Handle h, char* sheet);
extern int32_t test_write_multi(Handle h, char* sheet, int32_t count, char* value);

#ifdef __cplusplus
}
#endif
