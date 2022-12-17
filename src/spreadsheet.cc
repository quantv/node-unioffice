#include <node_api.h>
#include <assert.h>
#include <cstdlib>
#include <cstdio>
#include <errno.h>
#include "../include/libspreadsheet.h"

Handle get_handle(napi_env env, napi_value val){
    uint64_t value0;
    napi_status status;
    bool lossless;
    status = napi_get_value_bigint_uint64(env, val, &value0, &lossless);
    assert(status == napi_ok);
    return Handle(value0);
}

char* get_string(napi_env env, napi_value val){
    size_t      len;
    size_t      len2;
    char       *buff;
    napi_get_value_string_utf8(env, val, NULL, 0, &len);
    buff = (char *) malloc (len + 1);
    if (!buff){
        napi_throw_type_error(env, NULL, "Generic error");
        return NULL;
    }
    napi_get_value_string_utf8(env, val, buff, len + 1, &len2);
    return buff;
}
const char* ss_status_code(ss_status s){
    switch (s)
    {
        case ss_ok: return "ss_ok"; break;
        case ss_worksheet_error: return "ss_worksheet_error"; break;
        case ss_save_failed: return "ss_save_failed"; break;
        default:
            return "ss_unknown";
            break;
    }
}

static napi_value New(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    status = napi_create_bigint_uint64(env, uint64_t(ss_new()), &r);
    assert(status == napi_ok);
    return r;
}
static napi_value Open(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;

    size_t argc = 1;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);
    if (argc < 1) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    char *file = get_string(env, args[0]);
    uint64_t hl = uint64_t(ss_open(file));
    delete file;
    if (hl == 0) {
        napi_throw_error(env, "FILE_NOT_FOUND", "no such file or directory");
        return NULL;
    }
    status = napi_create_bigint_uint64(env, hl, &r);
    assert(status == napi_ok);
    return r;
}
static napi_value Close(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 1;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);
    if (argc < 1) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    ss_close(wb);

    return NULL;
}
static napi_value Save(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 2;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 2) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *filepath = get_string(env, args[1]);

    ss_status rs = ss_save(wb, filepath);
    delete filepath;
    if (rs != ss_ok){
        napi_throw_error(env, ss_status_code(rs), "Save failed");
        return NULL;
    }
    return NULL;
}
static napi_value SavePdf(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 3;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 3) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet = get_string(env, args[1]);
    char *filepath = get_string(env, args[2]);

    ss_status rs = ss_save_pdf(wb, sheet, filepath);

    delete filepath;
    delete sheet;

    if (rs != ss_ok){
        napi_throw_error(env, ss_status_code(rs), "Save failed");
        return NULL;
    }
    return NULL;
}
static napi_value AddSheet(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
    size_t argc = 1;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 1) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);

    char *s = ss_add_sheet(wb);

    status = napi_create_string_utf8(env, s, NAPI_AUTO_LENGTH, &r);
    
    delete s;
    
    return r;
}
static napi_value AddRow(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
    size_t argc = 2;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 2) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);

    uint32_t row;
    ss_status rs = ss_add_row(wb, sheet_name, &row);
    
    delete sheet_name;

    if (rs != ss_ok){
        napi_throw_error(env, ss_status_code(rs), "Add row failed");
        return NULL;
    }
    napi_create_uint32(env, row, &r);

    return r;
}
static napi_value AddRows(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 3;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 3) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    int32_t count;
    status = napi_get_value_int32(env, args[2], &count);
    assert(status == napi_ok);
    ss_add_rows(wb, sheet_name, count);

    delete sheet_name;

    return args[2];
}

static napi_value InsertRows(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    int32_t rowNum;
    int32_t count;

    status = napi_get_value_int32(env, args[2], &rowNum);
    assert(status == napi_ok);

    status = napi_get_value_int32(env, args[3], &count);
    assert(status == napi_ok);

    ss_status rs = ss_insert_rows(wb, sheet_name, rowNum, count);

    delete sheet_name;

    if (rs != ss_ok){
        napi_throw_error(env, ss_status_code(rs), "Error insert rows");
        return NULL;
    }

    return args[2];
}

static napi_value CopyRows(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 5;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 5) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    int32_t source;
    int32_t count;
    int32_t dest;

    status = napi_get_value_int32(env, args[2], &source);
    assert(status == napi_ok);

    status = napi_get_value_int32(env, args[3], &dest);
    assert(status == napi_ok);

    status = napi_get_value_int32(env, args[4], &count);
    assert(status == napi_ok);

    ss_status rs = ss_copy_rows(wb, sheet_name, source, dest, count);

    delete sheet_name;

    if (rs != ss_ok){
        napi_throw_error(env, ss_status_code(rs), "Error copy rows");
        return NULL;
    }

    return args[2];
}

static napi_value AddCell(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
    size_t argc = 3;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 2) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    uint64_t value0;
    bool lossless;
    status = napi_get_value_bigint_uint64(env, args[0], &value0, &lossless);
    assert(status == napi_ok);

    char *sheet_name = get_string(env, args[1]);

    uint32_t row;
    napi_get_value_uint32(env, args[2], &row);

    char *cell = ss_add_cell(Handle(value0), sheet_name, row);

    napi_create_string_utf8(env, cell, NAPI_AUTO_LENGTH, &r);

    delete sheet_name;
    delete cell;

    return r;
}


static napi_value CheckSheet(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
    size_t argc = 2;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 2) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);

    ss_status rs = ss_check_sheet(wb, sheet_name);
    delete sheet_name;

    if (rs != ss_ok){
        napi_throw_error(env, ss_status_code(rs), "Sheet not exist");
        return NULL;
    }
    return r;
}

static napi_value SetCellString(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    char *val = get_string(env, args[3]);

    int32_t rs = ss_set_cell_string(wb, sheet_name, cell_name, val);

    delete sheet_name;
    delete cell_name;
    delete val;

    assert(rs == 0);

    return NULL;
}

static napi_value SetCellBool(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    bool val;
    
    status = napi_get_value_bool(env, args[3], &val);
    assert(status == napi_ok);

    int32_t rs = ss_set_cell_bool(wb, sheet_name, cell_name, val ? 1 : 0);

    delete sheet_name;
    delete cell_name;

    assert(rs == 0);

    return NULL;
}

#define DECLARE_NAPI_METHOD(name, func) {name, 0, func, 0, 0, 0, napi_default, 0}

static napi_value SetCellDate(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    double val;
    
    status = napi_get_date_value(env, args[3], &val);
    assert(status == napi_ok);

    int32_t rs = ss_set_cell_date(wb, sheet_name, cell_name, val);

    delete sheet_name;
    delete cell_name;

    assert(rs == 0);

    return NULL;
}

static napi_value SetCellNumber(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    double val;
    
    status = napi_get_value_double(env, args[3], &val);
    assert(status == napi_ok);

    int32_t rs = ss_set_cell_number(wb, sheet_name, cell_name, val);

    delete sheet_name;
    delete cell_name;

    assert(rs == 0);

    return NULL;
}

static napi_value SetCellFormulaRaw(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    char *val = get_string(env, args[3]);

    int32_t rs = ss_set_cell_formula_raw(wb, sheet_name, cell_name, val);

    delete sheet_name;
    delete cell_name;
    delete val;

    assert(rs == 0);

    return NULL;
}

static napi_value SetCellFormulaArray(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    char *val = get_string(env, args[3]);

    int32_t rs = ss_set_cell_formula_array(wb, sheet_name, cell_name, val);

    delete sheet_name;
    delete cell_name;
    delete val;

    assert(rs == 0);

    return NULL;
}

static napi_value SetCellFormulaShared(napi_env env, napi_callback_info info){
    napi_status status;
    
    size_t argc = 6;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 6) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    char *val = get_string(env, args[3]);
    u_int32_t rows, cols;
    status = napi_get_value_uint32(env, args[4], &rows);
    assert(status == napi_ok);
    status = napi_get_value_uint32(env, args[5], &cols);
    assert(status == napi_ok);

    int32_t rs = ss_set_cell_formula_shared(wb, sheet_name, cell_name, val, rows, cols);

    delete sheet_name;
    delete cell_name;
    delete val;

    assert(rs == 0);

    return NULL;
}

enum cell_value_type {
    cell_value_type_string=1, 
    cell_value_type_number=2,
    cell_value_type_date=3,
    cell_value_type_bool=4,
};

static napi_value CellGetValue(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value value;

    size_t argc = 4;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 4) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    char *cell_name = get_string(env, args[2]);
    
    int32_t t;
    status = napi_get_value_int32(env, args[3], &t);
    assert(status == napi_ok);

    double dv;
    int32_t bv;

    errno = 0;
    switch (t)
    {
        case cell_value_type_number:
            dv = ss_cell_get_as_number(wb, sheet_name, cell_name);
            status = napi_create_double(env, dv, &value);
            break;
        case cell_value_type_date:
            dv = ss_cell_get_date(wb, sheet_name, cell_name);
            status = napi_create_date(env, dv, &value);
            break;
        case cell_value_type_bool:
            bv = ss_cell_get_bool(wb, sheet_name, cell_name);
            status = napi_create_int32(env, bv, &value);
        break;
        case cell_value_type_string:
        default:
            char *sv = ss_cell_get_as_string(wb, sheet_name, cell_name);
            status = napi_create_string_utf8(env, sv, NAPI_AUTO_LENGTH, &value);
            delete sv;
            break;
    }
    
    delete sheet_name;
    delete cell_name;

    if (errno != 0){
        napi_throw_error(env, NULL, "Error getting cell value");
        return NULL;
    }
    assert(status == napi_ok);
    return value;
}

static napi_value RecalculateFormulas(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value value;

    size_t argc = 2;
    napi_value args[argc];
    status = napi_get_cb_info(env, info, &argc, args, NULL, NULL);
    assert(status == napi_ok);

    if (argc < 2) {
        napi_throw_type_error(env, NULL, "Wrong number of arguments");
        return NULL;
    }
    
    Handle wb = get_handle(env, args[0]);
    char *sheet_name = get_string(env, args[1]);
    
    ss_recalculate_formulas(wb, sheet_name);

    delete sheet_name;
    return NULL;
}

static napi_value Init(napi_env env, napi_value exports){
    napi_status status;
    napi_property_descriptor desc[] = {
        DECLARE_NAPI_METHOD("cell_get_value", CellGetValue),
        //DECLARE_NAPI_METHOD("sheet_get_row_values", GetRowValues),
        //set cell values
        DECLARE_NAPI_METHOD("set_cell_date", SetCellDate),
        DECLARE_NAPI_METHOD("set_cell_formula_array", SetCellFormulaArray),
        DECLARE_NAPI_METHOD("set_cell_formula_shared", SetCellFormulaShared),
        DECLARE_NAPI_METHOD("set_cell_formula_raw", SetCellFormulaRaw),
        DECLARE_NAPI_METHOD("set_cell_number", SetCellNumber),
        DECLARE_NAPI_METHOD("set_cell_string", SetCellString),
        DECLARE_NAPI_METHOD("set_cell_bool", SetCellBool),
        //workbooks
        DECLARE_NAPI_METHOD("new", New),
        DECLARE_NAPI_METHOD("open", Open),
        DECLARE_NAPI_METHOD("save", Save),
        DECLARE_NAPI_METHOD("save_pdf", SavePdf),
        DECLARE_NAPI_METHOD("sheet_recalculate_formulas", RecalculateFormulas),
        DECLARE_NAPI_METHOD("close", Close),
        DECLARE_NAPI_METHOD("add_sheet", AddSheet),
        DECLARE_NAPI_METHOD("add_row", AddRow),
        DECLARE_NAPI_METHOD("add_rows", AddRows),
        DECLARE_NAPI_METHOD("add_cell", AddCell),
        DECLARE_NAPI_METHOD("check_sheet", CheckSheet),
        DECLARE_NAPI_METHOD("insert_rows", InsertRows),
        DECLARE_NAPI_METHOD("copy_rows", CopyRows),
    };
    status = napi_define_properties(env, exports, sizeof(desc) / sizeof(*desc), desc);
    assert(status == napi_ok);
    return exports;
}

NAPI_MODULE(NODE_GYP_MODULE_NAME, Init);