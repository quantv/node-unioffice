#include <node_api.h>
#include <assert.h>
#include <cstdlib>
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

static napi_value New(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    status = napi_create_bigint_uint64(env, uint64_t(ss_new()), &r);
    assert(status == napi_ok);
    return r;
}
static napi_value Open(napi_env env, napi_callback_info info){

}
static napi_value Close(napi_env env, napi_callback_info info){
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
    ss_close(wb);

    return NULL;
}
static napi_value Save(napi_env env, napi_callback_info info){
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
    char *filepath = get_string(env, args[1]);

    ss_save(wb, filepath);
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

    uint32_t row = ss_add_row(wb, sheet_name);

    napi_create_uint32(env, row, &r);
    return r;
}
static napi_value AddRows(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
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
    return r;
}
static napi_value SetCellString(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
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

    int32_t row = ss_set_cell_string(wb, sheet_name, cell_name, val);

    napi_create_int32(env, row, &r);
    return r;
}

static napi_value TestWrite(napi_env env, napi_callback_info info){
    napi_status status;
    napi_value r;
    
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
    char *val = get_string(env, args[2]);
    int32_t count;
    status = napi_get_value_int32(env, args[3], &count);
    assert(status == napi_ok);

    test_write_multi(wb, sheet_name, count, val);

    return NULL;
}

static napi_value copyRows(napi_env env, napi_callback_info info){

}


#define DECLARE_NAPI_METHOD(name, func) {name, 0, func, 0, 0, 0, napi_default, 0}

static napi_value Init(napi_env env, napi_value exports){
    napi_status status;
    napi_property_descriptor desc[] = {
        DECLARE_NAPI_METHOD("new", New),
        DECLARE_NAPI_METHOD("open", Open),
        DECLARE_NAPI_METHOD("save", Save),
        DECLARE_NAPI_METHOD("close", Close),
        DECLARE_NAPI_METHOD("add_sheet", AddSheet),
        DECLARE_NAPI_METHOD("add_row", AddRow),
        DECLARE_NAPI_METHOD("add_rows", AddRows),
        DECLARE_NAPI_METHOD("add_cell", AddCell),
        DECLARE_NAPI_METHOD("set_cell_string", SetCellString),
        DECLARE_NAPI_METHOD("test", TestWrite),
    };
    status = napi_define_properties(env, exports, sizeof(desc) / sizeof(*desc), desc);
    assert(status == napi_ok);
    return exports;
}

NAPI_MODULE(NODE_GYP_MODULE_NAME, Init);