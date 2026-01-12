/*
PROMPTS USED:
in gemini/xlsximport.c create C code for a SQLite extension named xlsximport. This xlsximport code uses the SQLite extension zipfile to open a XLSX file and gather this content:
     xl/sharedStrings.xml
     xl/worksheets/sheet1.xml to  xl/worksheets/sheetN.xml
     xl/workbook.xml
The name of each sheet is in xl/workbook.xml
The individual sheets are kept in xl/worksheets/sheet1.xml  to  xl/worksheets/sheetN.xml
To save on space, Microsoft stores all the character literal values in one common xl/sharedStrings.xml dictionary file. 
The individual cell value found for this string in the actual sheet1.xml file is just an index into this dictionary.
Microsoft does not store empty cells or rows in xl worksheets sheet1.xml, so any gaps between values have to be taken care by the code.
To figure out the number of skipped columns one need to be able to figure out the distance between, say, cell "AB67" and "C67". The way columns are named: A through Z, then AA through AZ, then AAA through AAZ, etc., suggests that we may assume they are using a base-26 system and therefore use a simple conversion method from a base-26 to the decimal system and then use subtraction to find out the number of empty cells between columns.
Create a SQL function named xlsx_import that creates one table for each of the sheets in the XLSX files, table name equal to sheet name, and column names equal to the values in first row of the sheet.
xlsx_import returns the number of tables imported.
The first parameter is the XLSX filename. Subsequent optional parameters are sheet names or sheet numbers (1-based) to import.
Add table valued SQL function named xlsx_import_sheetnames with parameter XLSX filename that returns the names of the sheets in the file,
declared as "CREATE TABLE x(sheet_num INTEGER, sheet_name TEXT, filename HIDDEN)"
Use expat for XML parsing. Add support for both shared strings and inline strings.
Support sparsely populated spreadsheets.
Add SQL function xlsx_import_version returning "2026-01-07 Gemini 3 Pro (High)".
Add all user prompts as comments.
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1
#include <expat.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

/* String Buffer Helper */
typedef struct {
    char *data;
    size_t len;
    size_t cap;
} StrBuf;

static void strbuf_init(StrBuf *sb) {
    sb->data = NULL;
    sb->len = 0;
    sb->cap = 0;
}

static void strbuf_append(StrBuf *sb, const char *s, int len) {
    if (len < 0) len = (int)strlen(s);
    if (sb->len + len + 1 > sb->cap) {
        size_t new_cap = sb->cap ? sb->cap * 2 : 1024;
        while (new_cap < sb->len + len + 1) new_cap *= 2;
        sb->data = sqlite3_realloc(sb->data, new_cap);
        sb->cap = new_cap;
    }
    memcpy(sb->data + sb->len, s, len);
    sb->len += len;
    sb->data[sb->len] = '\0';
}

static void strbuf_free(StrBuf *sb) {
    sqlite3_free(sb->data);
    strbuf_init(sb);
}

/* Shared Strings */
typedef struct {
    char **strings;
    int count;
    int cap;
    StrBuf current_str;
    int in_t_tag; 
} SharedStrings;

static void shared_strings_start(void *userData, const char *name, const char **atts) {
    SharedStrings *ss = (SharedStrings *)userData;
    (void)atts;
    if (strcmp(name, "t") == 0) {
        ss->in_t_tag = 1;
        strbuf_init(&ss->current_str);
    }
}

static void shared_strings_end(void *userData, const char *name) {
    SharedStrings *ss = (SharedStrings *)userData;
    if (strcmp(name, "t") == 0) {
        ss->in_t_tag = 0;
        if (ss->count == ss->cap) {
            ss->cap = ss->cap ? ss->cap * 2 : 1024;
            ss->strings = sqlite3_realloc(ss->strings, ss->cap * sizeof(char *));
        }
        ss->strings[ss->count++] = ss->current_str.data; /* takes ownership */
    }
}

static void shared_strings_char(void *userData, const XML_Char *s, int len) {
    SharedStrings *ss = (SharedStrings *)userData;
    if (ss->in_t_tag) {
        strbuf_append(&ss->current_str, s, len);
    }
}

/* Base-26 Column Conversion */
static int col_to_int(const char *col) {
    int res = 0;
    while (*col && isalpha(*col)) {
        res = res * 26 + (toupper(*col) - 'A' + 1);
        col++;
    }
    return res;
}

static void int_to_col(int n, char *buf) {
    char temp[16];
    int i = 0;
    n++; /* 1-based */
    while (n > 0) {
        n--;
        temp[i++] = (char)('A' + (n % 26));
        n /= 26;
    }
    int j;
    for (j = 0; j < i; j++) {
        buf[j] = temp[i - 1 - j];
    }
    buf[i] = '\0';
}

static void extract_col_row(const char *ref, int *col, int *row) {
    const char *p = ref;
    while (*p && isalpha(*p)) p++;
    size_t len = p - ref;
    char c_part[16];
    if (len >= sizeof(c_part)) len = sizeof(c_part) - 1;
    memcpy(c_part, ref, len);
    c_part[len] = '\0';
    *col = col_to_int(c_part);
    *row = atoi(p);
}

/* Workbook / Sheets Info */
typedef struct {
    char *name;
    int id; /* r:id or sheetId */
} SheetInfo;

typedef struct {
    SheetInfo *sheets;
    int count;
    int cap;
} Workbook;

static void workbook_free(Workbook *wb) {
    int i;
    for (i = 0; i < wb->count; i++) sqlite3_free(wb->sheets[i].name);
    sqlite3_free(wb->sheets);
    wb->sheets = NULL;
    wb->count = 0;
    wb->cap = 0;
}

static void workbook_start(void *userData, const char *name, const char **atts) {
    Workbook *wb = (Workbook *)userData;
    if (strcmp(name, "sheet") == 0) {
        const char *sheet_name = NULL;
        /* const char *sheet_id_str = NULL; */
        int i;
        for (i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "name") == 0) sheet_name = atts[i+1];
            /* if (strcmp(atts[i], "sheetId") == 0) sheet_id_str = atts[i+1]; */
        }
        
        if (wb->count == wb->cap) {
            wb->cap = wb->cap ? wb->cap * 2 : 8;
            wb->sheets = sqlite3_realloc(wb->sheets, wb->cap * sizeof(SheetInfo));
        }
        wb->sheets[wb->count].name = sqlite3_mprintf("%s", sheet_name);
        wb->sheets[wb->count].id = wb->count + 1; /* Assume 1-based index maps to sheetN.xml */
        wb->count++;
    }
}

/* Parsing Sheet Data */
typedef struct {
    sqlite3 *db;
    char *table_name;
    SharedStrings *ss;
    
    int current_row;
    int current_col;
    
    /* Cell state */
    int in_row;
    int in_c;
    int in_v;
    int in_t; /* inlineStr value */
    char col_ref[16];
    char cell_type[16]; /* "s", "inlineStr", etc. */
    StrBuf cell_val;
    
    /* SQL generation */
    StrBuf sql_buf;
    int header_row_processed;
    StrBuf header_cols; /* CSV of column names */
    int num_cols;
    int dimension_max_col;
    int dimension_max_row;
    int last_row_inserted;
    int table_created;
} SheetCtx;

static void sheet_start(void *userData, const char *name, const char **atts) {
    SheetCtx *ctx = (SheetCtx *)userData;
    int i;

    if (strcmp(name, "dimension") == 0) {
        const char *ref = NULL;
        for (i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "ref") == 0) ref = atts[i+1];
        }
        if (ref) {
           /* ref is "A1" or "A1:C5" */
           const char *p = strchr(ref, ':');
           if (p) {
               int c, r;
               extract_col_row(p + 1, &c, &r);
               ctx->dimension_max_col = c;
               ctx->dimension_max_row = r;
           } else {
               int c, r;
               extract_col_row(ref, &c, &r);
               ctx->dimension_max_col = c;
               ctx->dimension_max_row = r;
           }
        }
    } else if (strcmp(name, "row") == 0) {
        ctx->in_row = 1;
        const char *r_attr = NULL;
        for (i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "r") == 0) r_attr = atts[i+1];
        }
        if (r_attr) ctx->current_row = atoi(r_attr);
        ctx->current_col = 0; /* Convert 1..N to 0..N-1? No, col_to_int returns 1-based */
        
        if (ctx->header_row_processed) {
            /* Handle Row Gaps */
            while (ctx->last_row_inserted + 1 < ctx->current_row) {
                ctx->last_row_inserted++;
                strbuf_init(&ctx->sql_buf);
                strbuf_append(&ctx->sql_buf, "INSERT INTO \"", -1);
                strbuf_append(&ctx->sql_buf, ctx->table_name, -1);
                strbuf_append(&ctx->sql_buf, "\" VALUES (", -1);
                int k;
                for (k = 0; k < ctx->num_cols; k++) {
                    if (k > 0) strbuf_append(&ctx->sql_buf, ", ", 2);
                    strbuf_append(&ctx->sql_buf, "NULL", 4);
                }
                strbuf_append(&ctx->sql_buf, ");", 2);
                char *zErr = NULL;
                int rc = sqlite3_exec(ctx->db, ctx->sql_buf.data, NULL, NULL, &zErr);
                if (rc != SQLITE_OK) {
                    fprintf(stderr, "xlsximport: row gap insert error: %s\n", zErr);
                    sqlite3_free(zErr);
                }
                strbuf_free(&ctx->sql_buf);
            }

            strbuf_init(&ctx->sql_buf);
            strbuf_append(&ctx->sql_buf, "INSERT INTO \"", -1);
            strbuf_append(&ctx->sql_buf, ctx->table_name, -1);
            strbuf_append(&ctx->sql_buf, "\" VALUES (", -1);
        }
    } else if (strcmp(name, "c") == 0) {
        ctx->in_c = 1;
        ctx->cell_type[0] = '\0';
        for (i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "r") == 0) {
                int c, r;
                extract_col_row(atts[i+1], &c, &r);
                /* Calculate gap */
                /* If new row, current_col is 0. If data is B2 (col 2), gap is 1 (A). */
                /* But wait, if header row processed, we need to fill gaps with NULL in INSERT */
                if (ctx->header_row_processed) {
                     /* Logic: previous col was `current_col`. New col is `c`. 
                        Gaps = c - current_col - 1. 
                        Example: prev 1 (A), new 3 (C). c-current-1 = 3-1-1 = 1 gap (B).
                     */
                     int gaps = c - ctx->current_col - 1;
                     int k;
                     if (ctx->current_col == 0) gaps = c - 1; /* First cell in row */
                     
                     for (k = 0; k < gaps; k++) {
                         if (ctx->current_col + k + 1 <= ctx->num_cols) {
                             if (ctx->sql_buf.data[ctx->sql_buf.len-1] != '(') {
                                 strbuf_append(&ctx->sql_buf, ", ", 2);
                             }
                             strbuf_append(&ctx->sql_buf, "NULL", 4);
                         }
                     }
                } else {
                     /* Header Row Gap Filling */
                     int gaps = c - ctx->current_col - 1;
                     if (ctx->current_col == 0) gaps = c - 1;
                     
                     int k;
                     for (k = 0; k < gaps; k++) {
                         if (ctx->header_cols.len > 0) strbuf_append(&ctx->header_cols, ", ", 2);
                         char col_name[32];
                         char letter[16];
                         /* Determine column index being filled */
                         int fill_idx = ctx->current_col + k; /* 0-based index of column */
                         int_to_col(fill_idx, letter);
                         snprintf(col_name, sizeof(col_name), "\"%s\"", letter);
                         strbuf_append(&ctx->header_cols, col_name, -1);
                         ctx->num_cols++;
                     }
                }
                ctx->current_col = c;
            }
            if (strcmp(atts[i], "t") == 0) strncpy(ctx->cell_type, atts[i+1], 15);
        }
        strbuf_init(&ctx->cell_val);
    } else if (strcmp(name, "v") == 0 || strcmp(name, "t") == 0) { 
        /* 'v' for value, 't' can be used inside inlineStr */
        if (strcmp(name, "v") == 0) ctx->in_v = 1;
        if (strcmp(name, "t") == 0 && strcmp(ctx->cell_type, "inlineStr") == 0) ctx->in_t = 1;
    }
}

static void sheet_char(void *userData, const XML_Char *s, int len) {
    SheetCtx *ctx = (SheetCtx *)userData;
    if (ctx->in_v || ctx->in_t) {
        strbuf_append(&ctx->cell_val, s, len);
    }
}

static void sheet_end(void *userData, const char *name) {
    SheetCtx *ctx = (SheetCtx *)userData;
    
    if (strcmp(name, "v") == 0) ctx->in_v = 0;
    else if (strcmp(name, "t") == 0) ctx->in_t = 0;
    else if (strcmp(name, "c") == 0) {
        ctx->in_c = 0;
        char *val = ctx->cell_val.data; /* can be NULL */
        char *final_val = NULL;

        if (val) {
             if (strcmp(ctx->cell_type, "s") == 0) {
                 /* Shared string index */
                 int idx = atoi(val);
                 if (ctx->ss && idx >= 0 && idx < ctx->ss->count) {
                     final_val = ctx->ss->strings[idx];
                 }
             } else if (strcmp(ctx->cell_type, "inlineStr") == 0) {
                 final_val = val;
             } else {
                 final_val = val; /* Number or other */
             }
        }
        
        if (!final_val) final_val = "";

        if (!ctx->header_row_processed) {
            /* Building header list */
            if (ctx->header_cols.len > 0) strbuf_append(&ctx->header_cols, ", ", 2);
            strbuf_append(&ctx->header_cols, "\"", 1);
            strbuf_append(&ctx->header_cols, final_val, -1);
            strbuf_append(&ctx->header_cols, "\"", 1);
            ctx->num_cols++;
        } else {
             /* Check if this is the first value being added to buffer? */
             /* Buffer starts with "VALUES (". */
             /* If data[len-1] != '(' implies we added something. */
             
             if (ctx->sql_buf.len > 0 && ctx->sql_buf.data[ctx->sql_buf.len-1] != '(') {
                 strbuf_append(&ctx->sql_buf, ", ", 2);
             }
             
             char *escaped = sqlite3_mprintf("%Q", final_val);
             strbuf_append(&ctx->sql_buf, escaped, -1);
             sqlite3_free(escaped);
        }
        
        strbuf_free(&ctx->cell_val);
    } else if (strcmp(name, "row") == 0) {
        ctx->in_row = 0;
        if (!ctx->header_row_processed) {
            /* Check dimension max col */
            if (ctx->dimension_max_col > ctx->num_cols) {
                int k;
                int start = ctx->num_cols;
                int end = ctx->dimension_max_col;
                for (k = start; k < end; k++) {
                     if (ctx->header_cols.len > 0) strbuf_append(&ctx->header_cols, ", ", 2);
                     char col_name[32];
                     char letter[16];
                     int_to_col(k, letter);
                     snprintf(col_name, sizeof(col_name), "\"%s\"", letter);
                     strbuf_append(&ctx->header_cols, col_name, -1);
                     ctx->num_cols++;
                }
            }
            
            /* Create table */
            char *create_sql = sqlite3_mprintf("CREATE TABLE IF NOT EXISTS \"%w\" (%s);", ctx->table_name, ctx->header_cols.data);
            int rc = sqlite3_exec(ctx->db, create_sql, NULL, NULL, NULL);
            sqlite3_free(create_sql);
            if (rc == SQLITE_OK) {
                ctx->table_created = 1;
            }
            ctx->header_row_processed = 1;
            ctx->last_row_inserted = ctx->current_row;
        } else {
            /* Fill trailing gaps? */
            /* If row ends at col 3 but table has 5 cols. */
            int k;
            for (k = ctx->current_col; k < ctx->num_cols; k++) {
                if (ctx->sql_buf.data[ctx->sql_buf.len-1] != '(') {
                    strbuf_append(&ctx->sql_buf, ", ", 2);
                }
                strbuf_append(&ctx->sql_buf, "NULL", 4);
            }
            
            strbuf_append(&ctx->sql_buf, ");", -1);
            char *zErr = NULL;
            int rc = sqlite3_exec(ctx->db, ctx->sql_buf.data, NULL, NULL, &zErr);
            if (rc != SQLITE_OK) {
                fprintf(stderr, "xlsximport: insert error on row %d: %s\n", ctx->current_row, zErr);
                sqlite3_free(zErr);
            }
            strbuf_free(&ctx->sql_buf);
            ctx->last_row_inserted = ctx->current_row;
        }
    } else if (strcmp(name, "sheetData") == 0 || strcmp(name, "worksheet") == 0) {
        /* Handle trailing rows after last <row> tag */
        if (ctx->header_row_processed) {
            while (ctx->last_row_inserted < ctx->dimension_max_row) {
                ctx->last_row_inserted++;
                strbuf_init(&ctx->sql_buf);
                strbuf_append(&ctx->sql_buf, "INSERT INTO \"", -1);
                strbuf_append(&ctx->sql_buf, ctx->table_name, -1);
                strbuf_append(&ctx->sql_buf, "\" VALUES (", -1);
                int k;
                for (k = 0; k < ctx->num_cols; k++) {
                    if (k > 0) strbuf_append(&ctx->sql_buf, ", ", 2);
                    strbuf_append(&ctx->sql_buf, "NULL", 4);
                }
                strbuf_append(&ctx->sql_buf, ");", 2);
                char *zErr = NULL;
                int rc = sqlite3_exec(ctx->db, ctx->sql_buf.data, NULL, NULL, &zErr);
                if (rc != SQLITE_OK) {
                    fprintf(stderr, "xlsximport: trailing row insert error: %s\n", zErr);
                    sqlite3_free(zErr);
                }
                strbuf_free(&ctx->sql_buf);
            }
        }
    }
}

/* Helper to get file content from zipfile */
static int get_zip_content(sqlite3 *db, const char *zipname, const char *filename, void **buf, int *len) {
    sqlite3_stmt *stmt;
    char *sql = "SELECT data FROM zipfile(?) WHERE name = ?";
    int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    if (rc != SQLITE_OK) return rc;
    
    sqlite3_bind_text(stmt, 1, zipname, -1, SQLITE_STATIC);
    sqlite3_bind_text(stmt, 2, filename, -1, SQLITE_STATIC);
    
    rc = sqlite3_step(stmt);
    if (rc == SQLITE_ROW) {
        const void *blob = sqlite3_column_blob(stmt, 0);
        int bytes = sqlite3_column_bytes(stmt, 0);
        *buf = sqlite3_malloc(bytes);
        memcpy(*buf, blob, bytes);
        *len = bytes;
        rc = SQLITE_OK;
    } else {
        rc = SQLITE_ERROR;
    }
    sqlite3_finalize(stmt);
    return rc;
}

/* Helper to check if a sheet should be imported */
static int should_import(int argc, sqlite3_value **argv, int sheet_idx, const char *sheet_name) {
    if (argc <= 1) return 1; /* Import all if no extra args */
    
    int i;
    for (i = 1; i < argc; i++) {
        int type = sqlite3_value_type(argv[i]);
        if (type == SQLITE_INTEGER) {
            int req_idx = sqlite3_value_int(argv[i]);
            if (req_idx == sheet_idx + 1) return 1;
        } else if (type == SQLITE_TEXT) {
            const char *req_name = (const char *)sqlite3_value_text(argv[i]);
            if (req_name && strcmp(req_name, sheet_name) == 0) return 1;
        }
    }
    return 0;
}

/* Main Function */
static void xlsx_import_func(sqlite3_context *context, int argc, sqlite3_value **argv) {
    if (argc < 1) {
        sqlite3_result_error(context, "xlsx_import requires at least 1 argument", -1);
        return;
    }
    const char *fname = (const char *)sqlite3_value_text(argv[0]);
    sqlite3 *db = sqlite3_context_db_handle(context);
    
    /* 1. Parse Shared Strings */
    SharedStrings ss = {0};
    void *xml_data = NULL;
    int xml_len = 0;
    
    if (get_zip_content(db, fname, "xl/sharedStrings.xml", &xml_data, &xml_len) == SQLITE_OK) {
        XML_Parser parser = XML_ParserCreate(NULL);
        XML_SetUserData(parser, &ss);
        XML_SetElementHandler(parser, shared_strings_start, shared_strings_end);
        XML_SetCharacterDataHandler(parser, shared_strings_char);
        XML_Parse(parser, xml_data, xml_len, 1);
        XML_ParserFree(parser);
        sqlite3_free(xml_data);
    }
    
    /* 2. Parse Workbook to get sheets */
    Workbook wb = {0};
    if (get_zip_content(db, fname, "xl/workbook.xml", &xml_data, &xml_len) == SQLITE_OK) {
        XML_Parser parser = XML_ParserCreate(NULL);
        XML_SetUserData(parser, &wb);
        XML_SetElementHandler(parser, workbook_start, NULL);
        XML_Parse(parser, xml_data, xml_len, 1);
        XML_ParserFree(parser);
        sqlite3_free(xml_data);
    }
    
    /* 3. Process Sheets */
    int i;
    int sheets_imported = 0;
    for (i = 0; i < wb.count; i++) {
        if (!should_import(argc, argv, i, wb.sheets[i].name)) continue;

        char sheet_info_path[64];
        /* Assume sequential numbering based on requirement "xl/worksheets/sheet1.xml to ...sheetN.xml" */
        snprintf(sheet_info_path, sizeof(sheet_info_path), "xl/worksheets/sheet%d.xml", i+1);
        
        if (get_zip_content(db, fname, sheet_info_path, &xml_data, &xml_len) == SQLITE_OK) {
            SheetCtx ctx = {0};
            ctx.db = db;
            ctx.table_name = wb.sheets[i].name;
            ctx.ss = &ss;
            
            XML_Parser parser = XML_ParserCreate(NULL);
            XML_SetUserData(parser, &ctx);
            XML_SetElementHandler(parser, sheet_start, sheet_end);
            XML_SetCharacterDataHandler(parser, sheet_char);
            XML_Parse(parser, xml_data, xml_len, 1);
            XML_ParserFree(parser);
            
            if (ctx.table_created) {
                sheets_imported++;
            }

            sqlite3_free(xml_data);
            strbuf_free(&ctx.header_cols);
            strbuf_free(&ctx.sql_buf);
        }
    }
    
    /* Cleanup */
    for (i = 0; i < ss.count; i++) sqlite3_free(ss.strings[i]);
    sqlite3_free(ss.strings);
    workbook_free(&wb);
    
    sqlite3_result_int(context, sheets_imported);
}

/* Version Function */
static void xlsx_import_version(sqlite3_context *context, int argc, sqlite3_value **argv) {
    (void)argc; (void)argv;
    sqlite3_result_text(context, "2026-01-07 Gemini 3 Pro (High)", -1, SQLITE_STATIC);
}

/* 
** xlsx_import_sheetnames Table-Valued Function 
*/
typedef struct {
    sqlite3_vtab base;
    sqlite3 *db;
} SheetNamesVtab;

typedef struct {
    sqlite3_vtab_cursor base;
    Workbook wb;
    int current_idx;
} SheetNamesCursor;

static int sheetnames_connect(sqlite3 *db, void *pAux, int argc, const char *const *argv, sqlite3_vtab **ppVtab, char **pzErr) {
    (void)pAux; (void)argc; (void)argv; (void)pzErr;
    int rc = sqlite3_declare_vtab(db, "CREATE TABLE x(sheet_num INTEGER, sheet_name TEXT, filename HIDDEN)");
    if (rc != SQLITE_OK) return rc;
    SheetNamesVtab *vtab = sqlite3_malloc(sizeof(SheetNamesVtab));
    if (!vtab) return SQLITE_NOMEM;
    memset(vtab, 0, sizeof(SheetNamesVtab));
    vtab->db = db;
    *ppVtab = &vtab->base;
    return SQLITE_OK;
}

static int sheetnames_disconnect(sqlite3_vtab *vtab) {
    sqlite3_free(vtab);
    return SQLITE_OK;
}

static int sheetnames_open(sqlite3_vtab *vtab, sqlite3_vtab_cursor **ppCursor) {
    (void)vtab;
    SheetNamesCursor *cur = sqlite3_malloc(sizeof(SheetNamesCursor));
    if (!cur) return SQLITE_NOMEM;
    memset(cur, 0, sizeof(SheetNamesCursor));
    *ppCursor = &cur->base;
    return SQLITE_OK;
}

static int sheetnames_close(sqlite3_vtab_cursor *cur) {
    SheetNamesCursor *pCur = (SheetNamesCursor *)cur;
    workbook_free(&pCur->wb);
    sqlite3_free(pCur);
    return SQLITE_OK;
}

static int sheetnames_next(sqlite3_vtab_cursor *cur) {
    SheetNamesCursor *pCur = (SheetNamesCursor *)cur;
    pCur->current_idx++;
    return SQLITE_OK;
}

static int sheetnames_column(sqlite3_vtab_cursor *cur, sqlite3_context *ctx, int i) {
    SheetNamesCursor *pCur = (SheetNamesCursor *)cur;
    if (i == 0) {
        sqlite3_result_int(ctx, pCur->current_idx + 1);
    } else if (i == 1) {
        if (pCur->current_idx < pCur->wb.count) {
            sqlite3_result_text(ctx, pCur->wb.sheets[pCur->current_idx].name, -1, SQLITE_TRANSIENT);
        }
    }
    return SQLITE_OK;
}

static int sheetnames_rowid(sqlite3_vtab_cursor *cur, sqlite_int64 *pRowid) {
    SheetNamesCursor *pCur = (SheetNamesCursor *)cur;
    *pRowid = pCur->current_idx + 1;
    return SQLITE_OK;
}

static int sheetnames_eof(sqlite3_vtab_cursor *cur) {
    SheetNamesCursor *pCur = (SheetNamesCursor *)cur;
    return pCur->current_idx >= pCur->wb.count;
}

static int sheetnames_filter(sqlite3_vtab_cursor *cur, int idxNum, const char *idxStr, int argc, sqlite3_value **argv) {
    (void)idxNum; (void)idxStr;
    SheetNamesCursor *pCur = (SheetNamesCursor *)cur;
    SheetNamesVtab *vtab = (SheetNamesVtab *)cur->pVtab;
    
    workbook_free(&pCur->wb);
    pCur->current_idx = 0;
    
    if (argc < 1) {
        vtab->base.zErrMsg = sqlite3_mprintf("xlsx_import_sheetnames requires a filename argument");
        return SQLITE_ERROR;
    }
    
    const char *fname = (const char *)sqlite3_value_text(argv[0]);
    if (!fname) return SQLITE_ERROR;
    
    void *xml_data = NULL;
    int xml_len = 0;
    
    int rc = get_zip_content(vtab->db, fname, "xl/workbook.xml", &xml_data, &xml_len);
    if (rc != SQLITE_OK) {
        /* If we can't read workbook.xml, maybe the file doesn't exist or isn't a valid xlsx/zip. */
        /* For a vtab filter, returning ERROR might break the query hard. 
           Should we return OK with empty set? 
           Let's return error to inform user.
        */
        vtab->base.zErrMsg = sqlite3_mprintf("Failed to read workbook from %s", fname);
        return SQLITE_ERROR;
    }
    
    XML_Parser parser = XML_ParserCreate(NULL);
    XML_SetUserData(parser, &pCur->wb);
    XML_SetElementHandler(parser, workbook_start, NULL);
    if (XML_Parse(parser, xml_data, xml_len, 1) == XML_STATUS_ERROR) {
        XML_ParserFree(parser);
        sqlite3_free(xml_data);
        return SQLITE_ERROR;
    }
    XML_ParserFree(parser);
    sqlite3_free(xml_data);
    
    return SQLITE_OK;
}

static int sheetnames_bestindex(sqlite3_vtab *vtab, sqlite3_index_info *pIdxInfo) {
    (void)vtab;
    int i;
    int filename_idx = -1;
    
    for (i = 0; i < pIdxInfo->nConstraint; i++) {
        if (pIdxInfo->aConstraint[i].usable && pIdxInfo->aConstraint[i].iColumn == 2 && pIdxInfo->aConstraint[i].op == SQLITE_INDEX_CONSTRAINT_EQ) {
            filename_idx = i;
            break;
        }
    }
    
    if (filename_idx == -1) return SQLITE_CONSTRAINT;
    
    pIdxInfo->aConstraintUsage[filename_idx].argvIndex = 1;
    pIdxInfo->aConstraintUsage[filename_idx].omit = 1;
    pIdxInfo->estimatedCost = 1000.0;
    pIdxInfo->estimatedRows = 10;
    return SQLITE_OK;
}

static sqlite3_module sheetnamesModule = {
    0,
    sheetnames_connect,
    sheetnames_connect,
    sheetnames_bestindex,
    sheetnames_disconnect,
    sheetnames_disconnect,
    sheetnames_open,
    sheetnames_close,
    sheetnames_filter,
    sheetnames_next,
    sheetnames_eof,
    sheetnames_column,
    sheetnames_rowid,
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
};

/* Init */
#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsximport_init(sqlite3 *db, char **pzErrMsg, const sqlite3_api_routines *pApi) {
    int rc = SQLITE_OK;
    SQLITE_EXTENSION_INIT2(pApi);
    (void)pzErrMsg;
    
    rc = sqlite3_create_function(db, "xlsx_import", -1, SQLITE_UTF8, NULL, xlsx_import_func, NULL, NULL);
    if (rc == SQLITE_OK) {
        rc = sqlite3_create_function(db, "xlsx_import_version", 0, SQLITE_UTF8, NULL, xlsx_import_version, NULL, NULL);
    }
    if (rc == SQLITE_OK) {
        rc = sqlite3_create_module(db, "xlsx_import_sheetnames", &sheetnamesModule, NULL);
    }
    
    return rc;
}
