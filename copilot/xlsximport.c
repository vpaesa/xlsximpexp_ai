/*
User prompts (included verbatim as requested):

Create C code for a SQLite extension named xlsximport. Use the SQLite zipfile extension to open a XLSX file and gather this content:
    xl/sharedStrings.xml
    xl/worksheets/sheet1.xml to  xl/worksheets/sheetN.xml
    xl/workbook.xml
The name of each sheet is in xl/workbook.xml
The individual sheets are kept in xl/worksheets/sheet1.xml  to  xl/worksheets/sheetN.xml
To save on space, Microsoft stores all the character literal values in one common xl/sharedStrings.xml dictionary file. The individual cell value found for this string in the actual sheet1.xml file is just an index into this dictionary.
Microsoft does not store empty cells or rows in xl/worksheets/sheet1.xml, so any gaps between values have to be taken care by the code.
Excel cell content length limit is 32767 characters.
Create a SQL function named xlsx_import that creates one table for each of the sheets in the XLSX files, table name equal to sheet name, and column names equal to the values in first row of the sheet.
The first parameter is the XLSX filename. Subsequent optional parameters are sheet names or sheet numbers (1-based) to import.
Use expat for XML parsing. Add support for both shared and inline strings.  
Do not perform table name and column sanititazion. Use proper quoting instead. 
Add SQL function xlsx_import_version returning "2025-12-30 Copilot Think Deeper (GPT 5.1?)". 
Add all user prompts as comments.

Usage:
.load xlsximport.so
SELECT xlsx_import('filename.xlsx');  -- Import all sheets
SELECT xlsx_import('filename.xlsx', 'Sheet1', 'Sheet2');  -- Import specific sheets by name
SELECT xlsx_import('filename.xlsx', 1, 3);  -- Import sheets by number (1-based)
SELECT xlsx_import('filename.xlsx', 'Sheet1', 2);  -- Mix of names and numbers
SELECT xlsx_import_version();
*/

/*
Limitations and notes:
- This code uses the zipfile virtual table via SQL queries to fetch file contents.
- The code is best-effort and does not implement every XLSX edge case (shared styles,
  relationships, external references, complex rich text formatting, etc.).
*/

#define _GNU_SOURCE
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <expat.h>
#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

/* Version function */
static void xlsx_import_version(sqlite3_context *context, int argc, sqlite3_value **argv){
    (void)argc; (void)argv;
    sqlite3_result_text(context, "2025-12-30 Copilot Think Deeper (GPT 5.1?)", -1, SQLITE_STATIC);
}

/* Quote an identifier for use as a SQLite identifier.
   Example:  Sheet "A"  ->  "Sheet ""A"""
   Returns malloc'd string, caller must free.
*/
static char *quote_identifier(const char *s){
    if(!s) s = "";
    size_t len = strlen(s);
    /* worst-case every char is a quote -> need 2*len + 2 for surrounding quotes + 1 for NUL */
    size_t cap = len * 2 + 3;
    char *out = (char*)malloc(cap);
    if(!out) return NULL;
    char *p = out;
    *p++ = '"';
    for(const char *q = s; *q; ++q){
        if(*q == '"'){
            *p++ = '"'; /* double the quote */
            *p++ = '"';
        } else {
            *p++ = *q;
        }
    }
    *p++ = '"';
    *p = '\0';
    return out;
}

/* Simple dynamic string buffer */
typedef struct {
    char *buf;
    size_t len;
    size_t cap;
} strbuf;

static void sb_init(strbuf *s){
    s->cap = 1024;
    s->len = 0;
    s->buf = (char*)malloc(s->cap);
    if(s->buf) s->buf[0] = '\0';
}
static void sb_append(strbuf *s, const char *t){
    if(!s || !s->buf) return;
    size_t tl = strlen(t);
    if(s->len + tl + 1 > s->cap){
        while(s->len + tl + 1 > s->cap) s->cap *= 2;
        s->buf = (char*)realloc(s->buf, s->cap);
    }
    memcpy(s->buf + s->len, t, tl+1);
    s->len += tl;
}
static void sb_append_buf(strbuf *s, const char *t, size_t len){
    if(!s || !s->buf) return;
    if(s->len + len + 1 > s->cap){
        while(s->len + len + 1 > s->cap) s->cap *= 2;
        s->buf = (char*)realloc(s->buf, s->cap);
    }
    memcpy(s->buf + s->len, t, len);
    s->len += len;
    s->buf[s->len] = '\0';
}
static void sb_free(strbuf *s){
    if(!s) return;
    free(s->buf);
    s->buf = NULL;
    s->len = s->cap = 0;
}

/* Shared strings container */
typedef struct {
    char **items;
    size_t n;
    size_t cap;
} sstrings;

static void sstrings_init(sstrings *ss){
    ss->n = 0; ss->cap = 64;
    ss->items = (char**)malloc(sizeof(char*) * ss->cap);
}
static void sstrings_add(sstrings *ss, const char *s){
    if(ss->n >= ss->cap){
        ss->cap *= 2;
        ss->items = (char**)realloc(ss->items, sizeof(char*) * ss->cap);
    }
    ss->items[ss->n++] = strdup(s ? s : "");
}
static void sstrings_free(sstrings *ss){
    for(size_t i=0;i<ss->n;i++) free(ss->items[i]);
    free(ss->items);
    ss->items = NULL; ss->n = ss->cap = 0;
}

/* Helper: convert Excel column letters to 0-based index (A->0, B->1, Z->25, AA->26) */
static int colname_to_index(const char *col){
    int idx = 0;
    for(const char *p = col; *p; ++p){
        if(*p >= 'A' && *p <= 'Z') idx = idx*26 + (*p - 'A' + 1);
        else if(*p >= 'a' && *p <= 'z') idx = idx*26 + (*p - 'a' + 1);
        else break;
    }
    return idx - 1;
}

/* --- Expat parsers --- */

/* Parser for sharedStrings.xml */
typedef struct {
    XML_Parser parser;
    sstrings *ss;
    int in_si;
    int in_t;
    strbuf cur;
} ss_parser_ctx;

static void ss_start(void *userData, const XML_Char *name, const XML_Char **atts){
    (void)atts;
    ss_parser_ctx *ctx = (ss_parser_ctx*)userData;
    if(strcmp(name, "si")==0){
        ctx->in_si = 1;
        ctx->cur.len = 0;
        ctx->cur.buf[0] = '\0';
    } else if(strcmp(name, "t")==0 && ctx->in_si){
        ctx->in_t = 1;
    }
}
static void ss_end(void *userData, const XML_Char *name){
    ss_parser_ctx *ctx = (ss_parser_ctx*)userData;
    if(strcmp(name, "si")==0){
        ctx->in_si = 0;
        sstrings_add(ctx->ss, ctx->cur.buf);
    } else if(strcmp(name, "t")==0){
        ctx->in_t = 0;
    }
}
static void ss_char(void *userData, const XML_Char *s, int len){
    ss_parser_ctx *ctx = (ss_parser_ctx*)userData;
    if(ctx->in_si && ctx->in_t){
        sb_append_buf(&ctx->cur, s, (size_t)len);
    }
}

/* Parser for workbook.xml to extract sheet names and sheetIds */
typedef struct {
    XML_Parser parser;
    char **names;
    int *sheetIds;
    size_t n;
    size_t cap;
} wb_parser_ctx;

static void wb_init(wb_parser_ctx *ctx){
    ctx->n = 0; ctx->cap = 16;
    ctx->names = (char**)malloc(sizeof(char*) * ctx->cap);
    ctx->sheetIds = (int*)malloc(sizeof(int) * ctx->cap);
}
static void wb_free(wb_parser_ctx *ctx){
    for(size_t i=0;i<ctx->n;i++) free(ctx->names[i]);
    free(ctx->names);
    free(ctx->sheetIds);
    ctx->names = NULL; ctx->sheetIds = NULL; ctx->n = ctx->cap = 0;
}
static void wb_start(void *userData, const XML_Char *name, const XML_Char **atts){
    wb_parser_ctx *ctx = (wb_parser_ctx*)userData;
    if(strcmp(name, "sheet")==0){
        const XML_Char *nm = NULL;
        const XML_Char *sid = NULL;
        for(int i=0; atts[i]; i+=2){
            if(strcmp(atts[i], "name")==0) nm = atts[i+1];
            else if(strcmp(atts[i], "sheetId")==0) sid = atts[i+1];
        }
        if(nm){
            if(ctx->n >= ctx->cap){
                ctx->cap *= 2;
                ctx->names = (char**)realloc(ctx->names, sizeof(char*) * ctx->cap);
                ctx->sheetIds = (int*)realloc(ctx->sheetIds, sizeof(int) * ctx->cap);
            }
            ctx->names[ctx->n] = strdup((const char*)nm);
            ctx->sheetIds[ctx->n] = sid ? atoi((const char*)sid) : (int)(ctx->n+1);
            ctx->n++;
        }
    }
}
static void wb_end(void *userData, const XML_Char *name){ (void)userData; (void)name; }
static void wb_char(void *userData, const XML_Char *s, int len){ (void)userData; (void)s; (void)len; }

/* Parser for worksheet XML (sheetN.xml) */
typedef struct {
    XML_Parser parser;
    sstrings *shared;
    int in_v;
    int in_t;
    int in_is;
    int in_c;
    char cur_cell_ref[64];
    char cur_cell_type[32];
    strbuf cur_text;
    int current_row;
    char **rowbuf;
    size_t rowcap;
    size_t maxcol;
    void (*emit_row)(int rownum, char **cols, size_t ncols, void *udata);
    void *emit_udata;
} sheet_parser_ctx;

static void ensure_rowcap(sheet_parser_ctx *ctx, size_t cols){
    if(cols <= ctx->rowcap) return;
    size_t newcap = ctx->rowcap ? ctx->rowcap : 16;
    while(newcap < cols) newcap *= 2;
    ctx->rowbuf = (char**)realloc(ctx->rowbuf, sizeof(char*) * newcap);
    for(size_t i=ctx->rowcap;i<newcap;i++) ctx->rowbuf[i] = NULL;
    ctx->rowcap = newcap;
}

static void sheet_start(void *userData, const XML_Char *name, const XML_Char **atts){
    sheet_parser_ctx *ctx = (sheet_parser_ctx*)userData;
    if(strcmp(name, "row")==0){
        ctx->current_row = 0;
        for(int i=0; atts[i]; i+=2){
            if(strcmp(atts[i], "r")==0) ctx->current_row = atoi(atts[i+1]);
        }
        if(ctx->rowbuf){
            for(size_t i=0;i<ctx->rowcap;i++){
                if(ctx->rowbuf[i]) { free(ctx->rowbuf[i]); ctx->rowbuf[i] = NULL; }
            }
        }
        ctx->maxcol = 0;
    } else if(strcmp(name, "c")==0){
        ctx->in_c = 1;
        ctx->cur_cell_ref[0] = '\0';
        ctx->cur_cell_type[0] = '\0';
        for(int i=0; atts[i]; i+=2){
            if(strcmp(atts[i], "r")==0) strncpy(ctx->cur_cell_ref, atts[i+1], sizeof(ctx->cur_cell_ref)-1);
            else if(strcmp(atts[i], "t")==0) strncpy(ctx->cur_cell_type, atts[i+1], sizeof(ctx->cur_cell_type)-1);
        }
        ctx->cur_text.len = 0;
        ctx->cur_text.buf[0] = '\0';
    } else if(strcmp(name, "v")==0){
        ctx->in_v = 1;
    } else if(strcmp(name, "is")==0){
        ctx->in_is = 1;
    } else if(strcmp(name, "t")==0){
        ctx->in_t = 1;
    }
}
static void sheet_end(void *userData, const XML_Char *name){
    sheet_parser_ctx *ctx = (sheet_parser_ctx*)userData;
    if(strcmp(name, "c")==0){
        char colletters[32] = {0};
        int i=0;
        while(ctx->cur_cell_ref[i] && !isdigit((unsigned char)ctx->cur_cell_ref[i]) && i < (int)sizeof(colletters)-1){
            colletters[i] = ctx->cur_cell_ref[i];
            i++;
        }
        colletters[i] = '\0';
        int colidx = colname_to_index(colletters);
        if(colidx < 0) colidx = 0;
        ensure_rowcap(ctx, (size_t)colidx+1);
        char *val = NULL;
        if(ctx->cur_cell_type[0] == 's' && ctx->cur_text.len > 0){
            int idx = atoi(ctx->cur_text.buf);
            if(idx >= 0 && (size_t)idx < ctx->shared->n){
                val = strdup(ctx->shared->items[idx]);
            } else {
                val = strdup("");
            }
        } else if(ctx->in_is){
            val = strdup(ctx->cur_text.buf);
        } else if(ctx->in_v){
            val = strdup(ctx->cur_text.buf);
        } else {
            val = strdup(ctx->cur_text.buf);
        }
        if(ctx->rowbuf[colidx]) free(ctx->rowbuf[colidx]);
        ctx->rowbuf[colidx] = val;
        if((size_t)colidx + 1 > ctx->maxcol) ctx->maxcol = (size_t)colidx + 1;
        ctx->in_c = 0;
        ctx->in_v = 0;
        ctx->in_is = 0;
        ctx->cur_text.len = 0;
        ctx->cur_text.buf[0] = '\0';
    } else if(strcmp(name, "row")==0){
        ctx->emit_row(ctx->current_row, ctx->rowbuf, ctx->maxcol, ctx->emit_udata);
        for(size_t i=0;i<ctx->rowcap;i++){
            if(ctx->rowbuf[i]) { free(ctx->rowbuf[i]); ctx->rowbuf[i] = NULL; }
        }
        ctx->maxcol = 0;
    } else if(strcmp(name, "v")==0){
        ctx->in_v = 0;
    } else if(strcmp(name, "t")==0){
        ctx->in_t = 0;
    } else if(strcmp(name, "is")==0){
        ctx->in_is = 0;
    }
}
static void sheet_char(void *userData, const XML_Char *s, int len){
    sheet_parser_ctx *ctx = (sheet_parser_ctx*)userData;
    if(ctx->in_v || ctx->in_t){
        sb_append_buf(&ctx->cur_text, s, (size_t)len);
    }
}

/* Helper: read a file from the .xlsx archive using the SQLite zipfile extension.
   This function queries the zipfile virtual table for the given archive and internal name.
   It returns a malloc'd null-terminated buffer (caller must free) and optionally sets out_len.
   If the file is not found, returns NULL.
*/
static char *read_zip_file_sqlite(sqlite3 *db, const char *archive, const char *internal_name, size_t *out_len){
    if(!db || !archive || !internal_name) return NULL;
    const char *sql =
        "SELECT data FROM zipfile(?) WHERE name = ? LIMIT 1;";
    sqlite3_stmt *stmt = NULL;
    if(sqlite3_prepare_v2(db, sql, -1, &stmt, NULL) != SQLITE_OK){
        return NULL;
    }
    sqlite3_bind_text(stmt, 1, archive, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(stmt, 2, internal_name, -1, SQLITE_TRANSIENT);
    char *result = NULL;
    int rc = sqlite3_step(stmt);
    if(rc == SQLITE_ROW){
        const void *blob = sqlite3_column_blob(stmt, 0);
        int bytes = sqlite3_column_bytes(stmt, 0);
        if(blob && bytes > 0){
            result = (char*)malloc((size_t)bytes + 1);
            memcpy(result, blob, (size_t)bytes);
            result[bytes] = '\0';
            if(out_len) *out_len = (size_t)bytes;
        } else {
            /* empty file -> return empty string */
            result = strdup("");
            if(out_len) *out_len = 0;
        }
    }
    sqlite3_finalize(stmt);
    return result;
}

/* sheet rows collector */
typedef struct {
    int rownum;
    char **cols;
    size_t ncols;
} sheet_row;

typedef struct {
    sheet_row *rows;
    size_t n;
    size_t cap;
} sheet_rows;

static void sheet_rows_init(sheet_rows *sr){
    sr->n = 0; sr->cap = 64;
    sr->rows = (sheet_row*)malloc(sizeof(sheet_row) * sr->cap);
}
static void sheet_rows_free(sheet_rows *sr){
    for(size_t i=0;i<sr->n;i++){
        for(size_t j=0;j<sr->rows[i].ncols;j++) if(sr->rows[i].cols[j]) free(sr->rows[i].cols[j]);
        free(sr->rows[i].cols);
    }
    free(sr->rows);
    sr->rows = NULL; sr->n = sr->cap = 0;
}
static void sheet_rows_emit(int rownum, char **cols, size_t ncols, void *udata){
    sheet_rows *sr = (sheet_rows*)udata;
    if(sr->n >= sr->cap){
        sr->cap *= 2;
        sr->rows = (sheet_row*)realloc(sr->rows, sizeof(sheet_row) * sr->cap);
    }
    char **copycols = (char**)malloc(sizeof(char*) * ncols);
    for(size_t i=0;i<ncols;i++){
        copycols[i] = cols[i] ? strdup(cols[i]) : NULL;
    }
    sr->rows[sr->n].rownum = rownum;
    sr->rows[sr->n].cols = copycols;
    sr->rows[sr->n].ncols = ncols;
    sr->n++;
}

/* Helper: check if a string is an integer (consists of digits only) */
static int is_integer_string(const char *s){
    if(!s || *s == '\0') return 0;
    const char *p = s;
    if(*p == '+' || *p == '-') p++;
    if(!*p) return 0;
    while(*p){
        if(!isdigit((unsigned char)*p)) return 0;
        p++;
    }
    return 1;
}

/* Helper: decide whether to import a sheet based on provided selectors.
   selectors: array of strings (names or integers). selector_count may be 0 -> import all.
   For integer selectors, match if selector == sheetId OR selector == (si+1) (1-based index).
   For name selectors, case-sensitive exact match against workbook name.
*/
static int should_import_sheet(const wb_parser_ctx *wb, size_t si, int sheetId, const char **selectors, int selector_count){
    if(selector_count <= 0) return 1; /* import all */
    for(int i=0;i<selector_count;i++){
        const char *sel = selectors[i];
        if(!sel) continue;
        if(is_integer_string(sel)){
            int val = atoi(sel);
            if(val == sheetId) return 1;
            if(val == (int)(si + 1)) return 1;
        } else {
            if(wb->names && si < wb->n && wb->names[si] && strcmp(wb->names[si], sel) == 0) return 1;
        }
    }
    return 0;
}

/* Main worker: parse sharedStrings.xml, workbook.xml, and each sheet, then create tables and insert rows.
   Uses read_zip_file_sqlite() to fetch files from the archive.
   New: accepts selectors array (sheet names or integers as strings) and selector_count.
   Uses quoting for table and column identifiers instead of sanitization.
*/
static int import_xlsx_to_db(sqlite3 *db, const char *filename, const char **selectors, int selector_count, sqlite3_context *ctx){
    if(!db || !filename){
        sqlite3_result_error(ctx, "Invalid arguments to import_xlsx_to_db", -1);
        return SQLITE_ERROR;
    }
    int tables_created = 0;

    /* 1) Read sharedStrings.xml if present */
    sstrings ss;
    sstrings_init(&ss);

    size_t tmp_len = 0;
    char *shared_buf = read_zip_file_sqlite(db, filename, "xl/sharedStrings.xml", &tmp_len);
    if(shared_buf){
        ss_parser_ctx sctx;
        sctx.parser = XML_ParserCreate(NULL);
        sctx.ss = &ss;
        sctx.in_si = sctx.in_t = 0;
        sb_init(&sctx.cur);
        XML_SetUserData(sctx.parser, &sctx);
        XML_SetElementHandler(sctx.parser, ss_start, ss_end);
        XML_SetCharacterDataHandler(sctx.parser, ss_char);
        if(XML_Parse(sctx.parser, shared_buf, (int)strlen(shared_buf), XML_TRUE) == XML_STATUS_ERROR){
            /* ignore parse errors but continue */
        }
        XML_ParserFree(sctx.parser);
        sb_free(&sctx.cur);
        free(shared_buf);
    }

    /* 2) Read workbook.xml to get sheet names and sheetIds */
    wb_parser_ctx wb;
    wb_init(&wb);
    char *wb_buf = read_zip_file_sqlite(db, filename, "xl/workbook.xml", &tmp_len);
    if(wb_buf){
        XML_Parser p = XML_ParserCreate(NULL);
        XML_SetUserData(p, &wb);
        XML_SetElementHandler(p, wb_start, wb_end);
        XML_SetCharacterDataHandler(p, wb_char);
        XML_Parse(p, wb_buf, (int)strlen(wb_buf), XML_TRUE);
        XML_ParserFree(p);
        free(wb_buf);
    } else {
        sqlite3_result_error(ctx, "xl/workbook.xml not found in archive (zipfile)", -1);
        sstrings_free(&ss);
        wb_free(&wb);
        return SQLITE_ERROR;
    }

    /* For each sheet in workbook, read corresponding sheet XML and import if selected */
    for(size_t si = 0; si < wb.n; ++si){
        const char *sheet_name_raw = wb.names[si];
        int sheetId = wb.sheetIds[si];

        if(!should_import_sheet(&wb, si, sheetId, selectors, selector_count)){
            continue; /* skip this sheet */
        }

        char sheet_internal[256];
        snprintf(sheet_internal, sizeof(sheet_internal), "xl/worksheets/sheet%d.xml", sheetId);

        char *sheet_buf = read_zip_file_sqlite(db, filename, sheet_internal, &tmp_len);
        if(!sheet_buf){
            /* fallback to sequential index */
            snprintf(sheet_internal, sizeof(sheet_internal), "xl/worksheets/sheet%lu.xml", (unsigned long)(si + 1));
            sheet_buf = read_zip_file_sqlite(db, filename, sheet_internal, &tmp_len);
            if(!sheet_buf){
                /* skip missing sheet */
                continue;
            }
        }

        sheet_rows rows;
        sheet_rows_init(&rows);

        sheet_parser_ctx sp;
        sp.parser = XML_ParserCreate(NULL);
        sp.shared = &ss;
        sp.in_v = sp.in_t = sp.in_is = sp.in_c = 0;
        sp.cur_cell_ref[0] = '\0';
        sp.cur_cell_type[0] = '\0';
        sb_init(&sp.cur_text);
        sp.rowbuf = NULL;
        sp.rowcap = 0;
        sp.maxcol = 0;
        sp.emit_row = sheet_rows_emit;
        sp.emit_udata = &rows;

        XML_SetUserData(sp.parser, &sp);
        XML_SetElementHandler(sp.parser, sheet_start, sheet_end);
        XML_SetCharacterDataHandler(sp.parser, sheet_char);
        if(XML_Parse(sp.parser, sheet_buf, (int)strlen(sheet_buf), XML_TRUE) == XML_STATUS_ERROR){
            /* continue best-effort */
        }
        XML_ParserFree(sp.parser);
        sb_free(&sp.cur_text);
        free(sp.rowbuf);
        free(sheet_buf);

        if(rows.n == 0){
            char *tblq = quote_identifier(sheet_name_raw);
            if(!tblq) { sheet_rows_free(&rows); continue; }
            char sql[1024];
            snprintf(sql, sizeof(sql), "CREATE TABLE IF NOT EXISTS %s (rowid INTEGER PRIMARY KEY);", tblq);
            char *errmsg = NULL;
            if(sqlite3_exec(db, sql, NULL, NULL, &errmsg) != SQLITE_OK){
                sqlite3_free(errmsg);
                free(tblq);
                sheet_rows_free(&rows);
                continue;
            }
            free(tblq);
            tables_created++;
            sheet_rows_free(&rows);
            continue;
        }

        /* Determine header row (first row encountered). */
        int min_rownum = rows.rows[0].rownum;
        for(size_t r=1;r<rows.n;r++) if(rows.rows[r].rownum < min_rownum) min_rownum = rows.rows[r].rownum;
        size_t header_idx = 0;
        for(size_t r=0;r<rows.n;r++) if(rows.rows[r].rownum == min_rownum){ header_idx = r; break; }

        size_t ncols = rows.rows[header_idx].ncols;
        /* Use raw header text as column names, but ensure uniqueness by appending suffixes when duplicates occur */
        char **colnames = (char**)malloc(sizeof(char*) * ncols);
        for(size_t c=0;c<ncols;c++){
            const char *raw = (c < rows.rows[header_idx].ncols && rows.rows[header_idx].cols[c]) ? rows.rows[header_idx].cols[c] : "";
            if(raw == NULL) raw = "";
            /* Start with raw header text (may be empty) */
            char *candidate = strdup(raw);
            int suffix = 1;
            while(1){
                int dup = 0;
                for(size_t j=0;j<c;j++){
                    if(strcmp(colnames[j], candidate) == 0){ dup = 1; break; }
                }
                if(!dup) break;
                char tmp[1024];
                snprintf(tmp, sizeof(tmp), "%s_%d", candidate, suffix++);
                free(candidate);
                candidate = strdup(tmp);
            }
            colnames[c] = candidate;
        }

        char *tblq = quote_identifier(sheet_name_raw);
        if(!tblq){
            for(size_t c=0;c<ncols;c++) free(colnames[c]);
            free(colnames);
            sheet_rows_free(&rows);
            continue;
        }

        /* Build CREATE TABLE SQL using quoted identifiers */
        strbuf create_sql;
        sb_init(&create_sql);
        sb_append(&create_sql, "CREATE TABLE IF NOT EXISTS ");
        sb_append(&create_sql, tblq);
        sb_append(&create_sql, " (");
        for(size_t c=0;c<ncols;c++){
            char *colq = quote_identifier(colnames[c]);
            if(!colq) colq = strdup("\"\""); /* fallback */
            sb_append(&create_sql, colq);
            sb_append(&create_sql, " TEXT");
            if(c+1 < ncols) sb_append(&create_sql, ", ");
            free(colq);
        }
        sb_append(&create_sql, ");");

        char *errmsg = NULL;
        if(sqlite3_exec(db, create_sql.buf, NULL, NULL, &errmsg) != SQLITE_OK){
            sqlite3_free(errmsg);
            for(size_t c=0;c<ncols;c++) free(colnames[c]);
            free(colnames);
            free(tblq);
            sb_free(&create_sql);
            sheet_rows_free(&rows);
            continue;
        }
        sb_free(&create_sql);

        /* Prepare INSERT statement with quoted column names and parameter placeholders */
        strbuf insert_sql;
        sb_init(&insert_sql);
        sb_append(&insert_sql, "INSERT INTO ");
        sb_append(&insert_sql, tblq);
        sb_append(&insert_sql, " (");
        for(size_t c=0;c<ncols;c++){
            char *colq = quote_identifier(colnames[c]);
            if(!colq) colq = strdup("\"\"");
            sb_append(&insert_sql, colq);
            if(c+1 < ncols) sb_append(&insert_sql, ", ");
            free(colq);
        }
        sb_append(&insert_sql, ") VALUES (");
        for(size_t c=0;c<ncols;c++){
            sb_append(&insert_sql, "?");
            if(c+1 < ncols) sb_append(&insert_sql, ", ");
        }
        sb_append(&insert_sql, ");");

        sqlite3_stmt *stmt = NULL;
        if(sqlite3_prepare_v2(db, insert_sql.buf, -1, &stmt, NULL) != SQLITE_OK){
            for(size_t c=0;c<ncols;c++) free(colnames[c]);
            free(colnames);
            free(tblq);
            sb_free(&insert_sql);
            sheet_rows_free(&rows);
            continue;
        }
        sb_free(&insert_sql);

        /* Insert rows: skip header row */
        for(size_t r=0;r<rows.n;r++){
            if(rows.rows[r].rownum == min_rownum) continue;
            sqlite3_reset(stmt);
            sqlite3_clear_bindings(stmt);
            for(size_t c=0;c<ncols;c++){
                const char *val = NULL;
                if(c < rows.rows[r].ncols) val = rows.rows[r].cols[c];
                if(val) sqlite3_bind_text(stmt, (int)c+1, val, -1, SQLITE_TRANSIENT);
                else sqlite3_bind_null(stmt, (int)c+1);
            }
            if(sqlite3_step(stmt) != SQLITE_DONE){
                /* ignore row insert errors */
            }
        }
        sqlite3_finalize(stmt);

        for(size_t c=0;c<ncols;c++) free(colnames[c]);
        free(colnames);
        free(tblq);
        sheet_rows_free(&rows);

        tables_created++;
    }

    sstrings_free(&ss);
    wb_free(&wb);

    sqlite3_result_int(ctx, tables_created);
    return SQLITE_OK;
}

/* SQLite user function wrapper: xlsx_import(filename, [sheet1, sheet2, ...])
   - If only filename is provided, import all sheets.
   - Additional parameters may be sheet names (string) or integers (sheet number or sheetId).
   - Example:
       SELECT xlsx_import('file.xlsx'); -- import all sheets
       SELECT xlsx_import('file.xlsx', 'Sheet1', '3', 'Sheet 4'); -- import Sheet1, sheet number 3, and sheet named "Sheet 4"
*/
static void xlsx_import_func(sqlite3_context *context, int argc, sqlite3_value **argv){
    sqlite3 *db = sqlite3_context_db_handle(context);
    if(argc < 1 || sqlite3_value_type(argv[0]) == SQLITE_NULL){
        sqlite3_result_error(context, "xlsx_import requires a filename argument", -1);
        return;
    }
    const unsigned char *fname = sqlite3_value_text(argv[0]);
    if(!fname){
        sqlite3_result_error(context, "Invalid filename", -1);
        return;
    }

    /* Collect selectors (argv[1]..argv[argc-1]) as strings */
    int selector_count = 0;
    const char **selectors = NULL;
    if(argc > 1){
        selector_count = argc - 1;
        selectors = (const char**)malloc(sizeof(char*) * selector_count);
        if(!selectors) selector_count = 0;
        for(int i=1;i<argc;i++){
            const unsigned char *v = sqlite3_value_text(argv[i]);
            if(v){
                selectors[i-1] = strdup((const char*)v);
            } else {
                selectors[i-1] = NULL;
            }
        }
    }

    /* Call main importer with selectors */
    import_xlsx_to_db(db, (const char*)fname, selectors, selector_count, context);

    /* free selectors */
    if(selectors){
        for(int i=0;i<selector_count;i++) if(selectors[i]) free((void*)selectors[i]);
        free(selectors);
    }
}

/* Extension entry point */
#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsximport_init(sqlite3 *db, char **pzErrMsg, const sqlite3_api_routines *pApi){
    SQLITE_EXTENSION_INIT2(pApi);
    (void)pzErrMsg;
    int rc = SQLITE_OK;
    rc = sqlite3_create_function(db, "xlsx_import", -1, SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL, xlsx_import_func, NULL, NULL);
    if(rc != SQLITE_OK) return rc;
    rc = sqlite3_create_function(db, "xlsx_import_version", 0, SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL, xlsx_import_version, NULL, NULL);
    return rc;
}

/* End of xlsximport.c */
