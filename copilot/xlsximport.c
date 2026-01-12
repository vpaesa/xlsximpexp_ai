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
Do not perform table name and column sanitization. Use proper quoting instead.
Add xlsx_import_sheetnames() table-valued function that returns the names of the sheets in the file.
Add SQL function xlsx_import_version returning "2026-01-07 Copilot Think Deeper (GPT 5.1?)". 
Add all user prompts as comments.

Usage:
.load xlsximport.so
SELECT xlsx_import('filename.xlsx');  -- Import all sheets
SELECT xlsx_import('filename.xlsx', 'Sheet1', 'Sheet2');  -- Import specific sheets by name
SELECT xlsx_import('filename.xlsx', 1, 3);  -- Import sheets by number (1-based)
SELECT xlsx_import('filename.xlsx', 'Sheet1', 2);  -- Mix of names and numbers
SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('filename.xlsx');
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

typedef struct xlsx_vtab {
    sqlite3_vtab base;
    sqlite3 *db; /* sqlite connection pointer */
} xlsx_vtab;

typedef struct xlsx_cursor {
    sqlite3_vtab_cursor base;
    char **sheet_names;
    int sheet_count;
    int rowid;
} xlsx_cursor;

/* Parser context for Expat */
typedef struct parser_ctx {
    char **names;
    int count;
    int capacity;
    char *error;
} parser_ctx;

/* Forward declarations */
static int xlsxConnect(sqlite3 *db, void *pAux, int argc, const char *const *argv,
                       sqlite3_vtab **ppVtab, char **pzErr);
static int xlsxDisconnect(sqlite3_vtab *pVtab);
static int xlsxOpen(sqlite3_vtab *pVtab, sqlite3_vtab_cursor **ppCursor);
static int xlsxClose(sqlite3_vtab_cursor*);
static int xlsxFilter(sqlite3_vtab_cursor*, int idxNum, const char *idxStr,
                      int argc, sqlite3_value **argv);
static int xlsxNext(sqlite3_vtab_cursor*);
static int xlsxEof(sqlite3_vtab_cursor*);
static int xlsxColumn(sqlite3_vtab_cursor*, sqlite3_context*, int);
static int xlsxRowid(sqlite3_vtab_cursor*, sqlite3_int64*);
static int xlsxBestIndex(sqlite3_vtab *pVTab, sqlite3_index_info*);

/* Module object */
static sqlite3_module xlsxModule = {
    .iVersion     = 0,
    .xCreate      = 0,
    .xConnect     = xlsxConnect,
    .xBestIndex   = xlsxBestIndex,
    .xDisconnect  = xlsxDisconnect,
    .xDestroy     = 0,
    .xOpen        = xlsxOpen,
    .xClose       = xlsxClose,
    .xFilter      = xlsxFilter,
    .xNext        = xlsxNext,
    .xEof         = xlsxEof,
    .xColumn      = xlsxColumn,
    .xRowid       = xlsxRowid,
    .xUpdate      = 0,
    .xBegin       = 0,
    .xSync        = 0,
    .xCommit      = 0,
    .xRollback    = 0,
    .xFindFunction= 0,
    .xRename      = 0,
    .xSavepoint   = 0,
    .xRelease     = 0,
    .xRollbackTo  = 0,
    .xShadowName  = 0,
    /* Newer fields (example) */
    .xIntegrity   = 0,   /* explicitly initialize the newer field(s) */
};

/* Helper: free sheet names in cursor */
static void free_sheet_names(xlsx_cursor *cur){
    if(!cur) return;
    if(cur->sheet_names){
        for(int i=0;i<cur->sheet_count;i++){
            free(cur->sheet_names[i]);
        }
        free(cur->sheet_names);
        cur->sheet_names = NULL;
    }
    cur->sheet_count = 0;
}

/* Helper: append name to parser_ctx */
static int parser_append_name(parser_ctx *ctx, const char *name){
    if(!ctx || !name) return 0;
    if(ctx->count + 1 > ctx->capacity){
        int newcap = ctx->capacity ? ctx->capacity * 2 : 8;
        char **tmp = (char**)realloc(ctx->names, newcap * sizeof(char*));
        if(!tmp) return 0;
        ctx->names = tmp;
        ctx->capacity = newcap;
    }
    ctx->names[ctx->count] = strdup(name);
    if(!ctx->names[ctx->count]) return 0;
    ctx->count++;
    return 1;
}

/* Utility: check if local name equals target (handles possible prefix "prefix:sheet") */
static int local_name_equals(const char *qname, const char *target){
    if(!qname || !target) return 0;
    const char *colon = strrchr(qname, ':');
    const char *local = colon ? colon + 1 : qname;
    return strcmp(local, target) == 0;
}

/* Expat start element handler */
static void XMLCALL start_element(void *userData, const XML_Char *name, const XML_Char **atts){
    parser_ctx *ctx = (parser_ctx*)userData;
    if(!ctx) return;

    /* Look for element named "sheet" (namespace prefixes possible) */
    if(local_name_equals((const char*)name, "sheet")){
        /* find attribute "name" (may be prefixed) */
        for(const XML_Char **a = atts; a && *a; a += 2){
            const char *attr_name = (const char*)a[0];
            const char *attr_val  = (const char*)a[1];
            if(local_name_equals(attr_name, "name")){
                if(!parser_append_name(ctx, attr_val)){
                    /* memory error: record it */
                    if(!ctx->error) ctx->error = strdup("Out of memory while collecting sheet names");
                }
                break;
            }
        }
    }
}

/*
 * parse_sheet_names_from_xlsx_via_zipfile_expat
 * - db: sqlite3 connection (must have zipfile module available)
 * - filename: path to XLSX file
 * - out_names/out_count: outputs (caller must free names)
 * - pzErr: sqlite3_mprintf'd error message on failure (caller must sqlite3_free)
 */
static int parse_sheet_names_from_xlsx_via_zipfile_expat(sqlite3 *db,
                                                         const char *filename,
                                                         char ***out_names,
                                                         int *out_count,
                                                         char **pzErr)
{
    sqlite3_stmt *stmt = NULL;
    const char *sql = "SELECT data FROM zipfile(?) WHERE name = 'xl/workbook.xml' LIMIT 1";
    int rc;
    parser_ctx ctx;
    XML_Parser parser = NULL;

    *out_names = NULL;
    *out_count = 0;
    if(!db || !filename){
        if(pzErr) *pzErr = sqlite3_mprintf("Invalid arguments");
        return SQLITE_MISUSE;
    }

    memset(&ctx, 0, sizeof(ctx));

    rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    if(rc != SQLITE_OK){
        if(pzErr) *pzErr = sqlite3_mprintf("Failed to prepare zipfile query: %s", sqlite3_errmsg(db));
        return rc;
    }

    rc = sqlite3_bind_text(stmt, 1, filename, -1, SQLITE_STATIC);
    if(rc != SQLITE_OK){
        if(pzErr) *pzErr = sqlite3_mprintf("Failed to bind filename: %s", sqlite3_errmsg(db));
        sqlite3_finalize(stmt);
        return rc;
    }

    rc = sqlite3_step(stmt);
    if(rc == SQLITE_ROW){
        const void *blob = sqlite3_column_blob(stmt, 0);
        int blob_size = sqlite3_column_bytes(stmt, 0);
        if(!blob || blob_size <= 0){
            if(pzErr) *pzErr = sqlite3_mprintf("workbook.xml is empty or unreadable");
            sqlite3_finalize(stmt);
            return SQLITE_ERROR;
        }

        /* Create Expat parser */
        parser = XML_ParserCreate(NULL);
        if(!parser){
            if(pzErr) *pzErr = sqlite3_mprintf("Failed to create XML parser");
            sqlite3_finalize(stmt);
            return SQLITE_NOMEM;
        }

        XML_SetUserData(parser, &ctx);
        XML_SetStartElementHandler(parser, start_element);

        /* Parse in one call (workbook.xml is typically small) */
        if(XML_Parse(parser, (const char*)blob, blob_size, XML_TRUE) == XML_STATUS_ERROR){
            enum XML_Error err = XML_GetErrorCode(parser);
            const char *msg = XML_ErrorString(err);
            if(pzErr) *pzErr = sqlite3_mprintf("XML parse error: %s", msg ? msg : "unknown");
            XML_ParserFree(parser);
            sqlite3_finalize(stmt);
            /* free any names collected */
            for(int i=0;i<ctx.count;i++) free(ctx.names[i]);
            free(ctx.names);
            return SQLITE_ERROR;
        }

        if(ctx.error){
            if(pzErr) *pzErr = sqlite3_mprintf("%s", ctx.error);
            free(ctx.error);
            XML_ParserFree(parser);
            sqlite3_finalize(stmt);
            for(int i=0;i<ctx.count;i++) free(ctx.names[i]);
            free(ctx.names);
            return SQLITE_NOMEM;
        }

        /* success */
        *out_names = ctx.names;
        *out_count = ctx.count;

        XML_ParserFree(parser);
    } else {
        if(rc == SQLITE_DONE){
            if(pzErr) *pzErr = sqlite3_mprintf("xl/workbook.xml not found in '%s' (zipfile returned no rows)", filename);
            sqlite3_finalize(stmt);
            return SQLITE_ERROR;
        } else {
            if(pzErr) *pzErr = sqlite3_mprintf("Error reading zipfile: %s", sqlite3_errmsg(db));
            sqlite3_finalize(stmt);
            return rc;
        }
    }

    sqlite3_finalize(stmt);
    return SQLITE_OK;
}

/* xConnect: declare the virtual table schema */
static int xlsxConnect(sqlite3 *db, void *pAux, int argc, const char *const *argv,
                       sqlite3_vtab **ppVtab, char **pzErr){
    (void)pAux; (void)argc; (void)argv; /* silence unused-parameter warning */
    int rc;
    xlsx_vtab *vtab = (xlsx_vtab*)sqlite3_malloc(sizeof(xlsx_vtab));
    if(!vtab) return SQLITE_NOMEM;
    memset(vtab, 0, sizeof(xlsx_vtab));

    vtab->db = db;

    rc = sqlite3_declare_vtab(db, "CREATE TABLE xlsx_import_sheetnames(sheet_num INTEGER, sheet_name TEXT, filename HIDDEN)");
    if(rc != SQLITE_OK){
        sqlite3_free(vtab);
        if(pzErr) *pzErr = sqlite3_mprintf("Failed to declare virtual table");
        return rc;
    }

    *ppVtab = (sqlite3_vtab*)vtab;
    return SQLITE_OK;
}

static int xlsxDisconnect(sqlite3_vtab *pVtab){
    if(pVtab) sqlite3_free(pVtab);
    return SQLITE_OK;
}

static int xlsxOpen(sqlite3_vtab *pVtab, sqlite3_vtab_cursor **ppCursor){
    (void)pVtab; /* silence unused-parameter warning */
    xlsx_cursor *cur = (xlsx_cursor*)sqlite3_malloc(sizeof(xlsx_cursor));
    if(!cur) return SQLITE_NOMEM;
    memset(cur, 0, sizeof(xlsx_cursor));
    *ppCursor = &cur->base;
    return SQLITE_OK;
}

static int xlsxClose(sqlite3_vtab_cursor *pCursor){
    xlsx_cursor *cur = (xlsx_cursor*)pCursor;
    if(cur){
        free_sheet_names(cur);
        sqlite3_free(cur);
    }
    return SQLITE_OK;
}

/* xFilter: start scan; expect filename as first argument */
static int xlsxFilter(sqlite3_vtab_cursor *pCursor, int idxNum, const char *idxStr,
                      int argc, sqlite3_value **argv){
    (void)idxNum; /* silence unused-parameter warning */
    xlsx_cursor *cur = (xlsx_cursor*)pCursor;
    const char *filename = NULL;
    char *err = NULL;
    int rc;

    free_sheet_names(cur);
    cur->rowid = 0;

    if(argc >= 1 && argv && argv[0]){
        if(sqlite3_value_type(argv[0]) != SQLITE_NULL){
            filename = (const char*)sqlite3_value_text(argv[0]);
        }
    }

    if(!filename && idxStr){
        filename = idxStr;
    }

    if(!filename){
        cur->sheet_count = 0;
        return SQLITE_OK;
    }

    xlsx_vtab *vtab = (xlsx_vtab*)pCursor->pVtab;
    if(!vtab || !vtab->db) return SQLITE_MISUSE;

    rc = parse_sheet_names_from_xlsx_via_zipfile_expat(vtab->db, filename, &cur->sheet_names, &cur->sheet_count, &err);
    if(rc != SQLITE_OK){
        if(err){
            sqlite3_log(SQLITE_ERROR, "%s", err);
            sqlite3_free(err);
        }
        cur->sheet_count = 0;
        return rc;
    }

    cur->rowid = 0;
    return SQLITE_OK;
}

static int xlsxNext(sqlite3_vtab_cursor *pCursor){
    xlsx_cursor *cur = (xlsx_cursor*)pCursor;
    cur->rowid++;
    return SQLITE_OK;
}

static int xlsxEof(sqlite3_vtab_cursor *pCursor){
    xlsx_cursor *cur = (xlsx_cursor*)pCursor;
    return cur->rowid >= cur->sheet_count;
}

static int xlsxColumn(sqlite3_vtab_cursor *pCursor, sqlite3_context *ctx, int col){
    xlsx_cursor *cur = (xlsx_cursor*)pCursor;
    if(col == 0){
        sqlite3_result_int(ctx, cur->rowid + 1);
    } else if(col == 1){
        if(cur->rowid < cur->sheet_count && cur->sheet_names[cur->rowid]){
            sqlite3_result_text(ctx, cur->sheet_names[cur->rowid], -1, SQLITE_TRANSIENT);
        } else {
            sqlite3_result_null(ctx);
        }
    } else {
        sqlite3_result_null(ctx);
    }
    return SQLITE_OK;
}

static int xlsxRowid(sqlite3_vtab_cursor *pCursor, sqlite3_int64 *pRowid){
    xlsx_cursor *cur = (xlsx_cursor*)pCursor;
    *pRowid = cur->rowid + 1;
    return SQLITE_OK;
}

static int xlsxBestIndex(sqlite3_vtab *pVTab, sqlite3_index_info *pIdxInfo){
    (void)pVTab;
    int idx = -1;
    /* Look for a constraint on the hidden column (index 2) */
    for(int i=0; i<pIdxInfo->nConstraint; i++){
        if(pIdxInfo->aConstraint[i].iColumn == 2){
            if(pIdxInfo->aConstraint[i].op == SQLITE_INDEX_CONSTRAINT_EQ){
                idx = i;
                break;
            }
        }
    }

    if(idx >= 0){
        pIdxInfo->aConstraintUsage[idx].argvIndex = 1;
        pIdxInfo->aConstraintUsage[idx].omit = 1; 
        pIdxInfo->estimatedCost = 10.0; /* Cheaper if we have the file */
    } else {
        pIdxInfo->estimatedCost = 1000000.0; /* Expensive check without filename */
    }
    return SQLITE_OK;
}

/* Version function */
static void xlsx_import_version(sqlite3_context *context, int argc, sqlite3_value **argv){
    (void)argc; (void)argv;
    sqlite3_result_text(context, "2026-01-07 Copilot Think Deeper (GPT 5.1?)", -1, SQLITE_STATIC);
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
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_import function");
        return rc;
    }
    rc = sqlite3_create_module(db, "xlsx_import_sheetnames", &xlsxModule, 0);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_import_sheetnames module");
        return rc;
    }
    rc = sqlite3_create_function(db, "xlsx_import_version", 0, SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL, xlsx_import_version, NULL, NULL);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_import_version function");
        return rc;
    }
    return rc;
}

/* End of xlsximport.c */
