/*
Prompts used (included per user request):

Prompt 1:
"Create the C code for a SQLite extension named xlsxexport that contains a SQL function named xlsx_export that saves multiple tables as a single XLSX spreadsheet, with the 
sheet names equal to the table names, and the sheet headers in bold and with autofilter. Use the SQLite zipfile extension to handle the ZIP container. Do not use external libraries to handle XML.
The table names are the arguments 2 to N of the xlsx_export function.
Sanitize the sheet names to conform to Excel restrictions. Comply with 32K characters Excel cell limit. 
Add SQL function xlsx_export_version returning "2025-12-30 Copilot Think Deeper (GPT 5.1?)".
Include as comments the prompts used."

Prompt 2:
"Use the SQLite virtual table zipfile extension"

Follow-up:
"worst-case expansion of plain text to XML is quote to '&quot;' (6 chars)"

User request:
"Replace inlineStr usage with a sharedStrings table to reduce file size for repeated strings"

Follow-up user request:
"use sqlite3_snprintf instead of snprintf in lines 389 and 432"

Follow-up user request:
"To insert a directory into the ZIP archive, set the "data" column to NULL when inserting into the zipfile virtual table"

User request:
"Integrate these directory-insert calls into the C code"
*/

/*
xlsxexport_shared_dirs.c

SQLite loadable extension implementing:
  - xlsx_export(filename, table1, table2, ...)
    Writes an XLSX file using the SQLite zipfile virtual table.
    Uses a sharedStrings table to reduce file size.
    Creates explicit directory entries in the ZIP by inserting rows with data == NULL.
    Header row uses bold style (xf index 1) and an autofilter is applied.
    Numeric types are written as numbers; text uses sharedStrings (t="s").
    Text is truncated to Excel's ~32K character limit per cell.

  - xlsx_export_version()
    Returns "2025-12-30 Copilot Think Deeper (GPT 5.1?)".

Notes:
  - Assumes a zipfile virtual table module is available and supports:
      CREATE VIRTUAL TABLE temp._xlsx_zip USING zipfile('archive.xlsx');
      INSERT OR REPLACE INTO temp._xlsx_zip(name, data) VALUES(...);
    If your zipfile vtab differs, adapt zipfile_insert_via_vtab accordingly.
  - No external XML libraries are used.
  - For very large datasets this builds each worksheet and sharedStrings in memory; adapt if streaming is required.
*/

#define _GNU_SOURCE
#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <stdarg.h>

/* Excel limits */
#define EXCEL_SHEETNAME_MAX 31
#define EXCEL_CELL_CHAR_LIMIT 32767

/* ---------- Utility helpers ---------- */

static char *xstrdup(const char *s){
    if(!s) return NULL;
    size_t n = strlen(s) + 1;
    char *p = (char*)malloc(n);
    if(p) memcpy(p, s, n);
    return p;
}

/* XML escape
   Worst-case expansion: '"' -> "&quot;" (6 chars), so allocate len*6 + 1.
   Caller must free returned buffer.
*/
static char *xml_escape(const char *s){
    if(!s) return xstrdup("");
    size_t len = strlen(s);
    size_t cap = len * 6 + 1;
    char *out = (char*)malloc(cap);
    if(!out) return NULL;
    char *p = out;
    for(const unsigned char *q = (const unsigned char*)s; *q; ++q){
        switch(*q){
            case '&': memcpy(p, "&amp;", 5); p += 5; break;
            case '<': memcpy(p, "&lt;", 4); p += 4; break;
            case '>': memcpy(p, "&gt;", 4); p += 4; break;
            case '"': memcpy(p, "&quot;", 6); p += 6; break;
            case '\'': memcpy(p, "&apos;", 6); p += 6; break;
            default:
                if(*q < 0x20 && *q != 0x09 && *q != 0x0A && *q != 0x0D){
                    *p++ = ' ';
                } else {
                    *p++ = *q;
                }
        }
    }
    *p = '\0';
    return out;
}

/* Sanitize sheet name to Excel rules and ensure uniqueness among existing[] */
static char *sanitize_sheet_name(const char *name, int idx, char **existing, int existing_count){
    if(!name) name = "";
    /* trim leading/trailing whitespace */
    while(*name && isspace((unsigned char)*name)) name++;
    size_t len = strlen(name);
    while(len > 0 && isspace((unsigned char)name[len-1])) len--;
    char *tmp = (char*)malloc(len + 1);
    if(!tmp) return NULL;
    memcpy(tmp, name, len);
    tmp[len] = '\0';

    /* remove forbidden characters */
    char *r = tmp, *w = tmp;
    while(*r){
        if(*r == ':' || *r == '\\' || *r == '/' || *r == '?' || *r == '*' || *r == '[' || *r == ']'){
            /* skip */
        } else {
            *w++ = *r;
        }
        r++;
    }
    *w = '\0';

    /* trim leading/trailing single quote */
    if(w - tmp > 0 && tmp[0] == '\'') memmove(tmp, tmp+1, strlen(tmp));
    size_t tlen = strlen(tmp);
    if(tlen > 0 && tmp[tlen-1] == '\'') tmp[tlen-1] = '\0';

    /* default name if empty */
    if(tmp[0] == '\0'){
        free(tmp);
        char buf[64];
        snprintf(buf, sizeof(buf), "Sheet%d", idx+1);
        tmp = xstrdup(buf);
        if(!tmp) return NULL;
    }

    /* truncate to 31 chars */
    if(strlen(tmp) > EXCEL_SHEETNAME_MAX) tmp[EXCEL_SHEETNAME_MAX] = '\0';

    /* ensure uniqueness */
    int suffix = 1;
    char *unique = NULL;
    while(1){
        int conflict = 0;
        for(int i=0;i<existing_count;i++){
            if(existing[i] && strcmp(existing[i], tmp) == 0){ conflict = 1; break; }
        }
        if(!conflict){
            unique = xstrdup(tmp);
            break;
        }
        char buf[128];
        snprintf(buf, sizeof(buf), "%s (%d)", tmp, suffix++);
        if(strlen(buf) > EXCEL_SHEETNAME_MAX) buf[EXCEL_SHEETNAME_MAX] = '\0';
        free(tmp);
        tmp = xstrdup(buf);
        if(!tmp){ unique = NULL; break; }
    }
    free(tmp);
    return unique;
}

/* Convert 0-based column index to Excel letters */
static void col_to_letters(int col, char *out, size_t outsz){
    char buf[16];
    int pos = 0;
    int v = col + 1;
    while(v > 0 && pos < (int)sizeof(buf)-1){
        int rem = (v - 1) % 26;
        buf[pos++] = 'A' + rem;
        v = (v - 1) / 26;
    }
    int j = 0;
    for(int i = pos-1; i >= 0 && j < (int)outsz-1; --i) out[j++] = buf[i];
    out[j] = '\0';
}

/* Simple dynamic writer */
typedef struct {
    char *data;
    size_t len;
    size_t cap;
} memwriter;

static int mw_init(memwriter *mw){
    mw->cap = 8192;
    mw->len = 0;
    mw->data = (char*)malloc(mw->cap);
    if(!mw->data) return SQLITE_NOMEM;
    mw->data[0] = '\0';
    return SQLITE_OK;
}
static void mw_free(memwriter *mw){
    if(mw->data) free(mw->data);
    mw->data = NULL; mw->len = mw->cap = 0;
}
static int mw_append(memwriter *mw, const char *fmt, ...){
    va_list ap;
    va_start(ap, fmt);
    int needed = vsnprintf(NULL, 0, fmt, ap);
    va_end(ap);
    if(needed < 0) return SQLITE_ERROR;
    if(mw->len + (size_t)needed + 1 > mw->cap){
        size_t newcap = mw->cap * 2;
        while(mw->len + (size_t)needed + 1 > newcap) newcap *= 2;
        char *p = (char*)realloc(mw->data, newcap);
        if(!p) return SQLITE_NOMEM;
        mw->data = p; mw->cap = newcap;
    }
    va_start(ap, fmt);
    vsnprintf(mw->data + mw->len, mw->cap - mw->len, fmt, ap);
    va_end(ap);
    mw->len += (size_t)needed;
    return SQLITE_OK;
}

/* ---------- sharedStrings table ---------- */

typedef struct {
    char **items;
    int count;
    int cap;
} sst;

static void sst_init(sst *s){ s->items = NULL; s->count = s->cap = 0; }
static void sst_free(sst *s){
    for(int i=0;i<s->count;i++) free(s->items[i]);
    free(s->items); s->items = NULL; s->count = s->cap = 0;
}

/* Return index of string, adding if necessary. Stores raw (unescaped) text. */
static int sst_index(sst *s, const char *txt){
    if(!txt) txt = "";
    for(int i=0;i<s->count;i++){
        if(strcmp(s->items[i], txt) == 0) return i;
    }
    if(s->count >= s->cap){
        int ncap = s->cap ? s->cap*2 : 64;
        char **p = (char**)realloc(s->items, ncap * sizeof(char*));
        if(!p) return -1;
        s->items = p; s->cap = ncap;
    }
    s->items[s->count] = xstrdup(txt);
    if(!s->items[s->count]) return -1;
    return s->count++;
}

/* Build sharedStrings.xml from sst */
static char *build_sharedstrings_xml(sst *s){
    int uniqueCount = s->count;
    size_t cap = 1024 + uniqueCount * 64;
    for(int i=0;i<uniqueCount;i++){
        size_t sl = strlen(s->items[i]);
        cap += sl * 6 + 64;
    }
    memwriter mw;
    if(mw_init(&mw) != SQLITE_OK) return NULL;
    mw_append(&mw, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    mw_append(&mw, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"%d\" uniqueCount=\"%d\">\n", uniqueCount, uniqueCount);
    for(int i=0;i<uniqueCount;i++){
        char *escaped = xml_escape(s->items[i]);
        mw_append(&mw, "  <si><t>%s</t></si>\n", escaped);
        free(escaped);
    }
    mw_append(&mw, "</sst>");
    char *out = mw.data;
    out[mw.len] = '\0';
    return out;
}

/* ---------- XML parts builders ---------- */

static char *build_styles_xml(void){
    const char *styles =
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
"  <fonts count=\"2\">"
"    <font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font>"
"    <font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font>"
"  </fonts>"
"  <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
"  <borders count=\"1\"><border/></borders>"
"  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
"  <cellXfs count=\"2\">"
"    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
"    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
"  </cellXfs>"
"  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
"</styleSheet>";
    return xstrdup(styles);
}

static char *build_content_types_xml(int sheet_count, int include_sharedstrings){
    size_t cap = 4096 + sheet_count * 256;
    char *buf = (char*)malloc(cap);
    if(!buf) return NULL;
    strcpy(buf,
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
"  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
"  <Default Extension=\"xml\" ContentType=\"application/xml\"/>"
"  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
"  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    if(include_sharedstrings){
        strcat(buf, "  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
    }
    for(int i=0;i<sheet_count;i++){
        char part[256];
        snprintf(part, sizeof(part),
                 "  <Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                 i+1);
        strcat(buf, part);
    }
    strcat(buf,
"  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
"  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
"</Types>");
    return buf;
}

static const char *rels_rels =
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
"  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\"/>"
"  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"/docProps/core.xml\"/>"
"  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"/docProps/app.xml\"/>"
"</Relationships>";

static const char *docprops_core =
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
"  <dc:creator>sqlite3 xlsxexport</dc:creator>"
"  <cp:lastModifiedBy>sqlite3 xlsxexport</cp:lastModifiedBy>"
"  <dcterms:created xsi:type=\"dcterms:W3CDTF\">2025-12-30T00:00:00Z</dcterms:created>"
"  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">2025-12-30T00:00:00Z</dcterms:modified>"
"</cp:coreProperties>";

static const char *docprops_app =
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
"  <Application>Microsoft Excel</Application>"
"</Properties>";

static char *build_workbook_rels(int sheet_count){
    size_t cap = 1024 + sheet_count * 180;
    char *buf = (char*)malloc(cap);
    if(!buf) return NULL;
    strcpy(buf,
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for(int i=0;i<sheet_count;i++){
        char part[180];
        snprintf(part, sizeof(part),
                 "  <Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>",
                 i+1, i+1);
        strcat(buf, part);
    }
    strcat(buf,
"  <Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
"</Relationships>");
    return buf;
}

static char *build_workbook_xml(char **sheet_names, int sheet_count){
    size_t cap = 1024 + sheet_count * 256;
    char *buf = (char*)malloc(cap);
    if(!buf) return NULL;
    strcpy(buf,
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
"  <sheets>");
    for(int i=0;i<sheet_count;i++){
        char part[512];
        snprintf(part, sizeof(part),
                 "    <sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>",
                 sheet_names[i], i+1, i+1);
        strcat(buf, part);
    }
    strcat(buf,
"  </sheets>"
"</workbook>");
    return buf;
}

/* ---------- Worksheet builder (uses sharedStrings) ---------- */

/* Build worksheet XML for a given table name. header_style_index is the xf index for header (1).
   sst is updated with any strings encountered; worksheet XML references shared string indices (t="s").
*/
static int build_worksheet_xml_with_sst(sqlite3 *db, const char *table, int header_style_index, sst *sstp, memwriter *out, char **pzErrMsg){
    int rc;
    sqlite3_stmt *pPragma = NULL;
    char pragma_sql[256];
    /* Use sqlite3_snprintf for SQL safety as requested */
    sqlite3_snprintf(sizeof(pragma_sql), pragma_sql, "PRAGMA table_info(%Q);", table);
    rc = sqlite3_prepare_v2(db, pragma_sql, -1, &pPragma, NULL);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("PRAGMA table_info failed for %s: %s", table, sqlite3_errmsg(db));
        return rc;
    }
    char **colnames = NULL;
    int colcount = 0;
    while((rc = sqlite3_step(pPragma)) == SQLITE_ROW){
        const unsigned char *cname = sqlite3_column_text(pPragma, 1);
        colnames = (char**)realloc(colnames, sizeof(char*) * (colcount + 1));
        colnames[colcount] = xstrdup((const char*)cname);
        colcount++;
    }
    sqlite3_finalize(pPragma);
    if(rc != SQLITE_DONE){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed reading PRAGMA table_info for %s: %s", table, sqlite3_errmsg(db));
        for(int i=0;i<colcount;i++) free(colnames[i]);
        free(colnames);
        return rc;
    }

    if(colcount == 0){
        mw_init(out);
        mw_append(out, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        mw_append(out, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
        mw_append(out, "<sheetData/>\n</worksheet>");
        for(int i=0;i<colcount;i++) free(colnames[i]);
        free(colnames);
        return SQLITE_OK;
    }

    /* Build SELECT */
    size_t qcap = 256 + colcount * 64;
    char *sql = (char*)malloc(qcap);
    if(!sql){
        for(int i=0;i<colcount;i++) free(colnames[i]);
        free(colnames);
        return SQLITE_NOMEM;
    }
    strcpy(sql, "SELECT ");
    for(int i=0;i<colcount;i++){
        char part[256];
        sqlite3_snprintf(sizeof(part), part, "%s%Q", (i==0)?"":", ", colnames[i]);
        strncat(sql, part, qcap - strlen(sql) - 1);
    }
    strncat(sql, " FROM ", qcap - strlen(sql) - 1);
    strncat(sql, table, qcap - strlen(sql) - 1);

    sqlite3_stmt *pStmt = NULL;
    rc = sqlite3_prepare_v2(db, sql, -1, &pStmt, NULL);
    free(sql);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to prepare SELECT for %s: %s", table, sqlite3_errmsg(db));
        for(int i=0;i<colcount;i++) free(colnames[i]);
        free(colnames);
        return rc;
    }

    rc = mw_init(out);
    if(rc != SQLITE_OK){
        sqlite3_finalize(pStmt);
        for(int i=0;i<colcount;i++) free(colnames[i]);
        free(colnames);
        return rc;
    }
    mw_append(out, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    mw_append(out, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
    mw_append(out, "<sheetData>\n");

    /* Header row: use sharedStrings for header text */
    mw_append(out, "<row r=\"1\">\n");
    for(int c=0;c<colcount;c++){
        char collet[8];
        col_to_letters(c, collet, sizeof(collet));
        char cellref[32];
        snprintf(cellref, sizeof(cellref), "%s1", collet);
        /* header raw text */
        const char *raw = colnames[c] ? colnames[c] : "";
        /* truncate header if needed */
        size_t rlen = strlen(raw);
        char *hdr = NULL;
        if(rlen > EXCEL_CELL_CHAR_LIMIT){
            hdr = (char*)malloc(EXCEL_CELL_CHAR_LIMIT + 1);
            if(hdr){
                memcpy(hdr, raw, EXCEL_CELL_CHAR_LIMIT);
                hdr[EXCEL_CELL_CHAR_LIMIT] = '\0';
            } else {
                hdr = xstrdup("");
            }
        } else {
            hdr = xstrdup(raw);
        }
        int idx = sst_index(sstp, hdr);
        free(hdr);
        if(idx < 0){
            sqlite3_finalize(pStmt);
            for(int i=0;i<colcount;i++) free(colnames[i]);
            free(colnames);
            return SQLITE_NOMEM;
        }
        mw_append(out, "<c r=\"%s\" s=\"%d\" t=\"s\"><v>%d</v></c>", cellref, header_style_index, idx);
    }
    mw_append(out, "\n</row>\n");

    /* Autofilter */
    char lastcol[8];
    col_to_letters(colcount - 1, lastcol, sizeof(lastcol));
    mw_append(out, "<autoFilter ref=\"A1:%s1\"/>\n", lastcol);

    /* Data rows */
    int rownum = 2;
    while((rc = sqlite3_step(pStmt)) == SQLITE_ROW){
        mw_append(out, "<row r=\"%d\">", rownum);
        for(int c=0;c<colcount;c++){
            int coltype = sqlite3_column_type(pStmt, c);
            char collet2[8];
            col_to_letters(c, collet2, sizeof(collet2));
            char cellref2[32];
            snprintf(cellref2, sizeof(cellref2), "%s%d", collet2, rownum);
            if(coltype == SQLITE_INTEGER){
                long long v = sqlite3_column_int64(pStmt, c);
                mw_append(out, "<c r=\"%s\"><v>%lld</v></c>", cellref2, v);
            } else if(coltype == SQLITE_FLOAT){
                double d = sqlite3_column_double(pStmt, c);
                char numbuf[64];
                snprintf(numbuf, sizeof(numbuf), "%.15g", d);
                mw_append(out, "<c r=\"%s\"><v>%s</v></c>", cellref2, numbuf);
            } else if(coltype == SQLITE_NULL){
                mw_append(out, "<c r=\"%s\"/>", cellref2);
            } else {
                const unsigned char *txt = sqlite3_column_text(pStmt, c);
                if(!txt) txt = (const unsigned char*)"";
                size_t tlen = strlen((const char*)txt);
                char *rawtxt = NULL;
                if(tlen > EXCEL_CELL_CHAR_LIMIT){
                    rawtxt = (char*)malloc(EXCEL_CELL_CHAR_LIMIT + 1);
                    if(rawtxt){
                        memcpy(rawtxt, txt, EXCEL_CELL_CHAR_LIMIT);
                        rawtxt[EXCEL_CELL_CHAR_LIMIT] = '\0';
                    } else {
                        rawtxt = xstrdup("");
                    }
                } else {
                    rawtxt = xstrdup((const char*)txt);
                }
                int idx = sst_index(sstp, rawtxt);
                free(rawtxt);
                if(idx < 0){
                    sqlite3_finalize(pStmt);
                    for(int i=0;i<colcount;i++) free(colnames[i]);
                    free(colnames);
                    return SQLITE_NOMEM;
                }
                mw_append(out, "<c r=\"%s\" t=\"s\"><v>%d</v></c>", cellref2, idx);
            }
        }
        mw_append(out, "</row>\n");
        rownum++;
        rc = SQLITE_OK; /* continue */
    }
    sqlite3_finalize(pStmt);

    mw_append(out, "</sheetData>\n</worksheet>");

    for(int i=0;i<colcount;i++) free(colnames[i]);
    free(colnames);
    return SQLITE_OK;
}

/* ---------- Zipfile virtual table insertion helper ---------- */

/*
 Insert a file or directory into the zip archive using the zipfile virtual table.
 If data == NULL, a directory entry is created (the zipfile vtab expects NULL in the data column).
*/
static int zipfile_insert_via_vtab(sqlite3 *db, const char *archive_filename, const char *path_in_zip, const void *data, sqlite3_int64 nData, char **pzErrMsg){
    int rc;
    sqlite3_stmt *pInsert = NULL;
    char *zCreate = NULL;
    char *zSql = NULL;
    const char *vtabName = "_xlsx_zip_vtab";

    fprintf(stderr, "zipfile_insert_via_vtab() %s %s %p %lld\n", archive_filename, path_in_zip, data, nData);
    zCreate = sqlite3_mprintf("CREATE VIRTUAL TABLE IF NOT EXISTS temp.%s USING zipfile(%Q);", vtabName, archive_filename);
    if(!zCreate) return SQLITE_NOMEM;

    rc = sqlite3_exec(db, "SAVEPOINT xlsx_export_sp;", NULL, NULL, NULL);
    if(rc != SQLITE_OK){
        sqlite3_free(zCreate);
        return rc;
    }

    rc = sqlite3_exec(db, zCreate, NULL, NULL, pzErrMsg);
    sqlite3_free(zCreate);
    if(rc != SQLITE_OK){
        sqlite3_exec(db, "ROLLBACK TO xlsx_export_sp;", NULL, NULL, NULL);
        sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, NULL);
        return rc;
    }

    //zSql = sqlite3_mprintf("INSERT OR REPLACE INTO temp.%s(name, data) VALUES(?, ?);", vtabName);
    zSql = sqlite3_mprintf("INSERT INTO temp.%s(name, data) VALUES(?, ?);", vtabName);
    if(!zSql){
        sqlite3_exec(db, "ROLLBACK TO xlsx_export_sp;", NULL, NULL, NULL);
        sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, NULL);
        return SQLITE_NOMEM;
    }

    rc = sqlite3_prepare_v2(db, zSql, -1, &pInsert, NULL);
    sqlite3_free(zSql);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to prepare zipfile insert: %s", sqlite3_errmsg(db));
        sqlite3_exec(db, "ROLLBACK TO xlsx_export_sp;", NULL, NULL, NULL);
        sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, NULL);
        return rc;
    }

    rc = sqlite3_bind_text(pInsert, 1, path_in_zip, -1, SQLITE_STATIC);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to bind name: %s", sqlite3_errmsg(db));
        sqlite3_finalize(pInsert);
        sqlite3_exec(db, "ROLLBACK TO xlsx_export_sp;", NULL, NULL, NULL);
        sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, NULL);
        return rc;
    }

    /* If data is NULL, bind NULL to create a directory entry; otherwise bind blob */
    if(data == NULL){
        rc = sqlite3_bind_null(pInsert, 2);
    } else {
        rc = sqlite3_bind_blob(pInsert, 2, data, (int)nData, SQLITE_TRANSIENT);
    }
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to bind data: %s", sqlite3_errmsg(db));
        sqlite3_finalize(pInsert);
        sqlite3_exec(db, "ROLLBACK TO xlsx_export_sp;", NULL, NULL, NULL);
        sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, NULL);
        return rc;
    }

    rc = sqlite3_step(pInsert);
    if(rc != SQLITE_DONE){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to insert into zipfile vtab: %s", sqlite3_errmsg(db));
        sqlite3_finalize(pInsert);
        sqlite3_exec(db, "ROLLBACK TO xlsx_export_sp;", NULL, NULL, NULL);
        sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, NULL);
        return rc;
    }
    sqlite3_finalize(pInsert);

    /* Best-effort drop of temp vtab */
    char *zDrop = sqlite3_mprintf("DROP TABLE IF EXISTS temp.%s;", vtabName);
    if(zDrop){
        sqlite3_exec(db, zDrop, NULL, NULL, NULL);
        sqlite3_free(zDrop);
    }

    rc = sqlite3_exec(db, "RELEASE xlsx_export_sp;", NULL, NULL, pzErrMsg);
    return rc == SQLITE_OK ? SQLITE_OK : rc;
}

/* ---------- xlsx_export function (uses sharedStrings and creates directories) ---------- */

static void xlsx_export_func(sqlite3_context *context, int argc, sqlite3_value **argv){
    sqlite3 *db = sqlite3_context_db_handle(context);
    if(argc < 2){
        sqlite3_result_error(context, "Usage: xlsx_export(filename, table1, table2, ...)", -1);
        return;
    }
    const unsigned char *filename = sqlite3_value_text(argv[0]);
    if(!filename){
        sqlite3_result_error(context, "filename must be a text value", -1);
        return;
    }

    int sheet_count = argc - 1;
    char **raw_table_names = (char**)malloc(sizeof(char*) * sheet_count);
    if(!raw_table_names){ sqlite3_result_error_nomem(context); return; }
    for(int i=0;i<sheet_count;i++){
        const unsigned char *t = sqlite3_value_text(argv[i+1]);
        raw_table_names[i] = t ? xstrdup((const char*)t) : xstrdup("");
    }

    /* Sanitize sheet names */
    char **sheet_names = (char**)malloc(sizeof(char*) * sheet_count);
    if(!sheet_names){ for(int i=0;i<sheet_count;i++) free(raw_table_names[i]); free(raw_table_names); sqlite3_result_error_nomem(context); return; }
    for(int i=0;i<sheet_count;i++){
        sheet_names[i] = sanitize_sheet_name(raw_table_names[i], i, sheet_names, i);
        if(!sheet_names[i]){
            for(int j=0;j<=i;j++) if(sheet_names[j]) free(sheet_names[j]);
            for(int j=0;j<sheet_count;j++) free(raw_table_names[j]);
            free(raw_table_names); free(sheet_names);
            sqlite3_result_error_nomem(context);
            return;
        }
    }

    char *pzErrMsg = NULL;
    int rc;

    /* Prepare sharedStrings structure */
    sst shared;
    sst_init(&shared);

    /* Build worksheets and populate sharedStrings */
    memwriter *worksheets = (memwriter*)malloc(sizeof(memwriter) * sheet_count);
    if(!worksheets){ rc = SQLITE_NOMEM; goto cleanup_error; }
    for(int i=0;i<sheet_count;i++){
        if(mw_init(&worksheets[i]) != SQLITE_OK){ rc = SQLITE_NOMEM; goto cleanup_error; }
        rc = build_worksheet_xml_with_sst(db, raw_table_names[i], 1, &shared, &worksheets[i], &pzErrMsg);
        if(rc != SQLITE_OK){
            goto cleanup_error;
        }
    }

    /* Insert explicit directory entries into the ZIP (data == NULL).
       Many ZIP readers infer directories from file entries, but some consumers
       expect explicit directory entries. We create the directories used by XLSX.
    */
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "_rels/", NULL, 0, &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "docProps/", NULL, 0, &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/", NULL, 0, &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/_rels/", NULL, 0, &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/worksheets/", NULL, 0, &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 1) [Content_Types].xml (include sharedStrings) */
    char *content_types = build_content_types_xml(sheet_count, 1);
    if(!content_types){ rc = SQLITE_NOMEM; goto cleanup_error; }
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "[Content_Types].xml", content_types, (sqlite3_int64)strlen(content_types), &pzErrMsg);
    free(content_types);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 2) _rels/.rels */
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "_rels/.rels", rels_rels, (sqlite3_int64)strlen(rels_rels), &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 3) docProps */
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "docProps/core.xml", docprops_core, (sqlite3_int64)strlen(docprops_core), &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "docProps/app.xml", docprops_app, (sqlite3_int64)strlen(docprops_app), &pzErrMsg);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 4) xl/styles.xml */
    char *styles_xml = build_styles_xml();
    if(!styles_xml){ rc = SQLITE_NOMEM; goto cleanup_error; }
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/styles.xml", styles_xml, (sqlite3_int64)strlen(styles_xml), &pzErrMsg);
    free(styles_xml);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 5) xl/sharedStrings.xml */
    char *shared_xml = build_sharedstrings_xml(&shared);
    if(!shared_xml){ rc = SQLITE_NOMEM; goto cleanup_error; }
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/sharedStrings.xml", shared_xml, (sqlite3_int64)strlen(shared_xml), &pzErrMsg);
    free(shared_xml);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 6) xl/worksheets/sheetN.xml for each sheet */
    for(int i=0;i<sheet_count;i++){
        char path[180];
        sqlite3_snprintf(sizeof(path), path, "xl/worksheets/sheet%d.xml", i+1);
        rc = zipfile_insert_via_vtab(db, (const char*)filename, path, worksheets[i].data, (sqlite3_int64)worksheets[i].len, &pzErrMsg);
        if(rc != SQLITE_OK) goto cleanup_error;
    }

    /* 7) xl/_rels/workbook.xml.rels */
    char *workbook_rels = build_workbook_rels(sheet_count);
    if(!workbook_rels){ rc = SQLITE_NOMEM; goto cleanup_error; }
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/_rels/workbook.xml.rels", workbook_rels, (sqlite3_int64)strlen(workbook_rels), &pzErrMsg);
    free(workbook_rels);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* 8) xl/workbook.xml */
    char *workbook_xml = build_workbook_xml(sheet_names, sheet_count);
    if(!workbook_xml){ rc = SQLITE_NOMEM; goto cleanup_error; }
    rc = zipfile_insert_via_vtab(db, (const char*)filename, "xl/workbook.xml", workbook_xml, (sqlite3_int64)strlen(workbook_xml), &pzErrMsg);
    free(workbook_xml);
    if(rc != SQLITE_OK) goto cleanup_error;

    /* success */
    for(int i=0;i<sheet_count;i++){
        mw_free(&worksheets[i]);
        free(raw_table_names[i]);
        free(sheet_names[i]);
    }
    free(worksheets);
    free(raw_table_names);
    free(sheet_names);
    sst_free(&shared);
    sqlite3_result_int(context, 0);
    return;

cleanup_error:
    if(pzErrMsg){
        sqlite3_result_error(context, pzErrMsg, -1);
        sqlite3_free(pzErrMsg);
    } else {
        sqlite3_result_error(context, sqlite3_errstr(rc), -1);
    }
    if(sheet_count){
        for(int i=0;i<sheet_count;i++){
            if(raw_table_names && raw_table_names[i]) free(raw_table_names[i]);
            if(sheet_names && sheet_names[i]) free(sheet_names[i]);
        }
    }
    if(raw_table_names) free(raw_table_names);
    if(sheet_names) free(sheet_names);
    if(sheet_count && worksheets){
        for(int i=0;i<sheet_count;i++) mw_free(&worksheets[i]);
        free(worksheets);
    }
    sst_free(&shared);
    return;
}

/* ---------- xlsx_export_version ---------- */

static void xlsx_export_version(sqlite3_context *context, int argc, sqlite3_value **argv){
    (void)argc; (void)argv;
    sqlite3_result_text(context, "2025-12-30 Copilot Think Deeper (GPT 5.1?)", -1, SQLITE_STATIC);
}

/* ---------- Extension entry point ---------- */

#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsxexport_init(sqlite3 *db, char **pzErrMsg, const sqlite3_api_routines *pApi){
    SQLITE_EXTENSION_INIT2(pApi);
    int rc;
    rc = sqlite3_create_function(db, "xlsx_export", -1, SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL, xlsx_export_func, NULL, NULL);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_export: %s", sqlite3_errmsg(db));
        return rc;
    }
    rc = sqlite3_create_function(db, "xlsx_export_version", 0, SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL, xlsx_export_version, NULL, NULL);
    if(rc != SQLITE_OK){
        if(pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_export_version: %s", sqlite3_errmsg(db));
        return rc;
    }
    return SQLITE_OK;
}
