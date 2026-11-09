/*
** xlsxexport.c
**
** SQLite extension implementing:
**
**   xlsx_export(filename [, table1, table2, ...])
**   xlsx_export_version()
**
** xlsx_export() writes selected tables into one XLSX file.
** If only a filename is supplied, all normal tables in the main
** schema are exported.
**
** The XLSX file is written as a ZIP archive using the SQLite
** zipfile extension. All XML parts are generated directly in C.
**
** Features:
**   - Sheet names are derived from table names and sanitized
**     for Excel.
**   - Header rows are bold.
**   - Autofilter is enabled on the header row.
**   - Excel's 32,767 character cell limit is enforced; longer
**     cells are truncated and a warning is logged.
**
** Version string:
**   2026-11-03 Qwen3.8-Max (Thinking)
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <string.h>
#include <stdarg.h>
#include <stdio.h>
#include <time.h>

#ifndef SQLITE_DIRECTONLY
#define SQLITE_DIRECTONLY 0
#endif

#ifndef SQLITE_DETERMINISTIC
#define SQLITE_DETERMINISTIC 0
#endif

/*
** Excel limits / constants.
*/
#define XLSX_MAX_CELL_CHARS   32767
#define XLSX_MAX_COLS         16384
#define XLSX_MAX_ROWS         1048576
#define XLSX_FILE_MODE        33188   /* 0100644 */

/*
** Small growable string buffer.
*/
typedef struct XStr {
  char *z;
  int n;
  int nAlloc;
  int err;
} XStr;

typedef struct XlsxTable {
  char *zTable;
  char *zSheet;
} XlsxTable;

typedef struct XlsxTableList {
  XlsxTable *a;
  int n;
  int nAlloc;
} XlsxTableList;

/*
** ------------------------------------------------------------------
** Utility functions
** ------------------------------------------------------------------
*/

static void xstr_init(XStr *p){
  p->z = 0;
  p->n = 0;
  p->nAlloc = 0;
  p->err = 0;
}

static void xstr_free(XStr *p){
  if( p ){
    sqlite3_free(p->z);
    p->z = 0;
    p->n = 0;
    p->nAlloc = 0;
    p->err = 0;
  }
}

static int xstr_reserve(XStr *p, int nExtra){
  int nNew;
  char *zNew;

  if( p->err ) return 1;
  if( p->n + nExtra + 1 <= p->nAlloc ) return 0;

  nNew = (p->nAlloc ? p->nAlloc : 256);
  while( nNew < p->n + nExtra + 1 ){
    if( nNew > 0x3fffffff ){
      p->err = 1;
      return 1;
    }
    nNew *= 2;
  }

  zNew = sqlite3_realloc(p->z, nNew);
  if( zNew==0 ){
    p->err = 1;
    return 1;
  }

  p->z = zNew;
  p->nAlloc = nNew;
  return 0;
}

static int xstr_append(XStr *p, const char *z, int n){
  if( p->err ) return 1;
  if( z==0 ) return 0;
  if( n<0 ) n = (int)strlen(z);

  if( xstr_reserve(p, n) ) return 1;

  memcpy(p->z + p->n, z, n);
  p->n += n;
  p->z[p->n] = 0;
  return 0;
}

static int xstr_appendall(XStr *p, const char *z){
  return xstr_append(p, z, -1);
}

static int xstr_appendchar(XStr *p, char c){
  return xstr_append(p, &c, 1);
}

static int xstr_appendf(XStr *p, const char *zFmt, ...){
  va_list ap;
  char *z;

  if( p->err ) return 1;

  va_start(ap, zFmt);
  z = sqlite3_vmprintf(zFmt, ap);
  va_end(ap);

  if( z==0 ){
    p->err = 1;
    return 1;
  }

  xstr_append(p, z, -1);
  sqlite3_free(z);
  return p->err;
}

static char *xlsx_strndup(const char *z, int n){
  char *p;
  if( z==0 ) return 0;
  p = sqlite3_malloc(n + 1);
  if( p ){
    memcpy(p, z, n);
    p[n] = 0;
  }
  return p;
}

static char *xlsx_strdup(const char *z){
  if( z==0 ) return 0;
  return xlsx_strndup(z, (int)strlen(z));
}

static int utf8_byte_len(unsigned char c){
  if( c < 0x80 ) return 1;
  if( (c & 0xE0)==0xC0 ) return 2;
  if( (c & 0xF0)==0xE0 ) return 3;
  if( (c & 0xF8)==0xF0 ) return 4;
  return 1;
}

static void trim_space_apos(char *z){
  int n;
  int i = 0;

  if( z==0 ) return;
  n = (int)strlen(z);

  while( i<n && (z[i]==' ' || z[i]=='\'') ){
    i++;
  }

  if( i>0 ){
    memmove(z, z + i, n - i + 1);
    n -= i;
  }

  while( n>0 && (z[n-1]==' ' || z[n-1]=='\'') ){
    n--;
  }

  z[n] = 0;
}

/*
** Duplicate at most maxChars UTF-8 characters.
** If the result is empty after trimming, return "Sheet".
*/
static char *xlsx_dup_utf8_maxchars(const char *z, int maxChars){
  int n;
  int i = 0;
  int chars = 0;
  char *p;

  if( z==0 ) return 0;
  if( maxChars <= 0 ) return xlsx_strdup("");

  n = (int)strlen(z);

  while( i<n && chars<maxChars ){
    unsigned char c = (unsigned char)z[i];
    int len;

    if( (c & 0xC0) != 0x80 ){
      len = utf8_byte_len(c);
      if( i + len > n ) len = n - i;
    }else{
      len = 1;
    }

    i += len;
    chars++;
  }

  p = xlsx_strndup(z, i);
  trim_space_apos(p);

  if( p && p[0]==0 ){
    sqlite3_free(p);
    p = xlsx_strdup("Sheet");
  }

  return p;
}

/*
** Characters that are unsafe or risky for Excel sheet names.
** Excel officially forbids: : \ / ? * [ ]
** We additionally sanitize controls and a few risky characters.
*/
static int is_bad_sheet_char(unsigned char c){
  return
     c==':'  || c=='\\' || c=='/' || c=='?' ||
     c=='*'  || c=='['  || c==']' || c=='<' ||
     c=='>'  || c=='"'  || c=='|';
}

/*
** Sanitize a candidate sheet name.
*/
static char *sanitize_sheet_base(const char *zIn){
  XStr s;
  char *z;
  char *z2;
  int i, n;

  xstr_init(&s);

  if( zIn==0 ) zIn = "";
  n = (int)strlen(zIn);

  for(i=0; i<n; i++){
    unsigned char c = (unsigned char)zIn[i];

    if( c < 0x20 || c==0x7f ){
      continue;
    }

    if( is_bad_sheet_char(c) ){
      xstr_appendall(&s, "_");
    }else{
      xstr_append(&s, (const char*)&c, 1);
    }
  }

  if( s.err ){
    xstr_free(&s);
    return 0;
  }

  if( s.n==0 ){
    z = xlsx_strdup("Sheet");
  }else{
    z = xlsx_strndup(s.z, s.n);
  }

  xstr_free(&s);

  if( z==0 ) return 0;

  trim_space_apos(z);

  if( z[0]==0 ){
    sqlite3_free(z);
    z = xlsx_strdup("Sheet");
    if( z==0 ) return 0;
  }

  z2 = xlsx_dup_utf8_maxchars(z, 31);
  sqlite3_free(z);

  if( z2==0 ) return 0;

  /*
  ** "History" is reserved by Excel in some scenarios.
  */
  if( sqlite3_stricmp(z2, "History")==0 ){
    char *z3 = sqlite3_mprintf("%s_", z2);
    sqlite3_free(z2);
    z2 = z3;
  }

  return z2;
}

static int sheet_name_exists(XlsxTable *a, int n, const char *z){
  int i;
  if( z==0 ) return 0;

  for(i=0; i<n; i++){
    if( a[i].zSheet && sqlite3_stricmp(a[i].zSheet, z)==0 ){
      return 1;
    }
  }

  return 0;
}

/*
** Create a unique Excel-compatible sheet name from a table name.
*/
static char *make_unique_sheet(const char *zTable, XlsxTable *a, int nExisting){
  char *zBase;
  int k;

  zBase = sanitize_sheet_base(zTable);
  if( zBase==0 ) return 0;

  if( !sheet_name_exists(a, nExisting, zBase) ){
    return zBase;
  }

  for(k=2; k<1000000; k++){
    char zSuffix[24];
    int nSuffix;
    int maxBase;
    char *zShort;
    char *zNew;

    sqlite3_snprintf(sizeof(zSuffix), zSuffix, "_%d", k);
    nSuffix = (int)strlen(zSuffix);
    maxBase = 31 - nSuffix;

    if( maxBase > 0 ){
      zShort = xlsx_dup_utf8_maxchars(zBase, maxBase);
    }else{
      zShort = xlsx_strdup("");
    }

    if( zShort==0 ){
      sqlite3_free(zBase);
      return 0;
    }

    zNew = sqlite3_mprintf("%s%s", zShort, zSuffix);
    sqlite3_free(zShort);

    if( zNew==0 ){
      sqlite3_free(zBase);
      return 0;
    }

    if( !sheet_name_exists(a, nExisting, zNew) ){
      sqlite3_free(zBase);
      return zNew;
    }

    sqlite3_free(zNew);
  }

  sqlite3_free(zBase);
  return xlsx_strdup("Sheet");
}

static void list_init(XlsxTableList *p){
  p->a = 0;
  p->n = 0;
  p->nAlloc = 0;
}

static int list_add(XlsxTableList *p, const char *zTable){
  if( p->n >= p->nAlloc ){
    int nNew = (p->nAlloc ? p->nAlloc * 2 : 8);
    XlsxTable *a = sqlite3_realloc(p->a, nNew * sizeof(XlsxTable));
    if( a==0 ) return 1;
    p->a = a;
    p->nAlloc = nNew;
  }

  p->a[p->n].zTable = xlsx_strdup(zTable);
  p->a[p->n].zSheet = 0;

  if( p->a[p->n].zTable==0 ) return 1;

  p->n++;
  return 0;
}

static void list_free(XlsxTableList *p){
  int i;

  if( p==0 ) return;

  for(i=0; i<p->n; i++){
    sqlite3_free(p->a[i].zTable);
    sqlite3_free(p->a[i].zSheet);
  }

  sqlite3_free(p->a);
  list_init(p);
}

/*
** Quote an SQL identifier with double quotes.
*/
static char *quote_ident(const char *z){
  XStr s;

  xstr_init(&s);

  if( z==0 ) z = "";

  xstr_appendchar(&s, '"');

  while( *z ){
    if( *z=='"' ){
      xstr_appendall(&s, "\"\"");
    }
    xstr_appendchar(&s, *z);
    z++;
  }

  xstr_appendchar(&s, '"');

  if( s.err ){
    xstr_free(&s);
    return 0;
  }

  return s.z;
}

/*
** ------------------------------------------------------------------
** XML helpers
** ------------------------------------------------------------------
*/

static int xml_append_byte_escaped(XStr *p, unsigned char c){
  switch( c ){
    case '&':
      return xstr_appendall(p, "&amp;");
    case '<':
      return xstr_appendall(p, "&lt;");
    case '>':
      return xstr_appendall(p, "&gt;");
    case '"':
      return xstr_appendall(p, "&quot;");
    case '\'':
      return xstr_appendall(p, "&apos;");
    default:
      if( c < 0x20 && c!=0x09 && c!=0x0a && c!=0x0d ){
        return 0;
      }
      if( c==0x7f ){
        return 0;
      }
      return xstr_append(p, (const char*)&c, 1);
  }
}

/*
** Append XML-escaped text, truncating to maxChars UTF-8 characters.
** Sets *pExceeded if truncation happened.
*/
static int xml_append_escaped_maxchars(
  XStr *p,
  const char *z,
  int n,
  int maxChars,
  int *pExceeded
){
  int i = 0;
  int nChar = 0;

  if( pExceeded ) *pExceeded = 0;
  if( z==0 ) return 0;
  if( n<0 ) n = (int)strlen(z);
  if( maxChars<0 ) maxChars = 0x7fffffff;

  while( i<n ){
    unsigned char c = (unsigned char)z[i];
    int isStart = ((c & 0xC0) != 0x80);
    int len;

    if( isStart ){
      if( nChar >= maxChars ){
        if( pExceeded ) *pExceeded = 1;
        break;
      }

      nChar++;
      len = utf8_byte_len(c);

      if( i + len > n ){
        len = n - i;
      }

      for(int j=0; j<len; j++){
        xml_append_byte_escaped(p, (unsigned char)z[i+j]);
      }

      i += len;
    }else{
      if( nChar >= maxChars ){
        if( pExceeded ) *pExceeded = 1;
        break;
      }

      nChar++;
      xml_append_byte_escaped(p, c);
      i++;
    }

    if( p->err ) return 1;
  }

  return 0;
}

/*
** Convert 1-based column number to Excel column letters.
*/
static void xlsx_column_ref(int iCol, char *zOut){
  char tmp[16];
  int n = 0;
  int i;

  if( iCol < 1 ) iCol = 1;

  while( iCol>0 && n<(int)sizeof(tmp)-1 ){
    int r = (iCol - 1) % 26;
    tmp[n++] = (char)('A' + r);
    iCol = (iCol - 1) / 26;
  }

  for(i=0; i<n; i++){
    zOut[i] = tmp[n - 1 - i];
  }

  zOut[n] = 0;
}

/*
** Append an inline-string cell.
*/
static int append_inline_cell(
  XStr *p,
  int iRow,
  int iCol,
  int bHeader,
  const char *zText,
  int nText,
  const char *zTable,
  const char *zCol
){
  char ref[16];
  int exceeded = 0;

  xlsx_column_ref(iCol, ref);

  if( xstr_appendf(p, "<c r=\"%s\"", ref) ) return 1;

  if( bHeader ){
    if( xstr_appendall(p, " s=\"1\"") ) return 1;
  }

  if( xstr_appendall(p, " t=\"inlineStr\"><is><t xml:space=\"preserve\">") ){
    return 1;
  }

  if( zText==0 ){
    zText = "";
    nText = 0;
  }

  if( nText<0 ) nText = (int)strlen(zText);

  if( xml_append_escaped_maxchars(p, zText, nText,
                                  XLSX_MAX_CELL_CHARS, &exceeded) ){
    return 1;
  }

  if( exceeded ){
    sqlite3_log(
      SQLITE_WARNING,
      "xlsx_export: cell value in table \"%s\" row %d column \"%s\" "
      "exceeds %d characters and was truncated",
      zTable ? zTable : "", iRow, zCol ? zCol : "",
      XLSX_MAX_CELL_CHARS
    );
  }

  if( xstr_appendall(p, "</t></is></c>") ) return 1;

  return 0;
}

static int append_null_cell(XStr *p, int iRow, int iCol){
  char ref[16];
  xlsx_column_ref(iCol, ref);
  return xstr_appendf(p, "<c r=\"%s\"/>", ref);
}

/*
** ------------------------------------------------------------------
** XLSX part builders
** ------------------------------------------------------------------
*/

static int build_content_types(XStr *p, int nSheets){
  int i;

  xstr_appendall(p,
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
  );

  for(i=1; i<=nSheets; i++){
    xstr_appendf(p,
      "<Override PartName=\"/xl/worksheets/sheet%d.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
      i
    );
  }

  xstr_appendall(p, "</Types>");

  return p->err ? SQLITE_NOMEM : SQLITE_OK;
}

static int build_root_rels(XStr *p){
  xstr_appendall(p,
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    "<Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
    "Target=\"xl/workbook.xml\"/>"
    "</Relationships>"
  );

  return p->err ? SQLITE_NOMEM : SQLITE_OK;
}

static int build_workbook(XStr *p, XlsxTableList *pList){
  int i;
  int dummy;

  xstr_appendall(p,
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
    "<sheets>"
  );

  for(i=0; i<pList->n; i++){
    xstr_appendall(p, "<sheet name=\"");

    if( xml_append_escaped_maxchars(p, pList->a[i].zSheet, -1,
                                    1000, &dummy) ){
      return SQLITE_NOMEM;
    }

    if( xstr_appendf(p, "\" sheetId=\"%d\" r:id=\"rId%d\"/>",
                     i + 1, i + 1) ){
      return SQLITE_NOMEM;
    }
  }

  xstr_appendall(p, "</sheets></workbook>");

  return p->err ? SQLITE_NOMEM : SQLITE_OK;
}

static int build_workbook_rels(XStr *p, int nSheets){
  int i;

  xstr_appendall(p,
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  );

  for(i=1; i<=nSheets; i++){
    xstr_appendf(p,
      "<Relationship Id=\"rId%d\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet%d.xml\"/>",
      i, i
    );
  }

  xstr_appendf(p,
    "<Relationship Id=\"rId%d\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
    "Target=\"styles.xml\"/>",
    nSheets + 1
  );

  xstr_appendall(p, "</Relationships>");

  return p->err ? SQLITE_NOMEM : SQLITE_OK;
}

static int build_styles(XStr *p){
  xstr_appendall(p,
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<fonts count=\"2\">"
    "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
    "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>"
    "</fonts>"
    "<fills count=\"2\">"
    "<fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"gray125\"/></fill>"
    "</fills>"
    "<borders count=\"1\">"
    "<border><left/><right/><top/><bottom/><diagonal/></border>"
    "</borders>"
    "<cellStyleXfs count=\"1\">"
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
    "</cellStyleXfs>"
    "<cellXfs count=\"2\">"
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
    "</cellXfs>"
    "<cellStyles count=\"1\">"
    "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
    "</cellStyles>"
    "</styleSheet>"
  );

  return p->err ? SQLITE_NOMEM : SQLITE_OK;
}

/*
** Build one worksheet XML part for a table.
*/
static int build_sheet_xml(
  sqlite3 *db,
  const char *zTable,
  char **pzOut,
  int *pnOut
){
  int rc;
  char *zIdent;
  char *zSql;
  sqlite3_stmt *pStmt = 0;
  XStr s;
  int nCol;
  int nColUse;
  int iRow = 1;
  int i;

  *pzOut = 0;
  *pnOut = 0;

  zIdent = quote_ident(zTable);
  if( zIdent==0 ) return SQLITE_NOMEM;

  zSql = sqlite3_mprintf("SELECT * FROM main.%s", zIdent);
  sqlite3_free(zIdent);

  if( zSql==0 ) return SQLITE_NOMEM;

  rc = sqlite3_prepare_v2(db, zSql, -1, &pStmt, 0);
  sqlite3_free(zSql);

  if( rc!=SQLITE_OK ){
    return rc;
  }

  xstr_init(&s);

  nCol = sqlite3_column_count(pStmt);
  nColUse = (nCol > XLSX_MAX_COLS ? XLSX_MAX_COLS : nCol);

  if( nCol > XLSX_MAX_COLS ){
    sqlite3_log(
      SQLITE_WARNING,
      "xlsx_export: table \"%s\" has more than %d columns; "
      "extra columns are omitted",
      zTable, XLSX_MAX_COLS
    );
  }

  xstr_appendall(&s,
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<sheetData>"
  );

  if( s.err ){
    rc = SQLITE_NOMEM;
    goto sheet_error;
  }

  /*
  ** Header row.
  */
  if( nColUse > 0 ){
    if( xstr_appendall(&s, "<row r=\"1\">") ){
      rc = SQLITE_NOMEM;
      goto sheet_error;
    }

    for(i=0; i<nColUse; i++){
      const char *zCol = sqlite3_column_name(pStmt, i);
      if( zCol==0 ) zCol = "";

      if( append_inline_cell(&s, 1, i+1, 1, zCol, -1, zTable, zCol) ){
        rc = SQLITE_NOMEM;
        goto sheet_error;
      }
    }

    if( xstr_appendall(&s, "</row>") ){
      rc = SQLITE_NOMEM;
      goto sheet_error;
    }
  }

  /*
  ** Data rows.
  */
  rc = sqlite3_step(pStmt);

  while( rc==SQLITE_ROW ){
    if( iRow >= XLSX_MAX_ROWS ){
      sqlite3_log(
        SQLITE_WARNING,
        "xlsx_export: table \"%s\" has more than %d rows; "
        "extra rows are omitted",
        zTable, XLSX_MAX_ROWS
      );
      rc = SQLITE_DONE;
      break;
    }

    iRow++;

    if( xstr_appendf(&s, "<row r=\"%d\">", iRow) ){
      rc = SQLITE_NOMEM;
      goto sheet_error;
    }

    for(i=0; i<nColUse; i++){
      int eType = sqlite3_column_type(pStmt, i);

      if( eType==SQLITE_NULL ){
        if( append_null_cell(&s, iRow, i+1) ){
          rc = SQLITE_NOMEM;
          goto sheet_error;
        }
        continue;
      }

      if( eType==SQLITE_BLOB ){
        int nBlob;
        char *zMsg;
        const char *zCol;

        nBlob = sqlite3_column_bytes(pStmt, i);
        zMsg = sqlite3_mprintf("(BLOB %d bytes)", nBlob);
        zCol = sqlite3_column_name(pStmt, i);

        if( zMsg==0 ){
          rc = SQLITE_NOMEM;
          goto sheet_error;
        }

        if( append_inline_cell(&s, iRow, i+1, 0, zMsg, -1, zTable, zCol) ){
          sqlite3_free(zMsg);
          rc = SQLITE_NOMEM;
          goto sheet_error;
        }

        sqlite3_free(zMsg);
      }else{
        const unsigned char *zVal = sqlite3_column_text(pStmt, i);
        int nVal = zVal ? sqlite3_column_bytes(pStmt, i) : 0;
        const char *zCol = sqlite3_column_name(pStmt, i);

        if( append_inline_cell(&s, iRow, i+1, 0,
                               (const char*)zVal, nVal,
                               zTable, zCol) ){
          rc = SQLITE_NOMEM;
          goto sheet_error;
        }
      }
    }

    if( xstr_appendall(&s, "</row>") ){
      rc = SQLITE_NOMEM;
      goto sheet_error;
    }

    rc = sqlite3_step(pStmt);
  }

  if( rc!=SQLITE_DONE ){
    goto sheet_error;
  }

  rc = SQLITE_OK;

  if( xstr_appendall(&s, "</sheetData>") ){
    rc = SQLITE_NOMEM;
    goto sheet_error;
  }

  if( nColUse > 0 ){
    char zLast[16];
    xlsx_column_ref(nColUse, zLast);

    if( xstr_appendf(&s, "<autoFilter ref=\"A1:%s%d\"/>", zLast, iRow) ){
      rc = SQLITE_NOMEM;
      goto sheet_error;
    }
  }

  if( xstr_appendall(&s, "</worksheet>") ){
    rc = SQLITE_NOMEM;
    goto sheet_error;
  }

  if( s.err ){
    rc = SQLITE_NOMEM;
    goto sheet_error;
  }

  sqlite3_finalize(pStmt);

  *pzOut = s.z;
  *pnOut = s.n;

  return SQLITE_OK;

sheet_error:
  if( pStmt ){
    sqlite3_finalize(pStmt);
  }
  xstr_free(&s);
  if( rc==SQLITE_OK ) rc = SQLITE_ERROR;
  return rc;
}

/*
** ------------------------------------------------------------------
** Table collection / validation
** ------------------------------------------------------------------
*/

static int table_exists(sqlite3 *db, const char *zTable){
  int res = 0;
  sqlite3_stmt *p = 0;
  int rc;

  rc = sqlite3_prepare_v2(
    db,
    "SELECT 1 FROM main.sqlite_master "
    "WHERE type='table' AND name=?1 COLLATE NOCASE",
    -1, &p, 0
  );

  if( rc==SQLITE_OK ){
    sqlite3_bind_text(p, 1, zTable, -1, SQLITE_TRANSIENT);
    rc = sqlite3_step(p);
    if( rc==SQLITE_ROW ){
      res = 1;
    }
    sqlite3_finalize(p);
  }

  return res;
}

static int can_select_table(sqlite3 *db, const char *zTable){
  char *zIdent;
  char *zSql;
  sqlite3_stmt *p = 0;
  int rc;

  zIdent = quote_ident(zTable);
  if( zIdent==0 ) return 0;

  zSql = sqlite3_mprintf("SELECT * FROM main.%s LIMIT 0", zIdent);
  sqlite3_free(zIdent);

  if( zSql==0 ) return 0;

  rc = sqlite3_prepare_v2(db, zSql, -1, &p, 0);
  sqlite3_free(zSql);

  if( rc==SQLITE_OK ){
    sqlite3_finalize(p);
    return 1;
  }

  return 0;
}

static int collect_all_tables(sqlite3 *db, XlsxTableList *pList){
  sqlite3_stmt *p = 0;
  int rc;

  rc = sqlite3_prepare_v2(
    db,
    "SELECT name FROM main.sqlite_master "
    "WHERE type='table' AND name NOT LIKE 'sqlite_%' "
    "ORDER BY rowid",
    -1, &p, 0
  );

  if( rc!=SQLITE_OK ) return rc;

  while( (rc = sqlite3_step(p))==SQLITE_ROW ){
    const char *z = (const char*)sqlite3_column_text(p, 0);
    if( z && list_add(pList, z) ){
      rc = SQLITE_NOMEM;
      break;
    }
  }

  sqlite3_finalize(p);

  if( rc==SQLITE_DONE ) rc = SQLITE_OK;
  return rc;
}

/*
** ------------------------------------------------------------------
** ZIP helpers
** ------------------------------------------------------------------
*/

static char *xlsx_random_hex(void){
  unsigned char rb[8];
  char *zHex;
  int i;

  zHex = sqlite3_malloc(17);
  if( zHex==0 ) return 0;

  sqlite3_randomness(sizeof(rb), rb);

  for(i=0; i<8; i++){
    sqlite3_snprintf(3, zHex + i*2, "%02x", rb[i]);
  }

  zHex[16] = 0;
  return zHex;
}

static int add_zip_part(
  sqlite3_stmt *pIns,
  const char *zName,
  const char *zData,
  int nData,
  sqlite3_int64 iMtime
){
  int rc;

  if( zData==0 ){
    zData = "";
    nData = 0;
  }

  if( nData<0 ) nData = (int)strlen(zData);

  sqlite3_reset(pIns);
  sqlite3_clear_bindings(pIns);

  rc = sqlite3_bind_text(pIns, 1, zName, -1, SQLITE_TRANSIENT);
  if( rc==SQLITE_OK ){
    rc = sqlite3_bind_int64(pIns, 2, iMtime);
  }
  if( rc==SQLITE_OK ){
    rc = sqlite3_bind_blob(pIns, 3, zData, nData, SQLITE_TRANSIENT);
  }

  if( rc!=SQLITE_OK ) return rc;

  rc = sqlite3_step(pIns);

  if( rc==SQLITE_DONE ){
    rc = SQLITE_OK;
  }

  return rc;
}

/*
** ------------------------------------------------------------------
** SQL function implementations
** ------------------------------------------------------------------
*/

static void xlsx_export_version_func(
  sqlite3_context *ctx,
  int argc,
  sqlite3_value **argv
){
  (void)argc;
  (void)argv;

  sqlite3_result_text(
    ctx,
    "2026-11-03 Qwen3.8-Max (Thinking)",
    -1,
    SQLITE_STATIC
  );
}

static void xlsx_export_func(
  sqlite3_context *ctx,
  int argc,
  sqlite3_value **argv
){
  sqlite3 *db;
  XlsxTableList list;
  XStr part;
  const char *zFile;
  char *zRand = 0;
  char *zZipTab = 0;
  char *zTmpFile = 0;
  char *zSql = 0;
  char *zErr = 0;
  sqlite3_stmt *pIns = 0;

  int rc = SQLITE_OK;
  int bZipCreated = 0;
  int nExported = 0;
  int i;
  int nKeep;
  sqlite3_int64 iMtime;

  if( argc < 1 ){
    sqlite3_result_error(
      ctx,
      "usage: xlsx_export(filename [, table1, table2, ...])",
      -1
    );
    return;
  }

  zFile = (const char*)sqlite3_value_text(argv[0]);

  if( zFile==0 || zFile[0]==0 ){
    sqlite3_result_error(
      ctx,
      "xlsx_export: missing or empty filename",
      -1
    );
    return;
  }

  db = sqlite3_context_db_handle(ctx);

  list_init(&list);
  xstr_init(&part);

  iMtime = (sqlite3_int64)time(0);

  /*
  ** Collect table names.
  */
  if( argc==1 ){
    rc = collect_all_tables(db, &list);
    if( rc!=SQLITE_OK ){
      zErr = sqlite3_mprintf(
        "xlsx_export: cannot list tables: %s",
        sqlite3_errmsg(db)
      );
      goto cleanup;
    }
  }else{
    for(i=1; i<argc; i++){
      const char *zTable = (const char*)sqlite3_value_text(argv[i]);

      if( zTable && zTable[0] ){
        if( table_exists(db, zTable) ){
          if( list_add(&list, zTable) ){
            rc = SQLITE_NOMEM;
            zErr = sqlite3_mprintf("xlsx_export: out of memory");
            goto cleanup;
          }
        }else{
          sqlite3_log(
            SQLITE_WARNING,
            "xlsx_export: table \"%s\" not found in main schema; skipped",
            zTable
          );
        }
      }
    }
  }

  /*
  ** Keep only tables that can be read.
  */
  nKeep = 0;

  for(i=0; i<list.n; i++){
    if( can_select_table(db, list.a[i].zTable) ){
      if( i != nKeep ){
        list.a[nKeep].zTable = list.a[i].zTable;
        list.a[nKeep].zSheet = 0;
      }
      nKeep++;
    }else{
      sqlite3_log(
        SQLITE_WARNING,
        "xlsx_export: cannot read table \"%s\"; skipped",
        list.a[i].zTable
      );
      sqlite3_free(list.a[i].zTable);
    }
  }

  list.n = nKeep;

  if( list.n==0 ){
    rc = SQLITE_OK;
    goto cleanup;
  }

  /*
  ** Assign sanitized unique sheet names.
  */
  for(i=0; i<list.n; i++){
    list.a[i].zSheet = make_unique_sheet(list.a[i].zTable, list.a, i);
    if( list.a[i].zSheet==0 ){
      rc = SQLITE_NOMEM;
      zErr = sqlite3_mprintf("xlsx_export: out of memory");
      goto cleanup;
    }
  }

  /*
  ** Create temporary ZIP output file name and zipfile virtual table.
  */
  zRand = xlsx_random_hex();
  if( zRand==0 ){
    rc = SQLITE_NOMEM;
    zErr = sqlite3_mprintf("xlsx_export: out of memory");
    goto cleanup;
  }

  zZipTab = sqlite3_mprintf("xlsxzip_%s", zRand);
  zTmpFile = sqlite3_mprintf("%s.tmp-%s", zFile, zRand);
  sqlite3_free(zRand);
  zRand = 0;

  if( zZipTab==0 || zTmpFile==0 ){
    rc = SQLITE_NOMEM;
    zErr = sqlite3_mprintf("xlsx_export: out of memory");
    goto cleanup;
  }

  zSql = sqlite3_mprintf("DROP TABLE IF EXISTS temp.%s", zZipTab);
  if( zSql ){
    sqlite3_exec(db, zSql, 0, 0, 0);
    sqlite3_free(zSql);
    zSql = 0;
  }

  zSql = sqlite3_mprintf(
    "CREATE VIRTUAL TABLE temp.%s USING zipfile(%Q)",
    zZipTab, zTmpFile
  );

  if( zSql==0 ){
    rc = SQLITE_NOMEM;
    zErr = sqlite3_mprintf("xlsx_export: out of memory");
    goto cleanup;
  }

  rc = sqlite3_exec(db, zSql, 0, 0, 0);
  sqlite3_free(zSql);
  zSql = 0;

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: cannot create ZIP archive \"%s\" "
      "(ensure the zipfile extension is loaded): %s",
      zTmpFile, sqlite3_errmsg(db)
    );
    rc = SQLITE_ERROR;
    goto cleanup;
  }

  bZipCreated = 1;

  zSql = sqlite3_mprintf(
    "INSERT INTO temp.%s (name, mode, mtime, data) "
    "VALUES (?, %d, ?, ?)",
    zZipTab, XLSX_FILE_MODE
  );

  if( zSql==0 ){
    rc = SQLITE_NOMEM;
    zErr = sqlite3_mprintf("xlsx_export: out of memory");
    goto cleanup;
  }

  rc = sqlite3_prepare_v2(db, zSql, -1, &pIns, 0);
  sqlite3_free(zSql);
  zSql = 0;

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: cannot insert into ZIP archive: %s",
      sqlite3_errmsg(db)
    );
    goto cleanup;
  }

  /*
  ** [Content_Types].xml
  */
  xstr_init(&part);
  rc = build_content_types(&part, list.n);
  if( rc==SQLITE_OK ){
    rc = add_zip_part(pIns, "[Content_Types].xml", part.z, part.n, iMtime);
  }
  xstr_free(&part);

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: failed to write [Content_Types].xml: %s",
      sqlite3_errmsg(db)
    );
    goto cleanup;
  }

  /*
  ** _rels/.rels
  */
  xstr_init(&part);
  rc = build_root_rels(&part);
  if( rc==SQLITE_OK ){
    rc = add_zip_part(pIns, "_rels/.rels", part.z, part.n, iMtime);
  }
  xstr_free(&part);

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: failed to write _rels/.rels: %s",
      sqlite3_errmsg(db)
    );
    goto cleanup;
  }

  /*
  ** xl/workbook.xml
  */
  xstr_init(&part);
  rc = build_workbook(&part, &list);
  if( rc==SQLITE_OK ){
    rc = add_zip_part(pIns, "xl/workbook.xml", part.z, part.n, iMtime);
  }
  xstr_free(&part);

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: failed to write xl/workbook.xml: %s",
      sqlite3_errmsg(db)
    );
    goto cleanup;
  }

  /*
  ** xl/_rels/workbook.xml.rels
  */
  xstr_init(&part);
  rc = build_workbook_rels(&part, list.n);
  if( rc==SQLITE_OK ){
    rc = add_zip_part(pIns, "xl/_rels/workbook.xml.rels",
                      part.z, part.n, iMtime);
  }
  xstr_free(&part);

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: failed to write xl/_rels/workbook.xml.rels: %s",
      sqlite3_errmsg(db)
    );
    goto cleanup;
  }

  /*
  ** xl/styles.xml
  */
  xstr_init(&part);
  rc = build_styles(&part);
  if( rc==SQLITE_OK ){
    rc = add_zip_part(pIns, "xl/styles.xml", part.z, part.n, iMtime);
  }
  xstr_free(&part);

  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf(
      "xlsx_export: failed to write xl/styles.xml: %s",
      sqlite3_errmsg(db)
    );
    goto cleanup;
  }

  /*
  ** Worksheets.
  */
  for(i=0; i<list.n; i++){
    char *zPart = sqlite3_mprintf("xl/worksheets/sheet%d.xml", i + 1);
    char *zData = 0;
    int nData = 0;

    if( zPart==0 ){
      rc = SQLITE_NOMEM;
      zErr = sqlite3_mprintf("xlsx_export: out of memory");
      break;
    }

    rc = build_sheet_xml(db, list.a[i].zTable, &zData, &nData);

    if( rc==SQLITE_OK ){
      rc = add_zip_part(pIns, zPart, zData, nData, iMtime);
      if( rc==SQLITE_OK ){
        nExported++;
      }
    }

    if( rc!=SQLITE_OK ){
      if( zErr==0 ){
        zErr = sqlite3_mprintf(
          "xlsx_export: failed to export table \"%s\": %s",
          list.a[i].zTable, sqlite3_errmsg(db)
        );
      }
      sqlite3_free(zData);
      sqlite3_free(zPart);
      break;
    }

    sqlite3_free(zData);
    sqlite3_free(zPart);
  }

cleanup:
  if( pIns ){
    sqlite3_finalize(pIns);
    pIns = 0;
  }

  if( bZipCreated && zZipTab ){
    zSql = sqlite3_mprintf("DROP TABLE IF EXISTS temp.%s", zZipTab);
    if( zSql ){
      int rcDrop = sqlite3_exec(db, zSql, 0, 0, 0);
      sqlite3_free(zSql);
      zSql = 0;

      if( rcDrop!=SQLITE_OK && rc==SQLITE_OK ){
        rc = rcDrop;
        zErr = sqlite3_mprintf(
          "xlsx_export: cannot finalize ZIP archive: %s",
          sqlite3_errmsg(db)
        );
      }
    }
  }

  if( rc==SQLITE_OK ){
    if( zTmpFile ){
      remove(zFile);
      if( rename(zTmpFile, zFile)!=0 ){
        rc = SQLITE_ERROR;
        zErr = sqlite3_mprintf(
          "xlsx_export: cannot move temporary file to \"%s\"",
          zFile
        );
        remove(zTmpFile);
      }
    }
  }

  if( rc!=SQLITE_OK ){
    if( zTmpFile ){
      remove(zTmpFile);
    }

    if( zErr ){
      sqlite3_result_error(ctx, zErr, -1);
    }else{
      sqlite3_result_error_code(ctx, rc);
    }
  }else{
    sqlite3_result_int(ctx, nExported);
  }

  sqlite3_free(zErr);
  sqlite3_free(zRand);
  sqlite3_free(zZipTab);
  sqlite3_free(zTmpFile);

  list_free(&list);
  xstr_free(&part);
}

/*
** ------------------------------------------------------------------
** Extension entry point
** ------------------------------------------------------------------
*/

#ifdef _WIN32
#define XLSX_EXPORT __declspec(dllexport)
#else
#define XLSX_EXPORT
#endif

XLSX_EXPORT int sqlite3_xlsxexport_init(
  sqlite3 *db,
  char **pzErrMsg,
  const sqlite3_api_routines *pApi
){
  int rc;

  SQLITE_EXTENSION_INIT2(pApi);
  (void)pzErrMsg;

  rc = sqlite3_create_function(
    db,
    "xlsx_export",
    -1,
    SQLITE_UTF8 | SQLITE_DIRECTONLY,
    0,
    xlsx_export_func,
    0,
    0
  );

  if( rc==SQLITE_OK ){
    rc = sqlite3_create_function(
      db,
      "xlsx_export_version",
      0,
      SQLITE_UTF8 | SQLITE_DETERMINISTIC,
      0,
      xlsx_export_version_func,
      0,
      0
    );
  }

  return rc;
}