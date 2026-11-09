/*
** 2026-11-03
**
** xlsxexport.c -- SQLite loadable extension: export tables as an XLSX
** (Excel 2007+) workbook.
**
** SQL FUNCTIONS
**
**   xlsx_export(FILENAME)                        -- export every table
**   xlsx_export(FILENAME, TABLE [, TABLE ...])   -- export named tables
**   xlsx_export_version()                        -- build identification
**
** xlsx_export() writes the XLSX file FILENAME and returns the number of
** tables that were exported.  With a single argument every table of the
** "main" schema is exported (internal sqlite_* tables are excluded;
** views are not included -- name them explicitly if you want them).
** Each exported table becomes one worksheet:
**
**   * the sheet name is the (sanitized) table name,
**   * the first row holds the column names, in bold,
**   * the sheet carries an auto-filter over that header row,
**   * INTEGER and REAL values are written as numbers, TEXT values as
**     (inline) strings, NULL cells are left empty, and BLOB values are
**     written as X'HEX' text.
**
** An .xlsx file is a ZIP archive of XML parts.  The archive is built with
** the zipfile() aggregate function of the SQLite "zipfile" extension
** (https://sqlite.org/zipfile.html), which must therefore be loaded into
** the same database connection before xlsx_export() is used:
**
**   .load ./zipfile
**   .load ./xlsxexport
**   SELECT xlsx_export('dump.xlsx');
**   SELECT xlsx_export('report.xlsx', 'sales', 'customers');
**
** No external XML library is used: every XML part is written directly by
** this file.
**
** BUILDING
**
**   Unix:    gcc -O2 -fPIC -shared xlsxexport.c -o xlsxexport.so
**   Windows: cl /O2 xlsxexport.c -link -dll -out:xlsxexport.dll
**
** EXCEL LIMITS ENFORCED (with a warning on stderr when data is affected)
**
**   * cell text is limited to 32,767 characters -- longer values are
**     truncated,
**   * at most 1,048,576 rows (including the header row) and 16,384
**     columns per worksheet -- excess rows/columns are dropped,
**   * integers that a 64-bit double cannot represent exactly (beyond
**     +/-2^53) are written as text so that no digits are lost,
**   * non-finite REAL values (inf, nan) are written as text,
**   * a BLOB larger than the cell limit is truncated in its hex form.
**
** SHEET NAME SANITIZATION
**
**   * the forbidden characters : \ / ? * [ ] and control characters are
**     replaced by "_",
**   * the name is truncated to 31 characters (never splitting a Unicode
**     character),
**   * a name may not begin or end with an apostrophe,
**   * the reserved name "History" is avoided,
**   * duplicates (case-insensitive) get "_2", "_3", ... suffixes.
**
** NOTES
**
**   * In full-schema mode a table that cannot be read (e.g. a virtual
**     table whose module is missing) is skipped with a warning; when
**     tables are named explicitly such a failure raises an error.
**   * xlsx_export() writes a file and is therefore registered
**     SQLITE_DIRECTONLY (like the fileio extension's writefile()).
**     Compile with -DXLSXEXPORT_ALLOW_UNSAFE to permit its use from
**     triggers and views as well.
**   * Requires SQLite 3.8.3 or newer (zipfile extension + CTEs).
**
** xlsx_export_version() returns "2026-11-03 GLM-5.3-Flash (Max)".
*/
#include "sqlite3ext.h"            /* Includes sqlite3.h */
SQLITE_EXTENSION_INIT1

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>
#include <errno.h>
#include <float.h>

/* Excel hard limits */
#define XLSX_MAX_CHARS    32767       /* max characters per cell          */
#define XLSX_MAX_NAME       31       /* max characters per sheet name    */
#define XLSX_MAX_ROWS   1048576       /* max rows per worksheet (2^20)    */
#define XLSX_MAX_COLS     16384       /* max cols per worksheet (2^14)    */
#define XLSX_MAX_EXACT_INT 9007199254740992.0  /* 2^53                  */

/* Buffer large enough for any sanitized sheet name (31 UTF-16 units can
** need up to 4 bytes each in UTF-8, plus slack). */
#define XLSX_NAME_BUF   (4*XLSX_MAX_NAME+8)

#ifndef SQLITE_DIRECTONLY
# define SQLITE_DIRECTONLY 0
#endif
#ifndef SQLITE_WARNING
# define SQLITE_WARNING 28
#endif

/* xlsx_export() has side effects (it writes a file).  Like writefile()
** of the fileio extension it is kept out of triggers/views/checked
** schema contexts unless the builder explicitly allows it. */
#ifdef XLSXEXPORT_ALLOW_UNSAFE
# define XLSXEXPORT_FUNCFLAGS SQLITE_UTF8
#else
# define XLSXEXPORT_FUNCFLAGS (SQLITE_UTF8|SQLITE_DIRECTONLY)
#endif

#define XLSX_XML_DECL \
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"

/* Name of the temp table used when there are too many zip entries for
** the connection's bound-parameter limit. */
#define XLSX_ZIPTEMP "xlsxexport_zipentry"

/*
** Emit a warning on stderr and to the sqlite error log.
*/
static void xlsxWarn(const char *zFmt, ...){
  char zMsg[512];
  va_list ap;
  va_start(ap, zFmt);
  sqlite3_vsnprintf((int)sizeof(zMsg), zMsg, zFmt, ap);
  va_end(ap);
  fprintf(stderr, "xlsx_export: warning: %s\n", zMsg);
  fflush(stderr);
  sqlite3_log(SQLITE_WARNING, "xlsx_export: %s", zMsg);
}

/*************************************************************************
** Growable string buffer used to build every XML part in memory.
*/
typedef struct StrBuf {
  char *z;            /* NUL terminated buffer (NUL not counted in n) */
  size_t n;           /* bytes of payload */
  size_t nAlloc;      /* bytes allocated at z */
  int oom;            /* set when an allocation failed */
} StrBuf;

static void sbInit(StrBuf *p){
  p->n = 0;
  p->nAlloc = 256;
  p->z = (char*)sqlite3_malloc64(p->nAlloc);
  p->oom = (p->z==0);
  if( !p->oom ) p->z[0] = 0;
}
static void sbFree(StrBuf *p){
  sqlite3_free(p->z);
  memset(p, 0, sizeof(*p));
}
static int sbMakeRoom(StrBuf *p, size_t nAdd){
  size_t nNeed, nNew;
  char *zNew;
  if( p->oom ) return 0;
  nNeed = p->n + nAdd + 1;
  if( nNeed <= p->nAlloc ) return 1;
  nNew = p->nAlloc ? p->nAlloc : 256;
  while( nNew < nNeed ){
    size_t nDbl = nNew*2;
    if( nDbl<=nNew ){ nNew = nNeed; break; }   /* size_t overflow */
    nNew = nDbl;
  }
  zNew = (char*)sqlite3_realloc64(p->z, nNew);
  if( zNew==0 ){
    p->oom = 1;
    return 0;
  }
  p->z = zNew;
  p->nAlloc = nNew;
  return 1;
}
static void sbAppendN(StrBuf *p, const void *zV, size_t k){
  if( k==0 ) return;
  if( !sbMakeRoom(p, k) ) return;
  memcpy(p->z + p->n, zV, k);
  p->n += k;
  p->z[p->n] = 0;
}
static void sbAppend(StrBuf *p, const char *z){
  sbAppendN(p, z, strlen(z));
}
/*
** Append a formatted fragment.  All call sites produce short,
** bounded output (numeric fragments and fixed boilerplate); string
** payloads are appended with sbAppend()/sbAppendN().  The fallback
** loop is only a safety net for future call sites.
*/
static void sbAppendF(StrBuf *p, const char *zFmt, ...){
  char zStk[256];
  va_list ap;
  size_t n, sz;
  va_start(ap, zFmt);
  sqlite3_vsnprintf((int)sizeof(zStk), zStk, zFmt, ap);
  va_end(ap);
  n = strlen(zStk);
  if( n < sizeof(zStk)-1 ){
    sbAppendN(p, zStk, n);
    return;
  }
  for(sz=512; ; sz*=2){
    char *zBig = (char*)sqlite3_malloc64((sqlite3_uint64)sz);
    if( zBig==0 ){
      p->oom = 1;
      return;
    }
    va_start(ap, zFmt);
    sqlite3_vsnprintf((int)sz, zBig, zFmt, ap);
    va_end(ap);
    n = strlen(zBig);
    if( n<sz-1 || sz>=(size_t)1<<20 ){
      sbAppendN(p, zBig, n);
      sqlite3_free(zBig);
      return;
    }
    sqlite3_free(zBig);
  }
}

/*************************************************************************
** UTF-8 helpers.  SQLite text may contain invalid UTF-8; every decoder
** here is strict so that the generated XML stays well formed.
*/

/*
** Decode one code point.  Returns the number of bytes consumed (>=1).
** On malformed input *pCp is set to 0xFFFFFFFF and 1 byte is consumed.
*/
static int xlsxUtf8Decode(const unsigned char *z, size_t n, unsigned int *pCp){
  unsigned int c = z[0];
  *pCp = 0xFFFFFFFF;
  if( c<0x80 ){
    *pCp = c;
    return 1;
  }
  if( c<0xC0 ) return 1;                        /* stray continuation */
  if( c<0xE0 ){
    if( n>=2 && (z[1]&0xC0)==0x80 ){
      unsigned int v = ((c&0x1F)<<6) | (z[1]&0x3F);
      if( v>=0x80 ){ *pCp = v; return 2; }
    }
    return 1;
  }
  if( c<0xF0 ){
    if( n>=3 && (z[1]&0xC0)==0x80 && (z[2]&0xC0)==0x80 ){
      unsigned int v = ((c&0x0F)<<12) | ((z[1]&0x3F)<<6) | (z[2]&0x3F);
      if( v>=0x800 && (v<0xD800 || v>0xDFFF) ){ *pCp = v; return 3; }
    }
    return 1;
  }
  if( c<0xF8 ){
    if( n>=4 && (z[1]&0xC0)==0x80 && (z[2]&0xC0)==0x80 && (z[3]&0xC0)==0x80 ){
      unsigned int v = ((c&0x07)<<18) | ((z[1]&0x3F)<<12)
                      | ((z[2]&0x3F)<<6) | (z[3]&0x3F);
      if( v>=0x10000 && v<=0x10FFFF ){ *pCp = v; return 4; }
    }
  }
  return 1;
}

/* Encode one code point as UTF-8.  z must have room for 4 bytes. */
static int xlsxEncodeUtf8(char *z, unsigned int cp){
  if( cp<0x80 ){
    z[0] = (char)cp;
    return 1;
  }
  if( cp<0x800 ){
    z[0] = (char)(0xC0 | (cp>>6));
    z[1] = (char)(0x80 | (cp&0x3F));
    return 2;
  }
  if( cp<0x10000 ){
    z[0] = (char)(0xE0 | (cp>>12));
    z[1] = (char)(0x80 | ((cp>>6)&0x3F));
    z[2] = (char)(0x80 | (cp&0x3F));
    return 3;
  }
  z[0] = (char)(0xF0 | (cp>>18));
  z[1] = (char)(0x80 | ((cp>>12)&0x3F));
  z[2] = (char)(0x80 | ((cp>>6)&0x3F));
  z[3] = (char)(0x80 | (cp&0x3F));
  return 4;
}

/* Number of UTF-8 bytes in z[0..n) that fit in maxUnits UTF-16 units. */
static size_t xlsxTruncateUnits(const char *z, size_t n, int maxUnits){
  size_t i = 0;
  unsigned int u = 0;
  while( i<n ){
    unsigned int cp;
    unsigned int du;
    int k = xlsxUtf8Decode((const unsigned char*)z+i, n-i, &cp);
    du = (cp==0xFFFFFFFF || cp<=0xFFFF) ? 1 : 2;
    if( u+du > (unsigned)maxUnits ) break;
    u += du;
    i += (size_t)k;
  }
  return i;
}

static int xlsxIsWs(char c){ return c==' ' || c=='\t' || c=='\n' || c=='\r'; }
static int xlsxIsHex(char c){
  return (c>='0'&&c<='9') || (c>='a'&&c<='f') || (c>='A'&&c<='F');
}

/*************************************************************************
** Sheet name sanitation.
**
** Excel restrictions: max 31 characters, none of : \ / ? * [ ], not
** empty, no leading/trailing apostrophe, unique in the workbook
** (case-insensitive), and "History" is reserved.
*/

/*
** Sanitize table name zSrc into zOut (XLSX_NAME_BUF bytes).
*/
static void xlsxSanitizeSheetName(const char *zSrc, char *zOut){
  size_t i = 0, n, o = 0;
  unsigned int nUnits = 0;
  n = strlen(zSrc);
  while( i<n ){
    unsigned int cp;
    unsigned int du;
    int k = xlsxUtf8Decode((const unsigned char*)zSrc+i, n-i, &cp);
    if( cp==0xFFFFFFFF
     || cp==':' || cp=='\\' || cp=='/' || cp=='?' || cp=='*'
     || cp=='[' || cp==']' || cp<0x20 ){
      cp = '_';
    }
    du = (cp>0xFFFF) ? 2 : 1;
    if( nUnits+du > (unsigned)XLSX_MAX_NAME ) break;
    o += (size_t)xlsxEncodeUtf8(zOut+o, cp);
    nUnits += du;
    i += (size_t)k;
  }
  zOut[o] = 0;
  if( o==0 ){
    memcpy(zOut, "Sheet", 6);
    o = 5;
  }
  if( zOut[0]=='\'' ) zOut[0] = '_';
  if( zOut[o-1]=='\'' ) zOut[o-1] = '_';
  if( sqlite3_stricmp(zOut, "History")==0 ){
    zOut[o++] = '_';                 /* "History" is reserved by Excel */
    zOut[o] = 0;
  }
}

/*
** Make zName different (case-insensitively) from every name in
** azUsed[0..nUsed-1], appending "_2", "_3", ... while staying within
** 31 characters.
*/
static void xlsxUniquifySheetName(char *zName, char **azUsed, int nUsed){
  int i, k;
  char zBase[XLSX_NAME_BUF], zTry[XLSX_NAME_BUF];
  for(i=0; i<nUsed; i++){
    if( sqlite3_stricmp(azUsed[i], zName)==0 ) break;
  }
  if( i>=nUsed ) return;
  memcpy(zBase, zName, strlen(zName)+1);
  for(k=2; k<100000; k++){
    char zSuf[16];
    size_t nSuf, nBase;
    sqlite3_snprintf((int)sizeof(zSuf), zSuf, "_%d", k);
    nSuf = strlen(zSuf);
    nBase = xlsxTruncateUnits(zBase, strlen(zBase),
                              XLSX_MAX_NAME-(int)nSuf);
    memcpy(zTry, zBase, nBase);
    zTry[nBase] = 0;
    memcpy(zTry+nBase, zSuf, nSuf+1);
    for(i=0; i<nUsed; i++){
      if( sqlite3_stricmp(azUsed[i], zTry)==0 ) break;
    }
    if( i>=nUsed ){
      memcpy(zName, zTry, strlen(zTry)+1);
      return;
    }
  }
}

/*************************************************************************
** Worksheet XML generation.
*/

/* Append z[0..n) as escaped inline-string text.  Applies:
**   - XML entity escaping (& < > " '),
**   - Excel's _xHHHH_ escaping for characters that are illegal in XML,
**   - protection of literal "_xHHHH_" sequences (via _x005F_),
**   - replacement of invalid UTF-8 with U+FFFD,
**   - truncation at Excel's 32,767-character cell limit (+ warning).
*/
static void xlsxAppendEscapedText(StrBuf *p, const char *z, size_t n,
                                  const char *zTable, const char *zCol,
                                  sqlite3_int64 iRow){
  size_t i = 0;
  unsigned int nUnits = 0;
  int bTrunc = 0;
  while( i<n ){
    unsigned int cp, du;
    int k = xlsxUtf8Decode((const unsigned char*)z+i, n-i, &cp);
    if( cp==0xFFFFFFFF ) cp = 0xFFFD;
    du = (cp>0xFFFF) ? 2 : 1;
    if( nUnits+du > (unsigned)XLSX_MAX_CHARS ){ bTrunc = 1; break; }
    switch( cp ){
      case '&':  sbAppend(p, "&amp;");   break;
      case '<':  sbAppend(p, "&lt;");    break;
      case '>':  sbAppend(p, "&gt;");    break;
      case '"':  sbAppend(p, "&quot;");  break;
      case '\'': sbAppend(p, "&apos;");  break;
      default:
        if( cp=='_'
         && i+6<n && z[i+1]=='x'
         && xlsxIsHex(z[i+2]) && xlsxIsHex(z[i+3])
         && xlsxIsHex(z[i+4]) && xlsxIsHex(z[i+5])
         && z[i+6]=='_' ){
          sbAppend(p, "_x005F_");       /* protect literal _xHHHH_ */
        }else if( (cp<0x20 && cp!=0x09 && cp!=0x0A && cp!=0x0D)
               || cp==0xFFFE || cp==0xFFFF ){
          sbAppendF(p, "_x%04X_", cp);
        }else{
          char a[4];
          sbAppendN(p, a, (size_t)xlsxEncodeUtf8(a, cp));
        }
        break;
    }
    nUnits += du;
    i += (size_t)k;
  }
  if( bTrunc ){
    xlsxWarn("text in row %lld, column \"%s\" of table \"%s\" exceeds "
             "Excel's %d-character cell limit; content truncated",
             iRow, zCol, zTable, XLSX_MAX_CHARS);
  }
}

/* Append a BLOB as X'HEX' text (truncated to the cell limit). */
static void xlsxAppendBlobHex(StrBuf *p, const void *pV, size_t n,
                              const char *zTable, const char *zCol,
                              sqlite3_int64 iRow){
  static const char hexdig[] = "0123456789ABCDEF";
  const unsigned char *z = (const unsigned char*)pV;
  size_t mx = ((size_t)XLSX_MAX_CHARS-3)/2;
  size_t i;
  if( z==0 ) n = 0;
  if( n>mx ){
    xlsxWarn("blob in row %lld, column \"%s\" of table \"%s\" exceeds "
             "Excel's %d-character cell limit; hex dump truncated",
             iRow, zCol, zTable, XLSX_MAX_CHARS);
    n = mx;
  }
  sbAppend(p, "X'");
  for(i=0; i<n; i++){
    char a[2];
    a[0] = hexdig[z[i]>>4];
    a[1] = hexdig[z[i]&15];
    sbAppendN(p, a, 2);
  }
  sbAppend(p, "'");
}

/*
** Write one inline-string cell.  zRef is the A1-style cell reference,
** bHeader selects the bold style (xf 1).
*/
static void xlsxWriteStrCell(StrBuf *p, const char *zRef,
                             const char *z, size_t n, int bHeader,
                             const char *zTable, const char *zCol,
                             sqlite3_int64 iRow){
  int bPreserve;
  bPreserve = (n>0 && (xlsxIsWs(z[0]) || xlsxIsWs(z[n-1])))
           || (n>(size_t)XLSX_MAX_CHARS);
  sbAppendF(p, "<c r=\"%s\" t=\"inlineStr\"%s><is><t%s>",
            zRef, bHeader ? " s=\"1\"" : "",
            bPreserve ? " xml:space=\"preserve\"" : "");
  xlsxAppendEscapedText(p, z, n, zTable, zCol, iRow);
  sbAppend(p, "</t></is></c>");
}

/* 0-based column index -> letters ("A", "Z", "AA", "XFD"). zOut>=4 bytes */
static void xlsxColName(int iCol, char *zOut){
  char a[8];
  int n = 0, c = iCol, i;
  do{
    a[n++] = (char)('A' + (c%26));
    c = c/26 - 1;
  }while( c>=0 );
  for(i=0; i<n; i++) zOut[i] = a[n-1-i];
  zOut[n] = 0;
}

/*
** One zip entry of the future workbook.
*/
typedef struct XlsxEntry {
  char *zName;          /* zip member name, e.g. "xl/worksheets/sheet1.xml" */
  StrBuf data;          /* member payload (NUL-free UTF-8 XML) */
} XlsxEntry;

/*
** Generate the worksheet XML for table zTable into pEnt->data.
** Returns SQLITE_OK or an error message in *pzErr (sqlite3_mprintf'ed).
*/
static int xlsxGenSheet(sqlite3 *db, XlsxEntry *pEnt, const char *zTable,
                        char **pzErr){
  StrBuf sql;
  StrBuf *out = &pEnt->data;
  sqlite3_stmt *pStmt = 0;
  int rc = SQLITE_OK;
  int nCol, nColOut, j;
  sqlite3_int64 nDataRows = 0;
  size_t dimOff = 0;
  char zLastCol[8];
  char zRef[24];

  *pzErr = 0;
  sbInit(&sql);

  /* Build:  SELECT * FROM "<table>"   (quotes doubled) */
  sbAppend(&sql, "SELECT * FROM \"");
  {
    const char *p;
    for(p=zTable; *p; p++){
      if( *p=='"' ) sbAppend(&sql, "\"\"");
      else{
        char c = *p;
        sbAppendN(&sql, &c, 1);
      }
    }
  }
  sbAppend(&sql, "\"");
  if( sql.oom ) goto oom_exit;
  rc = sqlite3_prepare_v2(db, sql.z, -1, &pStmt, 0);
  sbFree(&sql);
  if( rc!=SQLITE_OK ){
    *pzErr = sqlite3_mprintf("table \"%s\": %s", zTable, sqlite3_errmsg(db));
    goto exit_fail;
  }

  nCol = sqlite3_column_count(pStmt);
  if( nCol<=0 ){
    *pzErr = sqlite3_mprintf("table \"%s\" has no columns", zTable);
    goto exit_fail;
  }
  nColOut = nCol;
  if( nColOut>XLSX_MAX_COLS ){
    xlsxWarn("table \"%s\" has %d columns; only the first %d are exported",
             zTable, nCol, XLSX_MAX_COLS);
    nColOut = XLSX_MAX_COLS;
  }
  xlsxColName(nColOut-1, zLastCol);

  /* worksheet prologue.  The dimension ref is a 13-byte placeholder that
  ** is patched with the true value once all rows are known. */
  sbAppend(out, XLSX_XML_DECL
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<dimension ref=\"");
  dimOff = out->n;
  sbAppend(out, "A1:ZZZZZZZZZZ");      /* 13 bytes, patched below */
  sbAppend(out,
    "\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
    "<sheetFormatPr defaultRowHeight=\"15\"/><sheetData>");

  /* ---- header row (bold, style xf 1) ---- */
  sbAppendF(out, "<row r=\"1\" spans=\"1:%d\">", nColOut);
  for(j=0; j<nColOut; j++){
    const char *zColName = sqlite3_column_name(pStmt, j);
    char zC[8];
    if( zColName==0 ) zColName = "";
    xlsxColName(j, zC);
    sqlite3_snprintf((int)sizeof(zRef), zRef, "%s1", zC);
    xlsxWriteStrCell(out, zRef, zColName, strlen(zColName), 1,
                     zTable, zColName, 1);
  }
  sbAppend(out, "</row>");

  /* ---- data rows ---- */
  while( (rc = sqlite3_step(pStmt))==SQLITE_ROW ){
    sqlite3_int64 r;
    if( nDataRows >= (sqlite3_int64)XLSX_MAX_ROWS-1 ){
      xlsxWarn("table \"%s\" has more than %d data rows; the remaining "
               "rows are not exported", zTable, XLSX_MAX_ROWS-1);
      rc = SQLITE_DONE;                 /* intentional stop, not an error */
      break;
    }
    nDataRows++;
    r = nDataRows + 1;                     /* header occupies row 1 */
    sbAppendF(out, "<row r=\"%lld\" spans=\"1:%d\">", r, nColOut);
    for(j=0; j<nColOut; j++){
      int eType = sqlite3_column_type(pStmt, j);
      const char *zCol = sqlite3_column_name(pStmt, j);
      char zC[8];
      if( zCol==0 ) zCol = "";
      xlsxColName(j, zC);
      sqlite3_snprintf((int)sizeof(zRef), zRef, "%s%lld", zC, r);
      switch( eType ){
        case SQLITE_NULL:
          break;                            /* leave the cell empty */
        case SQLITE_INTEGER: {
          sqlite3_int64 v = sqlite3_column_int64(pStmt, j);
          if( v>=-9007199254740992LL && v<=9007199254740992LL ){
            sbAppendF(out, "<c r=\"%s\"><v>%lld</v></c>", zRef, v);
          }else{
            xlsxWarn("integer %lld in row %lld, column \"%s\" of table "
                     "\"%s\" exceeds Excel's exact numeric range (+/-2^53); "
                     "written as text", v, r, zCol, zTable);
            char zNum[32];
            sqlite3_snprintf((int)sizeof(zNum), zNum, "%lld", v);
            xlsxWriteStrCell(out, zRef, zNum, strlen(zNum), 0,
                             zTable, zCol, r);
          }
          break;
        }
        case SQLITE_FLOAT: {
          double d = sqlite3_column_double(pStmt, j);
          if( d==d && d>=-DBL_MAX && d<=DBL_MAX ){       /* finite */
            sbAppendF(out, "<c r=\"%s\"><v>%!.15g</v></c>", zRef, d);
          }else{
            xlsxWarn("non-finite value in row %lld, column \"%s\" of "
                     "table \"%s\"; written as text", r, zCol, zTable);
            xlsxWriteStrCell(out, zRef,
                             d>0 ? "inf" : (d<0 ? "-inf" : "nan"),
                             d!=d ? 3 : (d<0 ? 4 : 3), 0, zTable, zCol, r);
          }
          break;
        }
        case SQLITE_TEXT: {
          const unsigned char *zV = sqlite3_column_text(pStmt, j);
          int nb = sqlite3_column_bytes(pStmt, j);
          if( zV==0 ) nb = 0;
          xlsxWriteStrCell(out, zRef, (const char*)zV,
                           nb<0 ? 0 : (size_t)nb, 0, zTable, zCol, r);
          break;
        }
        case SQLITE_BLOB: {
          const void *zV = sqlite3_column_blob(pStmt, j);
          int nb = sqlite3_column_bytes(pStmt, j);
          sbAppendF(out, "<c r=\"%s\" t=\"inlineStr\"><is><t>", zRef);
          xlsxAppendBlobHex(out, zV, nb<0 ? 0 : (size_t)nb,
                            zTable, zCol, r);
          sbAppend(out, "</t></is></c>");
          break;
        }
        default:
          break;
      }
    }
    sbAppend(out, "</row>");
    if( out->oom ) goto oom_exit;
  }
  if( rc!=SQLITE_DONE ){
    *pzErr = sqlite3_mprintf("reading table \"%s\": %s",
                             zTable, sqlite3_errmsg(db));
    goto exit_fail;
  }

  sbAppend(out, "</sheetData>");

  /* ---- patch the dimension placeholder with the real range ---- */
  sqlite3_snprintf((int)sizeof(zRef), zRef, "A1:%s%lld",
                   zLastCol, nDataRows+1);
  if( out->z && strlen(zRef)<=13 ){
    size_t nRef = strlen(zRef);
    memmove(out->z + dimOff + nRef,      /* shift the tail left  */
            out->z + dimOff + 13,
            out->n - dimOff - 13 + 1);   /* include the NUL      */
    memcpy(out->z + dimOff, zRef, nRef);
    out->n -= 13 - nRef;
  }

  /* ---- auto-filter over the (bold) header row ---- */
  sbAppendF(out, "<autoFilter ref=\"A1:%s1\"/></worksheet>", zLastCol);

  if( out->oom ) goto oom_exit;
  sqlite3_finalize(pStmt);
  return SQLITE_OK;

oom_exit:
  sbFree(&sql);
  sqlite3_finalize(pStmt);
  *pzErr = 0;
  return SQLITE_NOMEM;

exit_fail:
  sqlite3_finalize(pStmt);
  sbFree(&pEnt->data);                   /* leave empty for reuse */
  sbInit(&pEnt->data);
  if( *pzErr==0 ) *pzErr = sqlite3_mprintf("unknown error");
  return rc!=SQLITE_OK ? rc : SQLITE_ERROR;
}

/* Escape a (pre-sanitized, hence valid UTF-8) attribute value. */
static void sbAppendXmlAttr(StrBuf *p, const char *z){
  const char *pT;
  for(pT=z; *pT; pT++){
    switch( *pT ){
      case '&':  sbAppend(p, "&amp;");  break;
      case '<':  sbAppend(p, "&lt;");   break;
      case '>':  sbAppend(p, "&gt;");   break;
      case '"':  sbAppend(p, "&quot;"); break;
      case '\'': sbAppend(p, "&apos;"); break;
      default:   sbAppendN(p, pT, 1);   break;
    }
  }
}

/*************************************************************************
** The five "container" parts of the package.
**
** aEnt[0..4] are filled here; aEnt[5..] already hold the worksheets.
*/
static int xlsxBuildParts(XlsxEntry *aEnt, int nSheet, char **azSheet){
  StrBuf *p;
  int i;

  /* [Content_Types].xml */
  aEnt[0].zName = sqlite3_mprintf("%s", "[Content_Types].xml");
  p = &aEnt[0].data;
  sbAppend(p, XLSX_XML_DECL
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
  for(i=0; i<nSheet; i++){
    sbAppendF(p, "<Override PartName=\"/xl/worksheets/sheet%d.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument."
      "spreadsheetml.worksheet+xml\"/>", i+1);
  }
  sbAppend(p, "<Override PartName=\"/xl/styles.xml\" ContentType=\""
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
    "</Types>");

  /* _rels/.rels */
  aEnt[1].zName = sqlite3_mprintf("%s", "_rels/.rels");
  p = &aEnt[1].data;
  sbAppend(p, XLSX_XML_DECL
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    "<Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
    "Target=\"xl/workbook.xml\"/></Relationships>");

  /* xl/workbook.xml */
  aEnt[2].zName = sqlite3_mprintf("%s", "xl/workbook.xml");
  p = &aEnt[2].data;
  sbAppend(p, XLSX_XML_DECL
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
    "<bookViews><workbookView/></bookViews><sheets>");
  for(i=0; i<nSheet; i++){
    sbAppend(p, "<sheet name=\"");
    sbAppendXmlAttr(p, azSheet[i]);
    sbAppendF(p, "\" sheetId=\"%d\" r:id=\"rId%d\"/>", i+1, i+1);
  }
  sbAppend(p, "</sheets></workbook>");

  /* xl/_rels/workbook.xml.rels */
  aEnt[3].zName = sqlite3_mprintf("%s", "xl/_rels/workbook.xml.rels");
  p = &aEnt[3].data;
  sbAppend(p, XLSX_XML_DECL
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
  for(i=0; i<nSheet; i++){
    sbAppendF(p, "<Relationship Id=\"rId%d\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet%d.xml\"/>", i+1, i+1);
  }
  sbAppendF(p, "<Relationship Id=\"rId%d\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
    "Target=\"styles.xml\"/></Relationships>", nSheet+1);

  /* xl/styles.xml -- font 1 is bold and becomes cell format (xf) 1,
  ** which is what the header-row cells reference with s="1". */
  aEnt[4].zName = sqlite3_mprintf("%s", "xl/styles.xml");
  p = &aEnt[4].data;
  sbAppend(p, XLSX_XML_DECL
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
    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" "
      "applyFont=\"1\"/>"
    "</cellXfs>"
    "<cellStyles count=\"1\">"
    "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
    "</cellStyles>"
    "</styleSheet>");

  for(i=0; i<5; i++){
    if( aEnt[i].zName==0 || aEnt[i].data.oom ) return SQLITE_NOMEM;
  }
  return SQLITE_OK;
}

/*************************************************************************
** Build the .xlsx ZIP archive with the zipfile() aggregate and write it
** out.  The entries are fed to the aggregate through a CTE:
**
**   WITH e(n,d) AS (VALUES (?,?),(?,?),...) SELECT zipfile(n,d) FROM e
**
** When the connection's bound-parameter limit is too small for that many
** entries, a temporary table is used instead.
*/
static int xlsxWriteFile(const char *zFile, const void *pData, size_t n,
                         char **pzErr){
  FILE *f = fopen(zFile, "wb");
  if( f==0 ){
    *pzErr = sqlite3_mprintf("cannot open \"%s\" for writing: %s",
                             zFile, strerror(errno));
    return SQLITE_ERROR;
  }
  if( n>0 && fwrite(pData, 1, n, f)!=n ){
    *pzErr = sqlite3_mprintf("write error on \"%s\"", zFile);
    fclose(f);
    remove(zFile);
    return SQLITE_ERROR;
  }
  if( fclose(f)!=0 ){
    *pzErr = sqlite3_mprintf("write error on \"%s\": %s",
                             zFile, strerror(errno));
    remove(zFile);
    return SQLITE_ERROR;
  }
  return SQLITE_OK;
}

static int xlsxZipEntryTooBig(XlsxEntry *aEnt, int nEnt, char **pzErr){
  int i;
  for(i=0; i<nEnt; i++){
    if( aEnt[i].data.n > (size_t)0x7fffffff ){
      *pzErr = sqlite3_mprintf("part \"%s\" exceeds 2 GiB; "
        "the workbook is too large to export", aEnt[i].zName);
      return 1;
    }
  }
  return 0;
}

static int xlsxBuildZipFile(sqlite3 *db, XlsxEntry *aEnt, int nEnt,
                            const char *zFile, char **pzErr){
  sqlite3_stmt *pStmt = 0;
  StrBuf q;
  int rc = SQLITE_OK;
  int i;
  int mxVar = sqlite3_limit(db, SQLITE_LIMIT_VARIABLE_NUMBER, -1);
  int bTemp = (nEnt*2 > mxVar);

  *pzErr = 0;
  sbInit(&q);

  if( xlsxZipEntryTooBig(aEnt, nEnt, pzErr) ){
    sbFree(&q);
    return SQLITE_ERROR;
  }

  if( bTemp ){
    /* Feed the aggregate through a temp table instead of a CTE. */
    char *zExecErr = 0;
    rc = sqlite3_exec(db,
      "DROP TABLE IF EXISTS temp." XLSX_ZIPTEMP ";"
      "CREATE TEMP TABLE " XLSX_ZIPTEMP "(n TEXT,d BLOB);",
      0, 0, &zExecErr);
    if( rc!=SQLITE_OK ){
      *pzErr = sqlite3_mprintf("temp table: %s",
                               zExecErr ? zExecErr : "?");
      sqlite3_free(zExecErr);
      sbFree(&q);
      return rc;
    }
    rc = sqlite3_prepare_v2(db,
      "INSERT INTO temp." XLSX_ZIPTEMP "(n,d) VALUES(?1,?2)",
      -1, &pStmt, 0);
    if( rc==SQLITE_OK ){
      for(i=0; i<nEnt && rc==SQLITE_OK; i++){
        sqlite3_bind_text(pStmt, 1, aEnt[i].zName, -1, SQLITE_STATIC);
        sqlite3_bind_blob(pStmt, 2, aEnt[i].data.z,
                          (int)aEnt[i].data.n, SQLITE_STATIC);
        rc = sqlite3_step(pStmt);
        if( rc==SQLITE_DONE ) rc = SQLITE_OK;
        sqlite3_reset(pStmt);
      }
      if( rc==SQLITE_OK && sqlite3_errcode(db)==SQLITE_NOMEM ) rc = SQLITE_NOMEM;
    }
    sqlite3_finalize(pStmt);
    pStmt = 0;
    if( rc!=SQLITE_OK ){
      *pzErr = sqlite3_mprintf("cannot feed zipfile(): %s", sqlite3_errmsg(db));
      sbFree(&q);
      sqlite3_exec(db, "DROP TABLE IF EXISTS temp." XLSX_ZIPTEMP, 0, 0, 0);
      return rc;
    }
    sbAppend(&q, "SELECT zipfile(n,d) FROM temp." XLSX_ZIPTEMP);
  }else{
    sbAppend(&q, "WITH e(n,d) AS (VALUES ");
    for(i=0; i<nEnt; i++){
      sbAppend(&q, i ? ",(?,?)" : "(?,?)");
    }
    sbAppend(&q, ") SELECT zipfile(n,d) FROM e");
  }
  if( q.oom ){ rc = SQLITE_NOMEM; goto oom; }

  rc = sqlite3_prepare_v2(db, q.z, -1, &pStmt, 0);
  sbFree(&q);
  if( rc!=SQLITE_OK ){
    const char *zMsg = sqlite3_errmsg(db);
    if( strstr(zMsg, "zipfile")!=0 ){
      *pzErr = sqlite3_mprintf("%s (the zipfile extension must be loaded "
        "before xlsx_export can build the workbook)", zMsg);
    }else{
      *pzErr = sqlite3_mprintf("cannot build the XLSX container: %s", zMsg);
    }
    goto cleanup;
  }

  if( !bTemp ){
    for(i=0; i<nEnt && rc==SQLITE_OK; i++){
      rc = sqlite3_bind_text(pStmt, 2*i+1, aEnt[i].zName, -1, SQLITE_STATIC);
      if( rc==SQLITE_OK ){
        rc = sqlite3_bind_blob(pStmt, 2*i+2, aEnt[i].data.z,
                               (int)aEnt[i].data.n, SQLITE_STATIC);
      }
    }
    if( rc!=SQLITE_OK ){
      *pzErr = sqlite3_mprintf("bind: %s", sqlite3_errmsg(db));
      goto cleanup;
    }
  }

  rc = sqlite3_step(pStmt);
  if( rc==SQLITE_ROW ){
    const void *pZip = sqlite3_column_blob(pStmt, 0);
    int nZip = sqlite3_column_bytes(pStmt, 0);
    if( pZip==0 || nZip<=0 ){
      rc = SQLITE_ERROR;
      *pzErr = sqlite3_mprintf("zipfile() produced an empty archive");
    }else{
      rc = xlsxWriteFile(zFile, pZip, (size_t)nZip, pzErr);
    }
  }else{
    rc = SQLITE_ERROR;
    *pzErr = sqlite3_mprintf("zipfile() failed: %s", sqlite3_errmsg(db));
  }

cleanup:
  sqlite3_finalize(pStmt);
  if( bTemp ){
    sqlite3_exec(db, "DROP TABLE IF EXISTS temp." XLSX_ZIPTEMP, 0, 0, 0);
  }
  return rc;

oom:
  sbFree(&q);
  sqlite3_finalize(pStmt);
  if( bTemp ){
    sqlite3_exec(db, "DROP TABLE IF EXISTS temp." XLSX_ZIPTEMP, 0, 0, 0);
  }
  *pzErr = 0;
  return SQLITE_NOMEM;
}

/*************************************************************************
** xlsx_export(FILENAME [, TABLE...])
*/
static void xlsxResultError(sqlite3_context *ctx, const char *zMsg){
  if( zMsg==0 ) zMsg = "xlsx_export: out of memory";
  sqlite3_result_error(ctx, zMsg, -1);
}

static void xlsxExportFunc(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  sqlite3 *db = sqlite3_context_db_handle(ctx);
  char *zErr = 0;
  char **azTable = 0;      /* table names to export */
  char **azSheet = 0;      /* sanitized sheet names (successes only) */
  XlsxEntry *aEnt = 0;
  int nTable = 0, nSheet = 0;
  int i;
  int bAuto = (argc<=1);
  int rc = SQLITE_OK;
  const char *zFile;

  if( argc<1 ){
    xlsxResultError(ctx,
      "xlsx_export: wrong number of arguments; expected "
      "xlsx_export(FILENAME[, TABLE, ...])");
    return;
  }
  if( sqlite3_value_type(argv[0])==SQLITE_NULL ){
    xlsxResultError(ctx, "xlsx_export: the output file name must not be NULL");
    return;
  }
  zFile = (const char*)sqlite3_value_text(argv[0]);
  if( zFile==0 || zFile[0]==0 ){
    xlsxResultError(ctx, "xlsx_export: invalid output file name");
    return;
  }

  /* ---- collect the table names ---- */
  if( bAuto ){
    sqlite3_stmt *pList = 0;
    int bOom = 0;
    rc = sqlite3_prepare_v2(db,
      "SELECT name FROM main.sqlite_master"
      " WHERE type='table' AND name NOT LIKE 'sqlite\\_%' ESCAPE '\\'"
      " ORDER BY name", -1, &pList, 0);
    if( rc!=SQLITE_OK ){
      zErr = sqlite3_mprintf("cannot list tables: %s", sqlite3_errmsg(db));
      goto error_out;
    }
    while( (rc = sqlite3_step(pList))==SQLITE_ROW ){
      const unsigned char *z = sqlite3_column_text(pList, 0);
      char **aNew = (char**)sqlite3_realloc64(azTable,
                       (size_t)(nTable+1)*sizeof(char*));
      if( aNew==0 || z==0 ){ bOom = 1; break; }
      azTable = aNew;
      azTable[nTable] = sqlite3_mprintf("%s", (const char*)z);
      if( azTable[nTable]==0 ){ bOom = 1; break; }
      nTable++;
    }
    sqlite3_finalize(pList);
    if( bOom ){ rc = SQLITE_NOMEM; zErr = 0; goto error_out; }
    if( rc!=SQLITE_DONE ){
      zErr = sqlite3_mprintf("cannot list tables: %s", sqlite3_errmsg(db));
      goto error_out;
    }
    if( nTable==0 ){
      zErr = sqlite3_mprintf("xlsx_export: schema \"main\" contains no tables");
      goto error_out;
    }
  }else{
    nTable = argc-1;
    azTable = (char**)sqlite3_malloc64((size_t)nTable*sizeof(char*));
    if( azTable==0 ){ rc = SQLITE_NOMEM; zErr = 0; goto error_out; }
    memset(azTable, 0, (size_t)nTable*sizeof(char*));
    for(i=0; i<nTable; i++){
      if( sqlite3_value_type(argv[i+1])==SQLITE_NULL ){
        zErr = sqlite3_mprintf("xlsx_export: table argument %d must not "
                               "be NULL", i+2);
        goto error_out;
      }
      azTable[i] = sqlite3_mprintf("%s",
                      (const char*)sqlite3_value_text(argv[i+1]));
      if( azTable[i]==0 || azTable[i][0]==0 ){
        zErr = sqlite3_mprintf("xlsx_export: invalid table argument %d",
                               i+2);
        goto error_out;
      }
    }
  }

  /* ---- allocate ---- */
  azSheet = (char**)sqlite3_malloc64((size_t)nTable*sizeof(char*));
  aEnt = (XlsxEntry*)sqlite3_malloc64(
           ((size_t)nTable+5)*sizeof(XlsxEntry));
  if( azSheet==0 || aEnt==0 ){ rc = SQLITE_NOMEM; zErr = 0; goto error_out; }
  memset(azSheet, 0, (size_t)nTable*sizeof(char*));
  memset(aEnt, 0, ((size_t)nTable+5)*sizeof(XlsxEntry));
  for(i=0; i<nTable+5; i++){
    aEnt[i].zName = 0;
    sbInit(&aEnt[i].data);
    if( aEnt[i].data.oom ){ rc = SQLITE_NOMEM; zErr = 0; goto error_out; }
  }

  /* ---- generate the worksheets ---- */
  for(i=0; i<nTable; i++){
    char zSan[XLSX_NAME_BUF];
    char *zErr2 = 0;
    xlsxSanitizeSheetName(azTable[i], zSan);
    xlsxUniquifySheetName(zSan, azSheet, nSheet);
    rc = xlsxGenSheet(db, &aEnt[5+nSheet], azTable[i], &zErr2);
    if( rc==SQLITE_NOMEM ){
      sqlite3_free(zErr2);
      rc = SQLITE_NOMEM; zErr = 0;
      goto error_out;
    }
    if( rc!=SQLITE_OK ){
      if( bAuto ){
        /* full-schema mode: warn and skip the table */
        xlsxWarn("skipping table \"%s\": %s", azTable[i],
                 zErr2 ? zErr2 : "error");
        sqlite3_free(zErr2);
        continue;
      }
      zErr = zErr2;
      goto error_out;
    }
    aEnt[5+nSheet].zName = sqlite3_mprintf("xl/worksheets/sheet%d.xml",
                                           nSheet+1);
    azSheet[nSheet] = sqlite3_mprintf("%s", zSan);
    if( aEnt[5+nSheet].zName==0 || azSheet[nSheet]==0 ){
      sqlite3_free(zErr2);
      rc = SQLITE_NOMEM; zErr = 0;
      goto error_out;
    }
    nSheet++;
  }
  if( nSheet==0 ){
    zErr = sqlite3_mprintf("xlsx_export: no tables could be exported");
    goto error_out;
  }

  /* ---- container parts + zip + write ---- */
  if( xlsxBuildParts(aEnt, nSheet, azSheet)!=SQLITE_OK ){
    rc = SQLITE_NOMEM; zErr = 0;
    goto error_out;
  }
  rc = xlsxBuildZipFile(db, aEnt, nSheet+5, zFile, &zErr);
  if( rc!=SQLITE_OK ){
    if( rc==SQLITE_NOMEM ) zErr = 0;
    goto error_out;
  }

  sqlite3_result_int64(ctx, nSheet);
  goto cleanup;

error_out:
  xlsxResultError(ctx, zErr);
  sqlite3_free(zErr);

cleanup:
  if( aEnt ){
    for(i=0; i<nTable+5; i++){
      sqlite3_free(aEnt[i].zName);
      sbFree(&aEnt[i].data);
    }
    sqlite3_free(aEnt);
  }
  if( azSheet ){
    for(i=0; i<nSheet; i++) sqlite3_free(azSheet[i]);
    sqlite3_free(azSheet);
  }
  if( azTable ){
    for(i=0; i<nTable; i++) sqlite3_free(azTable[i]);
    sqlite3_free(azTable);
  }
}

/*************************************************************************
** xlsx_export_version()
*/
static void xlsxVersionFunc(sqlite3_context *ctx, int argc,
                            sqlite3_value **argv){
  (void)argc;
  (void)argv;
  sqlite3_result_text(ctx, "2026-11-03 GLM-5.3-Flash (Max)", -1,
                      SQLITE_STATIC);
}

#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsxexport_init(
  sqlite3 *db,
  char **pzErrMsg,
  const sqlite3_api_routines *pApi
){
  int rc;
  SQLITE_EXTENSION_INIT2(pApi);
  (void)pzErrMsg;
  rc = sqlite3_create_function(db, "xlsx_export", -1,
                               XLSXEXPORT_FUNCFLAGS, 0,
                               xlsxExportFunc, 0, 0);
  if( rc==SQLITE_OK ){
    rc = sqlite3_create_function(db, "xlsx_export_version", 0,
                                 SQLITE_UTF8|SQLITE_DETERMINISTIC, 0,
                                 xlsxVersionFunc, 0, 0);
  }
  return rc;
}
