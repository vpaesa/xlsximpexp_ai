/*
** 2026-11-03
**
** xlsximport.c -- A SQLite extension that imports Microsoft XLSX
** (Office Open XML) spreadsheet files into SQLite tables.
**
** The extension provides three SQL functions:
**
**   xlsx_import(FILENAME [, SHEET-SELECTOR...]) -> INT
**
**       Imports an XLSX file.  One table is created for each imported
**       worksheet.  The table name is the (sanitized) sheet name and the
**       column names are taken from the first row of the sheet that
**       contains data.  With no extra arguments every sheet of the
**       workbook is imported.  Otherwise each additional argument selects
**       one sheet, either by its 1-based number (an integer argument or an
**       all-digit text argument) or by its exact name.  The function
**       returns the number of sheets that were imported.
**
**       An existing table of the same name is dropped and recreated, so
**       re-running the import replaces the previous contents.  The whole
**       import runs inside a SAVEPOINT: if any sheet fails, every change
**       made by this call is rolled back.
**
**       Examples:
**         SELECT xlsx_import('/path/data.xlsx');            -- all sheets
**         SELECT xlsx_import('/path/data.xlsx', 'Sheet2');  -- by name
**         SELECT xlsx_import('/path/data.xlsx', 1, 3);      -- by number
**         SELECT xlsx_import('/path/data.xlsx', '2');       -- by number
**
**   xlsx_import_sheetnames(FILENAME) -> table
**
**       A table-valued function that lists the sheets of a workbook.  It
**       returns one row per sheet with the columns "sheet_num" (the
**       1-based sheet number, i.e. the position of the sheet in the
**       workbook) and "sheet_name" (the sheet name):
**
**         SELECT sheet_num, sheet_name
**           FROM xlsx_import_sheetnames('/path/data.xlsx');
**
**   xlsx_import_version() -> TEXT
**
**       Returns the version string of this extension.
**
** Requirements:
**
**   * The "zipfile" extension must be loaded into the same database
**     connection (it is used to read the members of the XLSX zip archive).
**     Recent sqlite3 shells have zipfile built in; otherwise run
**     ".load zipfile" or "SELECT load_extension('zipfile')" first.
**   * The expat XML parser library (development headers at build time).
**   * SQLite 3.30.0 or later at runtime (SQLITE_DIRECTONLY support).
**
** Build (Unix):
**
**   gcc -O2 -fPIC -Wall -Wextra -shared -o xlsximport.so xlsximport.c -lexpat
**
** Compatibility note (important):  expat 2.9.0 (released 2026-10-05) added
** new 64-bit location functions (XML_GetCurrentLineNumber64,
** XML_GetCurrentColumnNumber64, ...) and deprecated the older
** XML_GetCurrentLineNumber()/XML_GetCurrentColumnNumber().  The *64
** functions DO NOT EXIST in any earlier expat release, so hard-coding them
** causes "implicit declaration" build errors on those systems.  This file
** selects the right function at compile time via XML_LINE_NO() /
** XML_COLUMN_NO() below and therefore compiles cleanly against both old
** and new expat versions.
**
** Import behavior notes:
**
**   * Cell values: numbers are stored as INTEGER or REAL, shared/inline
**     strings as TEXT, booleans as 0/1, formula results as their cached
**     values, error cells as their text (e.g. '#DIV/0!').
**   * The header row supplies the column names and is not imported as
**     data; the data rows start with the first row after it.
**   * Table and column names are used verbatim (only unprintable control
**     characters become '_') and are always safely quoted in SQL.  Empty
**     headers fall back to spreadsheet-style column letters ("A", "B",
**     ..., "AA", ...); duplicate names get numbered ("Name_2").
**   * Dates: XLSX stores dates as serial numbers with a display format;
**     the raw serial number is imported (the format is not interpreted).
**   * Empty cells (including gaps before the first and after the last
**     cell of a row) become NULL.  Fully empty rows are skipped.
**   * Very wide sheets: before importing a sheet, xlsx_import() raises
**     the connection's SQLITE_LIMIT_COLUMN and SQLITE_LIMIT_VARIABLE_NUMBER
**     to the sheet's column count -- SQLite silently caps the raise at the
**     build's compile-time hard bounds -- and then re-queries both with
**     sqlite3_limit(db, ..., -1).  The sheet is imported with as many
**     columns as the effective limits allow: up to SQLITE_LIMIT_COLUMN
**     columns, of which the first SQLITE_LIMIT_VARIABLE_NUMBER are bound
**     as host parameters; if even that limit is too small, every row is
**     written with a row-specific INSERT whose values are all inlined as
**     SQL literals (no host parameters needed).  Columns that do not fit are dropped with
**     a warning on stderr instead of failing the import.  The raised
**     column and host-parameter limits are left in effect afterwards --
**     otherwise even "SELECT *" on the imported table would fail -- and
**     the SQL-statement-length limit, which is only needed while the
**     import runs, is restored.
**   * A workbook without any cell data yields a single-column empty table.
*/
#include "sqlite3ext.h"
#ifdef XLSXIMPORT_EXTERN_API
/* Combined single-binary builds (see xlsximport_zip_wrapper.c) link this
** file together with zipfile.c; the API pointer is then defined exactly
** once by the wrapper and only declared here. */
extern const sqlite3_api_routines *sqlite3_api;
#else
SQLITE_EXTENSION_INIT1
#endif
#include <expat.h>
#include <errno.h>
#include <stdarg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#ifndef SQLITE_DIRECTONLY
# define SQLITE_DIRECTONLY 0x000080000   /* SQLite >= 3.30.0 */
#endif

#ifndef UNUSED_PARAMETER
# define UNUSED_PARAMETER(X) (void)(X)
#endif

#define XLSXIMPORT_VERSION "2026-11-03 GLM-5.3-Flash (Max)"

/* Guard on the number of columns scanned in a worksheet: cells beyond
** this 0-based column index are ignored (silently, never an error).
** Real XLSX sheets have at most 16384 columns ("XFD").  The number of
** columns that actually gets imported is decided at run time from the
** database connection's own limits (SQLITE_LIMIT_COLUMN and
** SQLITE_LIMIT_VARIABLE_NUMBER), which xlsx_import() raises as far as
** the SQLite build's compile-time hard bounds allow. */
#ifndef XLSX_SCAN_MAX_COLUMNS
# if defined(XLSX_MAX_COLUMNS)
#  define XLSX_SCAN_MAX_COLUMNS XLSX_MAX_COLUMNS
# else
#  define XLSX_SCAN_MAX_COLUMNS 100000
# endif
#endif

/* Maximum length in bytes of a generated table or column name. */
#define XLSX_MAX_NAME 100

/*
** ---------------------------------------------------------------------
** expat location API compatibility  (the "implicit declaration of
** XML_GetCurrentLineNumber64" fix):  use the *64 functions only when the
** expat header knows about them (expat >= 2.9.0), otherwise fall back to
** the classic functions which exist in every expat release.  Both macros
** yield "unsigned long long" so the same "%llu" format is always valid.
** ---------------------------------------------------------------------
*/
#if XML_MAJOR_VERSION > 2 || (XML_MAJOR_VERSION == 2 && XML_MINOR_VERSION >= 9)
# define XML_LINE_NO(p)  ((unsigned long long)XML_GetCurrentLineNumber64(p))
# define XML_COLUMN_NO(p)((unsigned long long)XML_GetCurrentColumnNumber64(p))
#else
# define XML_LINE_NO(p)  ((unsigned long long)XML_GetCurrentLineNumber(p))
# define XML_COLUMN_NO(p)((unsigned long long)XML_GetCurrentColumnNumber(p))
#endif

/*
** ---------------------------------------------------------------------
** Small growable string buffer.
** ---------------------------------------------------------------------
*/
typedef struct StrBuf {
  char *z;               /* NUL-terminated (once anything was appended) */
  sqlite3_int64 nUsed;   /* bytes used, excluding the NUL terminator */
  sqlite3_int64 nAlloc;  /* bytes allocated, including room for the NUL */
  int oom;               /* set if an allocation has failed */
} StrBuf;

static void strbufInit(StrBuf *p){
  p->z = 0;
  p->nUsed = 0;
  p->nAlloc = 0;
  p->oom = 0;
}

static void strbufFree(StrBuf *p){
  sqlite3_free(p->z);
  strbufInit(p);
}

static void strbufReset(StrBuf *p){
  p->nUsed = 0;
  if( p->z ) p->z[0] = 0;
}

static void strbufAppend(StrBuf *p, const char *z, sqlite3_int64 n){
  if( n<0 ) n = (sqlite3_int64)strlen(z);
  if( p->oom ) return;
  if( p->nUsed + n + 1 > p->nAlloc ){
    sqlite3_int64 nNew = p->nAlloc ? p->nAlloc : 64;
    char *zNew;
    while( nNew < p->nUsed + n + 1 ) nNew *= 2;
    zNew = (char*)sqlite3_realloc64(p->z, (sqlite3_uint64)nNew);
    if( zNew==0 ){
      p->oom = 1;
      return;
    }
    p->z = zNew;
    p->nAlloc = nNew;
  }
  if( n>0 ) memcpy(p->z + p->nUsed, z, (size_t)n);
  p->nUsed += n;
  p->z[p->nUsed] = 0;
}

/* Append printf-style text.  Uses SQLite's own printf implementation. */
static void strbufPrintf(StrBuf *p, const char *zFmt, ...){
  va_list ap;
  char *z;
  va_start(ap, zFmt);
  z = sqlite3_vmprintf(zFmt, ap);
  va_end(ap);
  if( z==0 ){
    p->oom = 1;
    return;
  }
  strbufAppend(p, z, -1);
  sqlite3_free(z);
}

/*
** ---------------------------------------------------------------------
** Misc. small helpers.
** ---------------------------------------------------------------------
*/

/* Copy n bytes of z into a fresh NUL-terminated sqlite3_malloc buffer. */
static char *dupText(const char *z, sqlite3_int64 n){
  char *zNew;
  if( z==0 ) return 0;
  if( n<0 ) n = (sqlite3_int64)strlen(z);
  zNew = (char*)sqlite3_malloc64((sqlite3_uint64)n + 1);
  if( zNew ){
    if( n>0 ) memcpy(zNew, z, (size_t)n);
    zNew[n] = 0;
  }
  return zNew;
}

/* Look up an attribute by its literal (non-namespaced) name. */
static const char *xmlAttr(const XML_Char **azAtts, const char *zAttr){
  int i;
  for(i=0; azAtts[i]; i+=2){
    if( strcmp(azAtts[i], zAttr)==0 ) return azAtts[i+1];
  }
  return 0;
}

/* Look up the relationship id attribute of a <sheet> element.  The
** attribute is normally spelled "r:id" but any "<prefix>:id" (or plain
** "id") is accepted.  Note that "sheetId" is NOT a relationship id. */
static const char *ridAttr(const XML_Char **azAtts){
  int i;
  for(i=0; azAtts[i]; i+=2){
    const char *a = azAtts[i];
    size_t n = strlen(a);
    if( n>=3 && strcmp(a+n-3, ":id")==0 ) return azAtts[i+1];
    if( strcmp(a, "id")==0 ) return azAtts[i+1];
  }
  return 0;
}

/* Convert the column letters of a cell reference ("C5" -> 2, 0-based).
** Returns -1 if the reference is malformed. */
static int colFromRef(const char *zRef){
  int col = 0;
  int i;
  for(i=0; ; i++){
    char c = zRef[i];
    if( c>='a' && c<='z' ) c = (char)(c - 'a' + 'A');
    if( c<'A' || c>'Z' ) break;
    if( col > 100000 ) return -1;
    col = col*26 + (c - 'A' + 1);
  }
  if( i==0 ) return -1;
  col--;
  return col<0 ? -1 : col;
}

/* Inverse of colFromRef(): the spreadsheet-style letter of a 0-based
** column index (0 -> "A", 25 -> "Z", 26 -> "AA", ...). */
static void colLetter(int i, char *zOut){
  char zTmp[8];
  int n = 0;
  int k;
  i++;
  while( i>0 && n<7 ){
    i--;
    zTmp[n++] = (char)('A' + (i % 26));
    i /= 26;
  }
  for(k=0; k<n; k++) zOut[k] = zTmp[n-1-k];
  zOut[n] = 0;
}

/* True if z consists of n bytes of an optionally signed decimal integer. */
static int looksInteger(const char *z, sqlite3_int64 n){
  sqlite3_int64 i = 0;
  if( n==0 ) return 0;
  if( z[0]=='-' || z[0]=='+' ) i = 1;
  if( i>=n ) return 0;
  for(; i<n; i++){
    if( z[i]<'0' || z[i]>'9' ) return 0;
  }
  return 1;
}

/* True if z is a non-empty string of decimal digits only. */
static int allDigits(const char *z){
  int i;
  if( z==0 || z[0]==0 ) return 0;
  for(i=0; z[i]; i++){
    if( z[i]<'0' || z[i]>'9' ) return 0;
  }
  return 1;
}

/* True if any earlier entry of az[0..n-1] equals z (case-insensitive,
** like SQLite identifier comparison for ASCII). */
static int nameInList(char **az, int n, const char *z){
  int i;
  for(i=0; i<n; i++){
    if( az[i] && sqlite3_stricmp(az[i], z)==0 ) return 1;
  }
  return 0;
}

/* Turn raw sheet/header text into a usable SQL identifier.  The text is
** kept verbatim (SQL quoting happens elsewhere); only unprintable control
** characters become '_'.  The result is capped at XLSX_MAX_NAME bytes on
** a UTF-8 character boundary.  Returns a fresh buffer, or NULL if there
** is nothing usable (the caller then generates a fallback name). */
static char *sanitizeIdent(const char *zSrc, sqlite3_int64 nSrc){
  char *zOut;
  sqlite3_int64 i, j = 0;

  if( zSrc==0 ) return 0;
  if( nSrc<0 ) nSrc = (sqlite3_int64)strlen(zSrc);
  zOut = (char*)sqlite3_malloc64((sqlite3_uint64)nSrc + 1);
  if( zOut==0 ) return 0;
  for(i=0; i<nSrc; i++){
    unsigned char c = (unsigned char)zSrc[i];
    if( c<0x20 || c==0x7f ) c = '_';
    zOut[j++] = (char)c;
  }
  if( j>XLSX_MAX_NAME ){
    j = XLSX_MAX_NAME;
    while( j>0 && (((unsigned char)zOut[j]) & 0xC0)==0x80 ) j--;
  }
  zOut[j] = 0;
  if( j==0 ){
    sqlite3_free(zOut);
    return 0;
  }
  return zOut;
}

/*
** ---------------------------------------------------------------------
** Reading members of the XLSX zip archive through the zipfile extension.
** ---------------------------------------------------------------------
*/
typedef struct Blob {
  char *z;               /* content (sqlite3_malloc64, NUL-terminated) */
  sqlite3_int64 n;       /* content size in bytes */
} Blob;

/*
** Read the (decompressed) content of one archive member.  *pnFound is set
** to 1 if the member exists.  On success pBlob holds malloc'd content
** which may be zero bytes; the caller must sqlite3_free() pBlob->z.
*/
static int readZipEntry(
  sqlite3 *db,
  const char *zArchive,
  const char *zName,
  Blob *pBlob,
  int *pnFound,
  char **pzErr
){
  sqlite3_stmt *pStmt = 0;
  int rc;

  memset(pBlob, 0, sizeof(*pBlob));
  *pnFound = 0;
  rc = sqlite3_prepare_v2(db,
        "SELECT data FROM zipfile(?1) WHERE name=?2 LIMIT 1", -1, &pStmt, 0);
  if( rc!=SQLITE_OK ){
    const char *zDbErr = sqlite3_errmsg(db);
    if( zDbErr && strstr(zDbErr, "zipfile")!=0
     && (strstr(zDbErr, "no such module")!=0
         || strstr(zDbErr, "no such table")!=0) ){
      *pzErr = sqlite3_mprintf(
          "xlsx_import: the \"zipfile\" extension is required but not loaded "
          "into this database connection (%s).  Load it first, e.g. with "
          "\".load zipfile\" or \"SELECT load_extension('zipfile')\"; recent "
          "sqlite3 shells have it built in", zDbErr);
    }else{
      *pzErr = sqlite3_mprintf(
          "xlsx_import: cannot query zip archive \"%s\": %s",
          zArchive, zDbErr ? zDbErr : "unknown error");
    }
    return SQLITE_ERROR;
  }
  sqlite3_bind_text(pStmt, 1, zArchive, -1, SQLITE_STATIC);
  sqlite3_bind_text(pStmt, 2, zName, -1, SQLITE_STATIC);
  rc = sqlite3_step(pStmt);
  if( rc==SQLITE_ROW ){
    const void *pData = sqlite3_column_blob(pStmt, 0);
    sqlite3_int64 n = sqlite3_column_bytes(pStmt, 0);
    *pnFound = 1;
    if( n>0 ){
      pBlob->z = (char*)sqlite3_malloc64((sqlite3_uint64)n + 1);
      if( pBlob->z==0 ){
        rc = SQLITE_NOMEM;
      }else{
        memcpy(pBlob->z, pData, (size_t)n);
        pBlob->z[n] = 0;
        pBlob->n = n;
        rc = SQLITE_OK;
      }
    }else{
      rc = SQLITE_OK;
    }
  }else if( rc==SQLITE_DONE ){
    rc = SQLITE_OK;
  }else{
    *pzErr = sqlite3_mprintf(
        "xlsx_import: error reading \"%s\" from \"%s\": %s",
        zName, zArchive, sqlite3_errmsg(db));
    rc = SQLITE_ERROR;
  }
  sqlite3_finalize(pStmt);
  return rc;
}

/*
** ---------------------------------------------------------------------
** Generic expat driver.
** ---------------------------------------------------------------------
*/

/*
** Parse a complete in-memory XML document.  Any of the handler pointers
** may be NULL.  On error a human readable message (including the source
** line and column, see the expat compatibility note at the top of this
** file) is appended to pErr and SQLITE_ERROR is returned.
*/
static int xmlParseBuffer(
  Blob *pXml,
  XML_StartElementHandler xStart,
  XML_EndElementHandler xEnd,
  XML_CharacterDataHandler xChar,
  void *pUser,
  StrBuf *pErr
){
  XML_Parser pParser;
  enum XML_Status st;
  const char *zData;

  if( pXml->n > 2147483647 ){
    strbufAppend(pErr, "XML part is larger than 2GB and cannot be parsed", -1);
    return SQLITE_ERROR;
  }
  pParser = XML_ParserCreate(0);
  if( pParser==0 ){
    strbufAppend(pErr, "out of memory: could not allocate XML parser", -1);
    return SQLITE_NOMEM;
  }
  XML_SetUserData(pParser, pUser);
  if( xStart || xEnd ) XML_SetElementHandler(pParser, xStart, xEnd);
  if( xChar ) XML_SetCharacterDataHandler(pParser, xChar);
  zData = pXml->z ? pXml->z : "";
  st = XML_Parse(pParser, zData, (int)pXml->n, 1);
  if( st==XML_STATUS_ERROR ){
    const char *zXmlErr = XML_ErrorString(XML_GetErrorCode(pParser));
    strbufPrintf(pErr, "XML syntax error: %s at line %llu, column %llu",
                 zXmlErr ? zXmlErr : "unknown error",
                 XML_LINE_NO(pParser), XML_COLUMN_NO(pParser));
    XML_ParserFree(pParser);
    return SQLITE_ERROR;
  }
  XML_ParserFree(pParser);
  return SQLITE_OK;
}

/* Like xmlParseBuffer(), but reports errors with the name of the
** workbook part being parsed. */
static int parseXmlPart(
  Blob *pXml,
  XML_StartElementHandler xStart,
  XML_EndElementHandler xEnd,
  XML_CharacterDataHandler xChar,
  void *pUser,
  const char *zPartName,
  char **pzErr
){
  StrBuf err;
  int rc;
  strbufInit(&err);
  rc = xmlParseBuffer(pXml, xStart, xEnd, xChar, pUser, &err);
  if( rc!=SQLITE_OK && *pzErr==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: error parsing %s: %s",
                             zPartName, err.z ? err.z : "out of memory");
  }
  strbufFree(&err);
  return rc;
}

/*
** ---------------------------------------------------------------------
** Workbook parts: xl/workbook.xml and xl/_rels/workbook.xml.rels
** ---------------------------------------------------------------------
*/
typedef struct SheetInfo {
  char *zName;           /* sheet name from workbook.xml */
  char *zRid;            /* relationship id (r:id attribute) or NULL */
  char *zPath;           /* archive path of the worksheet XML */
} SheetInfo;

typedef struct SheetList {
  int n;
  int nAlloc;
  SheetInfo *a;
} SheetList;

typedef struct Rel {
  char *zId;
  char *zTarget;
} Rel;

typedef struct Rels {
  int n;
  int nAlloc;
  Rel *a;
} Rels;

static void freeSheetList(SheetList *pSL){
  int i;
  for(i=0; i<pSL->n; i++){
    sqlite3_free(pSL->a[i].zName);
    sqlite3_free(pSL->a[i].zRid);
    sqlite3_free(pSL->a[i].zPath);
  }
  sqlite3_free(pSL->a);
  memset(pSL, 0, sizeof(*pSL));
}

static int sheetListAppend(SheetList *pSL, const char *zName, const char *zRid){
  SheetInfo *p;
  if( pSL->n==pSL->nAlloc ){
    int nNew = pSL->nAlloc ? pSL->nAlloc*2 : 8;
    SheetInfo *aNew = (SheetInfo*)sqlite3_realloc64(
        pSL->a, sizeof(SheetInfo)*(sqlite3_uint64)nNew);
    if( aNew==0 ) return SQLITE_NOMEM;
    pSL->a = aNew;
    pSL->nAlloc = nNew;
  }
  p = &pSL->a[pSL->n];
  memset(p, 0, sizeof(*p));
  p->zName = dupText(zName ? zName : "", -1);
  p->zRid = zRid ? dupText(zRid, -1) : 0;
  if( p->zName==0 || (zRid && p->zRid==0) ){
    sqlite3_free(p->zName);
    sqlite3_free(p->zRid);
    p->zName = 0;
    p->zRid = 0;
    return SQLITE_NOMEM;
  }
  pSL->n++;
  return SQLITE_OK;
}

static int relsAppend(Rels *p, const char *zId, const char *zTarget){
  if( p->n==p->nAlloc ){
    int nNew = p->nAlloc ? p->nAlloc*2 : 8;
    Rel *aNew = (Rel*)sqlite3_realloc64(p->a, sizeof(Rel)*(sqlite3_uint64)nNew);
    if( aNew==0 ) return SQLITE_NOMEM;
    p->a = aNew;
    p->nAlloc = nNew;
  }
  p->a[p->n].zId = dupText(zId, -1);
  p->a[p->n].zTarget = dupText(zTarget, -1);
  if( p->a[p->n].zId==0 || p->a[p->n].zTarget==0 ){
    sqlite3_free(p->a[p->n].zId);
    sqlite3_free(p->a[p->n].zTarget);
    p->a[p->n].zId = 0;
    p->a[p->n].zTarget = 0;
    return SQLITE_NOMEM;
  }
  p->n++;
  return SQLITE_OK;
}

static void freeRels(Rels *p){
  int i;
  for(i=0; i<p->n; i++){
    sqlite3_free(p->a[i].zId);
    sqlite3_free(p->a[i].zTarget);
  }
  sqlite3_free(p->a);
  memset(p, 0, sizeof(*p));
}

static const char *relsFind(Rels *p, const char *zId){
  int i;
  for(i=0; i<p->n; i++){
    if( strcmp(p->a[i].zId, zId)==0 ) return p->a[i].zTarget;
  }
  return 0;
}

typedef struct WorkbookCtx {
  SheetList *pSL;
  int failed;
} WorkbookCtx;

typedef struct RelsCtx {
  Rels *pRels;
  int failed;
} RelsCtx;

static void wbStart(void *pUser, const XML_Char *zElem, const XML_Char **azAtts){
  WorkbookCtx *p = (WorkbookCtx*)pUser;
  const char *zName, *zRid;
  if( p->failed ) return;
  if( strcmp(zElem, "sheet")!=0 ) return;
  zName = xmlAttr(azAtts, "name");
  zRid = ridAttr(azAtts);
  if( sheetListAppend(p->pSL, zName, zRid)!=SQLITE_OK ) p->failed = 1;
}

static void relsStart(void *pUser, const XML_Char *zElem, const XML_Char **azAtts){
  RelsCtx *p = (RelsCtx*)pUser;
  const char *zId, *zTarget, *zType;
  size_t nT;
  if( p->failed ) return;
  if( strcmp(zElem, "Relationship")!=0 ) return;
  zId = xmlAttr(azAtts, "Id");
  zTarget = xmlAttr(azAtts, "Target");
  zType = xmlAttr(azAtts, "Type");
  if( zId==0 || zTarget==0 ) return;
  if( zType ){
    nT = strlen(zType);
    if( nT<10 || strcmp(zType+nT-10, "/worksheet")!=0 ) return;
  }
  if( relsAppend(p->pRels, zId, zTarget)!=SQLITE_OK ) p->failed = 1;
}

/*
** Build the list of sheets of workbook zFile.  On success pSL holds the
** sheet names (in workbook order) and the archive path of each worksheet
** part.  The caller must eventually freeSheetList(pSL).
*/
static int loadSheetList(sqlite3 *db, const char *zFile, SheetList *pSL, char **pzErr){
  Blob xml;
  int found = 0;
  int rc, i;
  WorkbookCtx wb;
  Rels rels;

  memset(pSL, 0, sizeof(*pSL));
  memset(&rels, 0, sizeof(rels));
  memset(&wb, 0, sizeof(wb));

  rc = readZipEntry(db, zFile, "xl/workbook.xml", &xml, &found, pzErr);
  if( rc ) return rc;
  if( !found || xml.n==0 ){
    *pzErr = sqlite3_mprintf(
        "xlsx_import: \"%s\" is not a valid XLSX file: mandatory member "
        "\"xl/workbook.xml\" is missing or empty", zFile);
    sqlite3_free(xml.z);
    return SQLITE_ERROR;
  }
  wb.pSL = pSL;
  rc = parseXmlPart(&xml, wbStart, 0, 0, &wb, "xl/workbook.xml", pzErr);
  sqlite3_free(xml.z);
  if( rc ) return rc;
  if( wb.failed ){
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory parsing xl/workbook.xml");
    return SQLITE_NOMEM;
  }
  if( pSL->n==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: workbook \"%s\" contains no sheets", zFile);
    return SQLITE_ERROR;
  }

  /* Optional relationship part: maps r:id -> worksheet part. */
  rc = readZipEntry(db, zFile, "xl/_rels/workbook.xml.rels", &xml, &found, pzErr);
  if( rc ) return rc;
  if( found && xml.n>0 ){
    RelsCtx rctx;
    memset(&rctx, 0, sizeof(rctx));
    rctx.pRels = &rels;
    rc = parseXmlPart(&xml, relsStart, 0, 0, &rctx,
                      "xl/_rels/workbook.xml.rels", pzErr);
    sqlite3_free(xml.z);
    if( rc ){
      freeRels(&rels);
      return rc;
    }
    if( rctx.failed ){
      freeRels(&rels);
      *pzErr = sqlite3_mprintf(
          "xlsx_import: out of memory parsing xl/_rels/workbook.xml.rels");
      return SQLITE_NOMEM;
    }
  }else{
    sqlite3_free(xml.z);
  }

  /* Resolve each sheet's relationship id to a worksheet part path. */
  for(i=0; i<pSL->n; i++){
    SheetInfo *pS = &pSL->a[i];
    const char *zTarget = (pS->zRid && rels.n) ? relsFind(&rels, pS->zRid) : 0;
    if( zTarget ){
      if( zTarget[0]=='/' ){
        pS->zPath = dupText(zTarget+1, -1);          /* package-absolute */
      }else if( strncmp(zTarget, "xl/", 3)==0 ){
        pS->zPath = dupText(zTarget, -1);            /* already prefixed */
      }else{
        pS->zPath = sqlite3_mprintf("xl/%s", zTarget);
      }
    }else{
      /* No usable relationship: fall back to the conventional name. */
      pS->zPath = sqlite3_mprintf("xl/worksheets/sheet%d.xml", i+1);
    }
    if( pS->zPath==0 ){
      freeRels(&rels);
      *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
      return SQLITE_NOMEM;
    }
  }
  freeRels(&rels);
  return SQLITE_OK;
}

/*
** ---------------------------------------------------------------------
** Shared strings: xl/sharedStrings.xml
** ---------------------------------------------------------------------
*/
typedef struct StrEntry {
  char *z;               /* UTF-8 text, NUL-terminated */
  sqlite3_int64 n;       /* text length in bytes */
} StrEntry;

typedef struct SharedStrings {
  int n;
  int nAlloc;
  StrEntry *a;
} SharedStrings;

static void freeSharedStrings(SharedStrings *pSS){
  int i;
  for(i=0; i<pSS->n; i++) sqlite3_free(pSS->a[i].z);
  sqlite3_free(pSS->a);
  memset(pSS, 0, sizeof(*pSS));
}

typedef struct SSCtx {
  SharedStrings *pSS;
  StrBuf si;             /* accumulated text of the current <si> */
  StrBuf t;              /* accumulated text of the current <t> */
  int inT;               /* nesting depth inside <t> */
  int nRPh;              /* nesting depth inside <rPh> (phonetic text) */
  int failed;
} SSCtx;

static void ssStart(void *pUser, const XML_Char *zElem, const XML_Char **azAtts){
  SSCtx *p = (SSCtx*)pUser;
  UNUSED_PARAMETER(azAtts);
  if( p->failed ) return;
  if( strcmp(zElem, "si")==0 ){
    strbufReset(&p->si);
  }else if( strcmp(zElem, "t")==0 ){
    p->inT++;
    strbufReset(&p->t);
  }else if( strcmp(zElem, "rPh")==0 ){
    p->nRPh++;           /* ignore phonetic (furigana) annotations */
  }
}

static void ssChar(void *pUser, const XML_Char *z, int len){
  SSCtx *p = (SSCtx*)pUser;
  if( p->failed ) return;
  if( p->inT>0 && p->nRPh==0 ) strbufAppend(&p->t, z, len);
}

static void ssEnd(void *pUser, const XML_Char *zElem){
  SSCtx *p = (SSCtx*)pUser;
  SharedStrings *pSS;
  if( p->failed ) return;
  if( strcmp(zElem, "t")==0 ){
    p->inT--;
    if( p->inT==0 && p->nRPh==0 ){
      strbufAppend(&p->si, p->t.z ? p->t.z : "", p->t.nUsed);
      if( p->si.oom ) p->failed = 1;
    }
  }else if( strcmp(zElem, "rPh")==0 ){
    p->nRPh--;
  }else if( strcmp(zElem, "si")==0 ){
    pSS = p->pSS;
    if( pSS->n==pSS->nAlloc ){
      int nNew = pSS->nAlloc ? pSS->nAlloc*2 : 64;
      StrEntry *aNew = (StrEntry*)sqlite3_realloc64(
          pSS->a, sizeof(StrEntry)*(sqlite3_uint64)nNew);
      if( aNew==0 ){
        p->failed = 1;
        return;
      }
      pSS->a = aNew;
      pSS->nAlloc = nNew;
    }
    pSS->a[pSS->n].z = dupText(p->si.z ? p->si.z : "", p->si.nUsed);
    if( pSS->a[pSS->n].z==0 ){
      p->failed = 1;
      return;
    }
    pSS->a[pSS->n].n = p->si.nUsed;
    pSS->n++;
  }
}

/* Load the shared string table.  A missing or empty part is not an
** error (the workbook simply has no shared strings). */
static int loadSharedStrings(sqlite3 *db, const char *zFile, SharedStrings *pSS, char **pzErr){
  Blob xml;
  int found = 0;
  int rc;
  SSCtx sctx;

  memset(pSS, 0, sizeof(*pSS));
  rc = readZipEntry(db, zFile, "xl/sharedStrings.xml", &xml, &found, pzErr);
  if( rc ) return rc;
  if( !found || xml.n==0 ){
    sqlite3_free(xml.z);
    return SQLITE_OK;
  }
  memset(&sctx, 0, sizeof(sctx));
  strbufInit(&sctx.si);
  strbufInit(&sctx.t);
  sctx.pSS = pSS;
  rc = parseXmlPart(&xml, ssStart, ssEnd, ssChar, &sctx,
                    "xl/sharedStrings.xml", pzErr);
  if( rc==SQLITE_OK && sctx.failed ){
    rc = SQLITE_ERROR;
    if( *pzErr==0 ){
      *pzErr = sqlite3_mprintf(
          "xlsx_import: out of memory parsing xl/sharedStrings.xml");
    }
  }
  strbufFree(&sctx.si);
  strbufFree(&sctx.t);
  if( rc!=SQLITE_OK ) freeSharedStrings(pSS);
  sqlite3_free(xml.z);
  return rc;
}

/*
** ---------------------------------------------------------------------
** Worksheet parsing.
** ---------------------------------------------------------------------
*/
#define CELL_NUMBER   0        /* t="n" or no t attribute */
#define CELL_SHARED   1        /* t="s": index into shared strings */
#define CELL_STRING   2        /* t="str": cached formula string result */
#define CELL_INLINE   3        /* t="inlineStr": <is><t>...</t></is> */
#define CELL_BOOL     4        /* t="b" */
#define CELL_ERROR    5        /* t="e" */
#define CELL_DATE     6        /* t="d": ISO 8601 date string */

#define CAP_NONE 0
#define CAP_V    1            /* capturing text of <v> */
#define CAP_T    2            /* capturing text of <is><t> */

/* Typed value of one spreadsheet cell (pass 2).  The conversion rules
** are shared by both insert paths (bound host parameters and inlined
** SQL literals). */
typedef struct CellVal {
  int eType;             /* SQLITE_NULL, SQLITE_INTEGER, SQLITE_FLOAT, SQLITE_TEXT */
  sqlite3_int64 i;       /* eType==SQLITE_INTEGER */
  double r;              /* eType==SQLITE_FLOAT */
  const char *z;         /* eType==SQLITE_TEXT: payload (borrowed) */
  sqlite3_int64 n;       /* TEXT payload length in bytes */
  int bStatic;           /* z points into the shared string table */
} CellVal;

/* Like CellVal, but owning its TEXT payload: used to stage the values
** of columns that are inlined into the INSERT as SQL literals because
** the connection's SQLITE_LIMIT_VARIABLE_NUMBER leaves fewer host
** parameters than the sheet has columns. */
typedef struct InlineVal {
  int eType;
  sqlite3_int64 i;
  double r;
  char *z;               /* owned payload, or NULL */
  sqlite3_int64 n;
} InlineVal;

typedef struct SheetParse SheetParse;
struct SheetParse {
  SharedStrings *pSS;    /* shared string table (borrowed) */
  int mode;              /* 1 = scan for columns, 2 = insert rows */
  sqlite3 *db;           /* mode 2 only */
  sqlite3_stmt *pIns;    /* INSERT statement (mode 2, fast path only) */
  int nCol;              /* number of table columns (mode 2 only) */
  int nBind;             /* host parameters in pIns: min(nCol, var limit) */
  int nInline;           /* nCol - nBind: columns inlined as literals */
  int skipRow;           /* mode 2: row index of the header, not imported */
  char *aBound;          /* per-column "already bound" flags (mode 2) */
  InlineVal *aInline;    /* staged values of the inlined columns (mode 2) */
  StrBuf rowSql;         /* row-specific INSERT text (mode 2 inline path) */
  const char *zTabIns;   /* target table name for rowSql (borrowed) */
  int nCellThisRow;      /* cells with a value in the current row */
  int nRowIns;           /* rows inserted (mode 2) */

  /* expat state */
  int inSheetData;
  int inCell;
  int capture;
  int cellType;
  int iRow;              /* 0-based index of the current row */
  int iCol;              /* 0-based index of the current cell */
  int nextRow;           /* row index to use if the next <row> has no r= */
  int nextCol;           /* column index to use if the next <c> has no r= */
  int headerRow;         /* 0-based index of the header row, -1 unknown */
  int nMaxCol;           /* 1 + highest column index holding a value */
  int hasValue;          /* current cell produced a value */
  int inSkipRow;         /* mode 2: current row is the header row */
  StrBuf vBuf;           /* capture buffer for <v>/<t> text */
  StrBuf cellBuf;        /* accumulated value of the current cell */
  StrEntry *aHdr;        /* header texts by column index (mode 1) */
  int nHdrAlloc;
  int failed;
  char *zErrMsg;
};

/* Record a fatal error (printf style).  The first error wins. */
static void pFail(SheetParse *p, const char *zFmt, ...){
  va_list ap;
  if( p->failed ) return;
  p->failed = 1;
  va_start(ap, zFmt);
  p->zErrMsg = sqlite3_vmprintf(zFmt, ap);
  va_end(ap);
}

/* Move a parse error into *pzErr, adding the sheet name for context. */
static void transferSheetError(SheetParse *p, const char *zSheet, char **pzErr){
  if( p->zErrMsg ){
    *pzErr = sqlite3_mprintf("xlsx_import: sheet \"%s\": %s",
                             zSheet ? zSheet : "", p->zErrMsg);
    sqlite3_free(p->zErrMsg);
    p->zErrMsg = 0;
  }else if( *pzErr==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: sheet \"%s\": out of memory",
                             zSheet ? zSheet : "");
  }
}

static void sheetParseInit(SheetParse *p){
  memset(p, 0, sizeof(*p));
  strbufInit(&p->vBuf);
  strbufInit(&p->cellBuf);
  strbufInit(&p->rowSql);
  p->headerRow = -1;
  p->skipRow = -1;
}

/* Reset the XML scan state so the same SheetParse can be reused for the
** second pass.  Header capture (aHdr) is intentionally kept. */
static void sheetParseClear(SheetParse *p){
  p->inSheetData = 0;
  p->inCell = 0;
  p->capture = CAP_NONE;
  p->cellType = CELL_NUMBER;
  p->iRow = 0;
  p->iCol = 0;
  p->nextRow = 0;
  p->nextCol = 0;
  p->headerRow = -1;
  p->nMaxCol = 0;
  p->hasValue = 0;
  p->failed = 0;
  p->nCellThisRow = 0;
  p->nRowIns = 0;
  p->inSkipRow = 0;
  strbufReset(&p->vBuf);
  strbufReset(&p->cellBuf);
}

static void sheetParseFree(SheetParse *p){
  int i;
  int nStaged = p->nInline>0 ? p->nCol : 0;
  strbufFree(&p->vBuf);
  strbufFree(&p->cellBuf);
  strbufFree(&p->rowSql);
  for(i=0; i<p->nHdrAlloc; i++) sqlite3_free(p->aHdr[i].z);
  sqlite3_free(p->aHdr);
  p->aHdr = 0;
  p->nHdrAlloc = 0;
  sqlite3_free(p->aBound);
  p->aBound = 0;
  for(i=0; i<nStaged; i++) sqlite3_free(p->aInline[i].z);
  sqlite3_free(p->aInline);
  p->aInline = 0;
  p->nInline = 0;
  sqlite3_free(p->zErrMsg);
  p->zErrMsg = 0;
}

static int typeFromT(const char *zT){
  if( zT==0 ) return CELL_NUMBER;
  switch( zT[0] ){
    case 's':
      if( strcmp(zT, "s")==0 ) return CELL_SHARED;
      if( strcmp(zT, "str")==0 ) return CELL_STRING;
      return CELL_NUMBER;
    case 'i':
      return CELL_INLINE;                      /* "inlineStr" */
    case 'b':
      return CELL_BOOL;
    case 'e':
      return CELL_ERROR;
    case 'd':
      return CELL_DATE;
    default:
      return CELL_NUMBER;                      /* "n" and anything unknown */
  }
}

/* Make sure aHdr can hold entries for columns 0..n-1 (zero filled). */
static int hdrEnsure(SheetParse *p, int n){
  if( n > p->nHdrAlloc ){
    int nNew = p->nHdrAlloc ? p->nHdrAlloc : 32;
    StrEntry *aNew;
    while( nNew < n ) nNew *= 2;
    aNew = (StrEntry*)sqlite3_realloc64(
        p->aHdr, sizeof(StrEntry)*(sqlite3_uint64)nNew);
    if( aNew==0 ){
      pFail(p, "out of memory");
      return 1;
    }
    memset(aNew + p->nHdrAlloc, 0,
           sizeof(StrEntry)*(size_t)(nNew - p->nHdrAlloc));
    p->aHdr = aNew;
    p->nHdrAlloc = nNew;
  }
  return 0;
}

/* Resolve the current cell into a header text (shared strings are
** looked up; everything else is taken verbatim). */
static int headerValue(SheetParse *p, const char **pzVal, sqlite3_int64 *pnVal){
  if( p->cellType==CELL_SHARED ){
    const char *z = p->cellBuf.z ? p->cellBuf.z : "";
    sqlite3_int64 idx;
    if( p->cellBuf.nUsed==0 ){
      *pzVal = "";
      *pnVal = 0;
      return SQLITE_OK;
    }
    idx = strtoll(z, 0, 10);
    if( idx<0 || idx>=p->pSS->n ){
      pFail(p, "row %d: shared string index %lld out of range (0..%d); "
               "is xl/sharedStrings.xml missing or incomplete?",
            p->iRow+1, idx, (sqlite3_int64)p->pSS->n - 1);
      return SQLITE_ERROR;
    }
    *pzVal = p->pSS->a[idx].z ? p->pSS->a[idx].z : "";
    *pnVal = p->pSS->a[idx].n;
  }else{
    *pzVal = p->cellBuf.z ? p->cellBuf.z : "";
    *pnVal = p->cellBuf.nUsed;
  }
  return SQLITE_OK;
}

/* Compute the typed value of the cell that just ended (captured text
** in p->cellBuf, type in p->cellType).  Returns SQLITE_OK and fills
** *pVal, or SQLITE_ERROR after pFail() (shared string index out of
** range). */
static int cellValue(SheetParse *p, CellVal *pVal){
  const char *z = p->cellBuf.z ? p->cellBuf.z : "";
  sqlite3_int64 n = p->cellBuf.nUsed;

  memset(pVal, 0, sizeof(*pVal));
  pVal->z = z;
  pVal->n = n;
  switch( p->cellType ){
    case CELL_SHARED: {
      sqlite3_int64 idx;
      if( n==0 ){
        pVal->eType = SQLITE_NULL;
        return SQLITE_OK;
      }
      idx = strtoll(z, 0, 10);
      if( idx<0 || idx>=p->pSS->n ){
        pFail(p, "row %d, column %d: shared string index %lld out of range "
                 "(0..%d); is xl/sharedStrings.xml missing or incomplete?",
              p->iRow+1, p->iCol+1, idx, (sqlite3_int64)p->pSS->n - 1);
        return SQLITE_ERROR;
      }
      /* The shared string table outlives the INSERT statement (it is
      ** freed only after sqlite3_finalize), so SQLITE_STATIC is safe. */
      pVal->eType = SQLITE_TEXT;
      pVal->z = p->pSS->a[idx].z ? p->pSS->a[idx].z : "";
      pVal->n = p->pSS->a[idx].n;
      pVal->bStatic = 1;
      return SQLITE_OK;
    }
    case CELL_NUMBER: {
      if( n==0 ){
        pVal->eType = SQLITE_NULL;
        return SQLITE_OK;
      }
      if( looksInteger(z, n) ){
        char *zEnd = 0;
        sqlite3_int64 v;
        errno = 0;
        v = strtoll(z, &zEnd, 10);
        if( errno==0 && zEnd && zEnd==z+n ){
          pVal->eType = SQLITE_INTEGER;
          pVal->i = v;
        }else{
          pVal->eType = SQLITE_FLOAT;
          pVal->r = strtod(z, 0);
        }
      }else{
        char *zEnd = 0;
        double d = strtod(z, &zEnd);
        if( zEnd>z && zEnd==z+n && d>=-9.2e18 && d<=9.2e18 ){
          sqlite3_int64 v = (sqlite3_int64)d;
          if( (double)v==d ){                 /* e.g. "42.0" -> INTEGER 42 */
            pVal->eType = SQLITE_INTEGER;
            pVal->i = v;
          }else{
            pVal->eType = SQLITE_FLOAT;
            pVal->r = d;
          }
        }else if( zEnd>z && zEnd==z+n ){
          pVal->eType = SQLITE_FLOAT;
          pVal->r = d;
        }else{
          pVal->eType = SQLITE_TEXT;          /* keep verbatim */
        }
      }
      return SQLITE_OK;
    }
    case CELL_BOOL: {
      if( n==0 ){
        pVal->eType = SQLITE_NULL;
      }else if( z[0]=='1' || z[0]=='t' || z[0]=='T' ){
        pVal->eType = SQLITE_INTEGER;
        pVal->i = 1;
      }else if( z[0]=='0' || z[0]=='f' || z[0]=='F' ){
        pVal->eType = SQLITE_INTEGER;
        pVal->i = 0;
      }else{
        pVal->eType = SQLITE_TEXT;
      }
      return SQLITE_OK;
    }
    default: {                /* INLINE, STRING, ERROR, DATE -> TEXT */
      pVal->eType = SQLITE_TEXT;
      return SQLITE_OK;
    }
  }
}

/* Append the SQL literal representation of an inlined cell value to
** pSb.  Used on the "more columns than host parameters" insert path:
** the generated text is part of a row-specific INSERT statement built
** in endRow(). */
static void appendInlineLiteral(StrBuf *pSb, const InlineVal *pV){
  switch( pV->eType ){
    case SQLITE_INTEGER:
      strbufPrintf(pSb, "%lld", pV->i);
      break;
    case SQLITE_FLOAT:
      strbufPrintf(pSb, "%!.15g", pV->r);
      break;
    case SQLITE_TEXT: {
      if( pV->z==0 || pV->n==0 ){
        strbufAppend(pSb, "''", 2);
      }else if( memchr(pV->z, 0, (size_t)pV->n)==0 ){
        /* Plain text: single quotes doubled via SQLite's %q. */
        strbufPrintf(pSb, "'%.*q'", (int)pV->n, pV->z);
      }else{
        /* Text with embedded NUL bytes: hex-encode so no byte is lost. */
        sqlite3_int64 k;
        strbufAppend(pSb, "CAST(X'", -1);
        for(k=0; k<pV->n; k++){
          strbufPrintf(pSb, "%02x", (unsigned)(pV->z[k] & 0xff));
        }
        strbufAppend(pSb, "' AS TEXT)", -1);
      }
      break;
    }
    default:                    /* SQLITE_NULL and anything unknown */
      strbufAppend(pSb, "NULL", 4);
      break;
  }
}

/* Process the value of the cell that just ended (pass 2).  Fast path
** (one prepared INSERT with a host parameter per column): bind the value.
** Inline path (more columns than the connection's
** SQLITE_LIMIT_VARIABLE_NUMBER allows): stage a private copy of the
** value; endRow() then writes the whole row into a row-specific INSERT
** with every value inlined as an SQL literal, which keeps the generated
** statements free of host parameters. */
static void bindCell(SheetParse *p){
  CellVal v;
  int rc;

  if( p->iCol >= p->nCol ) return;          /* defensive; cannot happen */
  if( cellValue(p, &v)!=SQLITE_OK ) return;
  if( p->nInline>0 ){
    /* Stage the value for the inline insert path (see endRow). */
    InlineVal *pV = &p->aInline[p->iCol];
    char *zCopy = 0;
    if( v.eType==SQLITE_TEXT && v.n>0 ){
      zCopy = dupText(v.z, v.n);
      if( zCopy==0 ){
        pFail(p, "out of memory");
        return;
      }
    }
    sqlite3_free(pV->z);
    pV->eType = v.eType;
    pV->i = v.i;
    pV->r = v.r;
    pV->z = zCopy;
    pV->n = zCopy ? v.n : 0;
    p->nCellThisRow++;
    return;
  }
  switch( v.eType ){
    case SQLITE_NULL:
      rc = sqlite3_bind_null(p->pIns, p->iCol+1);
      break;
    case SQLITE_INTEGER:
      rc = sqlite3_bind_int64(p->pIns, p->iCol+1, v.i);
      break;
    case SQLITE_FLOAT:
      rc = sqlite3_bind_double(p->pIns, p->iCol+1, v.r);
      break;
    default:
      rc = sqlite3_bind_text64(p->pIns, p->iCol+1, v.z, (sqlite3_uint64)v.n,
                               v.bStatic ? SQLITE_STATIC : SQLITE_TRANSIENT,
                               SQLITE_UTF8);
      break;
  }
  if( rc!=SQLITE_OK ){
    pFail(p, "out of memory binding cell values");
    return;
  }
  p->aBound[p->iCol] = 1;
  p->nCellThisRow++;
}

/* Finish the current cell: update the column scan state (pass 1) or bind
** its value into the INSERT statement (pass 2). */
static void endCell(SheetParse *p){
  if( p->failed || !p->hasValue ) return;
  if( p->iCol<0 ) return;                   /* defensive; cannot happen */
  if( p->mode==2 ){
    /* Pass 2: cells beyond the imported column count are dropped
    ** instead of failing the import (the connection's
    ** SQLITE_LIMIT_COLUMN may be smaller than the sheet is wide). */
    if( p->iCol >= p->nCol ) return;
  }else if( p->iCol >= XLSX_SCAN_MAX_COLUMNS ){
    /* Scan guard: ignore cells of absurdly wide sheets. */
    return;
  }
  if( p->iCol+1 > p->nMaxCol ) p->nMaxCol = p->iCol+1;
  if( p->mode==1 ){
    const char *zV = 0;
    sqlite3_int64 nV = 0;
    if( p->headerRow<0 ) p->headerRow = p->iRow;
    if( p->iRow!=p->headerRow ) return;
    if( headerValue(p, &zV, &nV)!=SQLITE_OK ) return;
    if( hdrEnsure(p, p->iCol+1) ) return;
    sqlite3_free(p->aHdr[p->iCol].z);
    p->aHdr[p->iCol].z = dupText(zV, nV);
    p->aHdr[p->iCol].n = nV;
    if( p->aHdr[p->iCol].z==0 ){
      pFail(p, "out of memory");
      return;
    }
  }else{
    bindCell(p);
  }
}

/* Reset the per-row staging data: the bound flags and the owned copies
** of the staged cell values (only used on the inline insert path, where
** one value is staged per table column). */
static void clearRowData(SheetParse *p){
  int i;
  int nStaged = p->nInline>0 ? p->nCol : 0;
  if( p->aBound ) memset(p->aBound, 0, (size_t)p->nBind);
  if( p->aInline ){
    for(i=0; i<nStaged; i++){
      sqlite3_free(p->aInline[i].z);
      p->aInline[i].z = 0;
      p->aInline[i].eType = SQLITE_NULL;
      p->aInline[i].n = 0;
    }
  }
  p->nCellThisRow = 0;
}

/* Finish the current row (pass 2: fill unbound columns with NULL and
** execute the INSERT; fully empty rows and the header row are skipped).
** Normally a single prepared INSERT with one host parameter per column
** is used.  When the sheet is wider than the connection's
** SQLITE_LIMIT_VARIABLE_NUMBER allows, a row-specific INSERT is built
** instead, with every value inlined as an SQL literal so that the
** generated statements need no host parameters at all. */
static void endRow(SheetParse *p){
  int i, rc;
  if( p->failed || p->mode!=2 || p->inSkipRow ) return;
  if( p->nCellThisRow==0 ){
    clearRowData(p);
    return;
  }
  if( p->nInline==0 ){
    /* Fast path: the prepared INSERT has a parameter per column. */
    for(i=0; i<p->nBind; i++){
      if( !p->aBound[i] ){
        rc = sqlite3_bind_null(p->pIns, i+1);
        if( rc!=SQLITE_OK ){
          pFail(p, "out of memory");
          clearRowData(p);
          return;
        }
      }
    }
    rc = sqlite3_step(p->pIns);
    if( rc!=SQLITE_DONE ){
      pFail(p, "row %d: INSERT failed: %s", p->iRow+1, sqlite3_errmsg(p->db));
      sqlite3_reset(p->pIns);
      clearRowData(p);
      return;
    }
    sqlite3_reset(p->pIns);
    clearRowData(p);
    p->nRowIns++;
  }else{
    /* Inline path: build, prepare, run and free a row-specific INSERT
    ** with the staged values inlined as SQL literals. */
    StrBuf *pSb = &p->rowSql;
    sqlite3_stmt *pStmt = 0;
    strbufReset(pSb);
    strbufPrintf(pSb, "INSERT INTO \"%w\" VALUES(", p->zTabIns);
    for(i=0; i<p->nCol; i++){
      if( i ) strbufAppend(pSb, ",", 1);
      appendInlineLiteral(pSb, &p->aInline[i]);
    }
    strbufAppend(pSb, ")", 1);
    if( pSb->oom ){
      pFail(p, "out of memory");
      clearRowData(p);
      return;
    }
    rc = sqlite3_prepare_v2(p->db, pSb->z, -1, &pStmt, 0);
    if( rc==SQLITE_OK ){
      rc = sqlite3_step(pStmt);
      if( rc==SQLITE_DONE ) p->nRowIns++;
    }
    if( rc!=SQLITE_OK && rc!=SQLITE_DONE ){
      pFail(p, "row %d: INSERT failed: %s", p->iRow+1, sqlite3_errmsg(p->db));
    }
    sqlite3_finalize(pStmt);
    clearRowData(p);
  }
}

static void sheetStart(void *pUser, const XML_Char *zElem, const XML_Char **azAtts){
  SheetParse *p = (SheetParse*)pUser;
  const char *zR;

  if( p->failed ) return;
  if( strcmp(zElem, "c")==0 ){
    const char *zT;
    if( !p->inSheetData || p->inSkipRow ) return;
    p->inCell = 1;
    p->hasValue = 0;
    p->capture = CAP_NONE;
    p->cellType = CELL_NUMBER;
    zT = xmlAttr(azAtts, "t");
    if( zT ) p->cellType = typeFromT(zT);
    strbufReset(&p->cellBuf);
    zR = xmlAttr(azAtts, "r");
    if( zR ){
      int iCol = colFromRef(zR);
      if( iCol<0 ){
        pFail(p, "row %d: malformed cell reference \"%s\"", p->iRow+1, zR);
        return;
      }
      p->iCol = iCol;
      p->nextCol = iCol+1;
    }else{
      p->iCol = p->nextCol++;
    }
  }else if( strcmp(zElem, "row")==0 ){
    if( !p->inSheetData ) return;
    p->nextCol = 0;
    p->nCellThisRow = 0;
    zR = xmlAttr(azAtts, "r");
    if( zR && zR[0]>='1' && zR[0]<='9' ){
      char *zEnd = 0;
      long r = strtol(zR, &zEnd, 10);
      if( zEnd && *zEnd==0 && r>=1 && r<=100000000 ){
        p->iRow = (int)(r-1);
        p->nextRow = (int)r;
      }else{
        p->iRow = p->nextRow++;
      }
    }else{
      p->iRow = p->nextRow++;
    }
    /* Pass 2: the header row supplied the column names and is not data. */
    p->inSkipRow = (p->mode==2 && p->iRow==p->skipRow);
    if( p->inSkipRow ) return;
  }else if( strcmp(zElem, "v")==0 ){
    if( p->inCell ){
      p->capture = CAP_V;
      strbufReset(&p->vBuf);
    }
  }else if( strcmp(zElem, "is")==0 ){
    if( p->inCell ) p->cellType = CELL_INLINE;
  }else if( strcmp(zElem, "t")==0 ){
    if( p->inCell && p->cellType==CELL_INLINE ){
      p->capture = CAP_T;
      strbufReset(&p->vBuf);
    }
  }else if( strcmp(zElem, "sheetData")==0 ){
    p->inSheetData = 1;
  }
}

static void sheetChar(void *pUser, const XML_Char *z, int len){
  SheetParse *p = (SheetParse*)pUser;
  if( p->failed ) return;
  if( p->capture!=CAP_NONE ) strbufAppend(&p->vBuf, z, len);
}

static void sheetEnd(void *pUser, const XML_Char *zElem){
  SheetParse *p = (SheetParse*)pUser;
  if( p->failed ) return;
  if( strcmp(zElem, "v")==0 ){
    if( p->inCell && p->capture==CAP_V ){
      p->capture = CAP_NONE;
      p->hasValue = 1;
      strbufReset(&p->cellBuf);
      strbufAppend(&p->cellBuf, p->vBuf.z ? p->vBuf.z : "", p->vBuf.nUsed);
      if( p->cellBuf.oom ) pFail(p, "out of memory");
    }
  }else if( strcmp(zElem, "t")==0 ){
    if( p->inCell && p->capture==CAP_T ){
      p->capture = CAP_NONE;
      p->hasValue = 1;
      strbufAppend(&p->cellBuf, p->vBuf.z ? p->vBuf.z : "", p->vBuf.nUsed);
      if( p->cellBuf.oom ) pFail(p, "out of memory");
    }
  }else if( strcmp(zElem, "c")==0 ){
    if( p->inCell ){
      p->inCell = 0;
      p->capture = CAP_NONE;
      endCell(p);
    }
  }else if( strcmp(zElem, "row")==0 ){
    if( p->inSheetData ) endRow(p);
  }else if( strcmp(zElem, "sheetData")==0 ){
    p->inSheetData = 0;
  }
}

/* Parse the worksheet XML blob with the current mode/settings. */
static int parseSheet(SheetParse *p, Blob *pXml){
  StrBuf err;
  int rc;
  strbufInit(&err);
  rc = xmlParseBuffer(pXml, sheetStart, sheetEnd, sheetChar, p, &err);
  if( rc!=SQLITE_OK && p->zErrMsg==0 ){
    p->zErrMsg = sqlite3_mprintf("XML error in worksheet part: %s",
                                 err.z ? err.z : "out of memory");
  }
  strbufFree(&err);
  return rc;
}

/*
** Build the final column name list for a table with nCol columns.
** Header texts come from pass 1; missing names are generated and
** duplicates are numbered.  Returns a malloc'd array of nCol malloc'd
** strings, or NULL on out-of-memory.
*/
static char **buildColumns(SheetParse *p, int nCol){
  char **azCol = (char**)sqlite3_malloc64(sizeof(char*)*(sqlite3_uint64)nCol);
  int i;
  if( azCol==0 ) return 0;
  memset(azCol, 0, sizeof(char*)*(size_t)nCol);
  for(i=0; i<nCol; i++){
    char *zBase = 0;
    if( i<p->nHdrAlloc && p->aHdr[i].z ){
      zBase = sanitizeIdent(p->aHdr[i].z, p->aHdr[i].n);
    }
    if( zBase==0 ){
      char zL[8];
      colLetter(i, zL);
      zBase = dupText(zL, -1);
      if( zBase==0 ) goto oom;
      azCol[i] = zBase;
      continue;
    }
    if( !nameInList(azCol, i, zBase) ){
      azCol[i] = zBase;
    }else{
      int k;
      for(k=2; ; k++){
        char *zTry = sqlite3_mprintf("%s_%d", zBase, k);
        if( zTry==0 ){
          sqlite3_free(zBase);
          goto oom;
        }
        if( !nameInList(azCol, i, zTry) ){
          azCol[i] = zTry;
          sqlite3_free(zBase);
          break;
        }
        sqlite3_free(zTry);
      }
    }
  }
  return azCol;

oom:
  for(i=0; i<nCol; i++) sqlite3_free(azCol[i]);
  sqlite3_free(azCol);
  return 0;
}

/*
** Import one worksheet into a new table.  Two expat passes over the
** worksheet XML: pass 1 determines the column count and captures the
** header texts, pass 2 binds and inserts the data rows.
*/
static int importOneSheet(
  sqlite3 *db,
  const char *zFile,
  SheetInfo *pSheet,
  int sheetno,
  char ***pazCreated,        /* table names created during this import */
  int *pnCreated,
  int *pnCreatedAlloc,
  char **pzErr
){
  Blob xml;
  SharedStrings ss;
  SheetParse sp;
  char **azCol = 0;
  char *zTab = 0, *zSql = 0;
  sqlite3_stmt *pIns = 0;
  StrBuf sb;
  int nCol = 0;
  int nHdrRow = -1;
  int nOrigCol = 0;
  int maxCol, maxVar, nBind, nInline;
  int curCol, curVar, curSql;
  int restoreSql = -1;
  int rc, i;

  memset(&xml, 0, sizeof(xml));
  memset(&ss, 0, sizeof(ss));
  sheetParseInit(&sp);
  strbufInit(&sb);

  rc = readZipEntry(db, zFile, pSheet->zPath, &xml, &i, pzErr);
  if( rc ) goto done;
  if( !i || xml.n==0 ){
    *pzErr = sqlite3_mprintf(
        "xlsx_import: sheet \"%s\": worksheet part \"%s\" is missing or empty",
        pSheet->zName ? pSheet->zName : "", pSheet->zPath);
    rc = SQLITE_ERROR;
    goto done;
  }

  rc = loadSharedStrings(db, zFile, &ss, pzErr);
  if( rc ) goto done;

  /* ---------------- pass 1: scan columns and headers ---------------- */
  sp.pSS = &ss;
  sp.mode = 1;
  rc = parseSheet(&sp, &xml);
  if( rc==SQLITE_OK && sp.failed ) rc = SQLITE_ERROR;
  if( rc!=SQLITE_OK ){
    transferSheetError(&sp, pSheet->zName, pzErr);
    goto done;
  }
  nOrigCol = sp.nMaxCol ? sp.nMaxCol : 1;  /* empty sheet -> 1 column */

  /* ------------------------------------------------------------------
  ** Make room for the sheet's width.  The connection's run-time limits
  ** are raised to the sheet's column count -- SQLite silently truncates
  ** the request to the build's hard upper bound (SQLITE_MAX_COLUMN,
  ** SQLITE_MAX_VARIABLE_NUMBER) -- and then the effective values are
  ** queried with a -1 argument, so the import can adapt to whatever the
  ** connection actually allows instead of failing: as many leading
  ** columns as SQLITE_LIMIT_COLUMN permits are imported, the first
  ** SQLITE_LIMIT_VARIABLE_NUMBER of them bound as host parameters and
  ** any remaining columns inlined as SQL literals.  Both raised limits
  ** are deliberately left in effect: with the old values restored, even
  ** a plain "SELECT *" on the imported table would fail with "too many
  ** columns in result set".  (The SQL-statement-length limit, which is
  ** only needed while rows are being inserted, is restored at the end.)
  ** ------------------------------------------------------------------ */
  curCol = sqlite3_limit(db, SQLITE_LIMIT_COLUMN, -1);
  if( curCol < nOrigCol ){
    /* Deliberately not restored at the end (see the block comment). */
    sqlite3_limit(db, SQLITE_LIMIT_COLUMN, nOrigCol);
  }
  maxCol = sqlite3_limit(db, SQLITE_LIMIT_COLUMN, -1);
  if( maxCol<1 ){
    *pzErr = sqlite3_mprintf(
        "xlsx_import: sheet \"%s\": the database connection allows 0 "
        "columns (SQLITE_LIMIT_COLUMN is 0)",
        pSheet->zName ? pSheet->zName : "");
    rc = SQLITE_ERROR;
    goto done;
  }
  nCol = nOrigCol < maxCol ? nOrigCol : maxCol;
  if( nCol < nOrigCol ){
    /* More columns in the sheet than the connection can hold: import
    ** the leading nCol columns and warn instead of failing. */
    fprintf(stderr,
            "xlsx_import: sheet \"%s\": importing only the first %d of %d "
            "columns; the effective SQLITE_LIMIT_COLUMN of this connection "
            "is %d (rebuild SQLite with a larger SQLITE_MAX_COLUMN to "
            "raise it further)\n",
            pSheet->zName ? pSheet->zName : "", nCol, nOrigCol, maxCol);
    fflush(stderr);
  }

  curVar = sqlite3_limit(db, SQLITE_LIMIT_VARIABLE_NUMBER, -1);
  if( curVar < nCol ){
    /* Deliberately not restored at the end (see the block comment). */
    sqlite3_limit(db, SQLITE_LIMIT_VARIABLE_NUMBER, nCol);
  }
  maxVar = sqlite3_limit(db, SQLITE_LIMIT_VARIABLE_NUMBER, -1);
  nBind = nCol < maxVar ? nCol : maxVar;
  nInline = nCol - nBind;
  /* The inline insert path writes the whole row as SQL literals: make
  ** sure the connection allows at least the default SQL statement
  ** length (restored at the end of the import). */
  curSql = sqlite3_limit(db, SQLITE_LIMIT_SQL_LENGTH, -1);
  if( curSql < 10000000 ){
    sqlite3_limit(db, SQLITE_LIMIT_SQL_LENGTH, 10000000);
    restoreSql = curSql;
  }

  azCol = buildColumns(&sp, nCol);
  if( azCol==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto done;
  }

  /* ---------------- determine the table name ---------------- */
  zTab = sanitizeIdent(pSheet->zName, -1);
  if( zTab==0 ) zTab = sqlite3_mprintf("Sheet%d", sheetno);
  if( zTab==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto done;
  }
  {
    int nSfx = 1;
    while( nameInList(*pazCreated, *pnCreated, zTab) ){
      char *zTry = sqlite3_mprintf("%s_%d", zTab, ++nSfx);
      if( zTry==0 ){
        *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
        rc = SQLITE_NOMEM;
        goto done;
      }
      sqlite3_free(zTab);
      zTab = zTry;
    }
  }
  if( *pnCreated==*pnCreatedAlloc ){
    int nNew = *pnCreatedAlloc ? *pnCreatedAlloc*2 : 8;
    char **azNew = (char**)sqlite3_realloc64(
        *pazCreated, sizeof(char*)*(sqlite3_uint64)nNew);
    if( azNew==0 ){
      *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
      rc = SQLITE_NOMEM;
      goto done;
    }
    *pazCreated = azNew;
    *pnCreatedAlloc = nNew;
  }
  (*pazCreated)[(*pnCreated)++] = dupText(zTab, -1);
  if( (*pazCreated)[*pnCreated-1]==0 ){
    (*pnCreated)--;
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto done;
  }

  /* ---------------- create the table ---------------- */

  zSql = sqlite3_mprintf("DROP TABLE IF EXISTS \"%w\";", zTab);
  if( zSql==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto done;
  }
  rc = sqlite3_exec(db, zSql, 0, 0, 0);
  if( rc!=SQLITE_OK ){
    *pzErr = sqlite3_mprintf("xlsx_import: DROP TABLE \"%s\" failed: %s",
                             zTab, sqlite3_errmsg(db));
    goto done;
  }
  sqlite3_free(zSql);
  zSql = 0;

  strbufPrintf(&sb, "CREATE TABLE \"%w\"(", zTab);
  for(i=0; i<nCol; i++){
    char *zC = sqlite3_mprintf("\"%w\"", azCol[i]);
    if( zC==0 ){
      strbufFree(&sb);
      *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
      rc = SQLITE_NOMEM;
      goto done;
    }
    if( i ) strbufAppend(&sb, ",", 1);
    strbufAppend(&sb, zC, -1);
    sqlite3_free(zC);
  }
  strbufAppend(&sb, ")", 1);
  if( sb.oom ){
    strbufFree(&sb);
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto done;
  }
  rc = sqlite3_exec(db, sb.z, 0, 0, 0);
  strbufFree(&sb);
  if( rc!=SQLITE_OK ){
    *pzErr = sqlite3_mprintf("xlsx_import: CREATE TABLE \"%s\" failed: %s",
                             zTab, sqlite3_errmsg(db));
    goto done;
  }

  if( nInline==0 ){
    /* Fast path: one prepared INSERT with a host parameter per column. */
    strbufPrintf(&sb, "INSERT INTO \"%w\" VALUES(", zTab);
    for(i=0; i<nCol; i++){
      strbufPrintf(&sb, "%s?%d", i ? "," : "", i+1);
    }
    strbufAppend(&sb, ")", 1);
    if( sb.oom ){
      strbufFree(&sb);
      *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
      rc = SQLITE_NOMEM;
      goto done;
    }
    rc = sqlite3_prepare_v2(db, sb.z, -1, &pIns, 0);
    strbufFree(&sb);
    if( rc!=SQLITE_OK ){
      *pzErr = sqlite3_mprintf(
          "xlsx_import: preparing INSERT for table \"%s\" failed: %s",
          zTab, sqlite3_errmsg(db));
      goto done;
    }
  }
  /* When nInline>0 there are more columns than host parameters; pass 2
  ** then writes each row with a row-specific INSERT whose values are
  ** all inlined as SQL literals (see endRow). */

  /* ---------------- pass 2: bind and insert rows ---------------- */
  nHdrRow = sp.headerRow;
  sheetParseClear(&sp);
  sp.mode = 2;
  sp.skipRow = nHdrRow;        /* do not import the header row as data */
  sp.db = db;
  sp.pIns = pIns;
  sp.nCol = nCol;
  sp.nBind = nBind;
  sp.nInline = nInline;
  sp.zTabIns = zTab;
  sp.aBound = (char*)sqlite3_malloc64((sqlite3_uint64)(nBind>0 ? nBind : 1));
  if( sp.aBound==0 ){
    *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto done;
  }
  memset(sp.aBound, 0, (size_t)nBind);
  if( nInline>0 ){
    /* Inline insert path: one staged value per table column. */
    sp.aInline = (InlineVal*)sqlite3_malloc64(
        sizeof(InlineVal)*(sqlite3_uint64)nCol);
    if( sp.aInline==0 ){
      *pzErr = sqlite3_mprintf("xlsx_import: out of memory");
      rc = SQLITE_NOMEM;
      goto done;
    }
    memset(sp.aInline, 0, sizeof(InlineVal)*(size_t)nCol);
  }
  rc = parseSheet(&sp, &xml);
  if( rc==SQLITE_OK && sp.failed ) rc = SQLITE_ERROR;
  if( rc!=SQLITE_OK ){
    transferSheetError(&sp, pSheet->zName, pzErr);
    goto done;
  }

done:
  /* Restore the SQL-statement-length limit; the column and
  ** host-parameter limits stay raised (see the block comment above). */
  if( restoreSql>=0 ) sqlite3_limit(db, SQLITE_LIMIT_SQL_LENGTH, restoreSql);
  sqlite3_finalize(pIns);
  sheetParseFree(&sp);
  strbufFree(&sb);
  if( azCol ){
    for(i=0; i<nCol; i++) sqlite3_free(azCol[i]);
    sqlite3_free(azCol);
  }
  sqlite3_free(zTab);
  sqlite3_free(zSql);
  freeSharedStrings(&ss);
  sqlite3_free(xml.z);
  return rc;
}

/*
** ---------------------------------------------------------------------
** SQL function implementations.
** ---------------------------------------------------------------------
*/
static void versionFunc(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  UNUSED_PARAMETER(argc);
  UNUSED_PARAMETER(argv);
  sqlite3_result_text(ctx, XLSXIMPORT_VERSION, -1, SQLITE_STATIC);
}

/*
** xlsx_import(FILENAME [, SHEET-SELECTOR...])
*/
static void importFunc(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  sqlite3 *db = sqlite3_context_db_handle(ctx);
  const char *zFile;
  SheetList sl;
  unsigned char *aSel = 0;
  char **azCreated = 0;
  int nCreated = 0, nCreatedAlloc = 0;
  char *zErr = 0;
  int rc, i, nImported = 0;

  memset(&sl, 0, sizeof(sl));
  if( argc<1 || sqlite3_value_type(argv[0])==SQLITE_NULL ){
    sqlite3_result_error(ctx,
        "xlsx_import: first argument must be the XLSX filename", -1);
    return;
  }
  zFile = (const char*)sqlite3_value_text(argv[0]);
  if( zFile==0 ){
    sqlite3_result_error_nomem(ctx);
    return;
  }

  rc = loadSheetList(db, zFile, &sl, &zErr);
  if( rc!=SQLITE_OK ) goto error_out;

  aSel = (unsigned char*)sqlite3_malloc64((sqlite3_uint64)sl.n + 1);
  if( aSel==0 ){
    zErr = sqlite3_mprintf("xlsx_import: out of memory");
    rc = SQLITE_NOMEM;
    goto error_out;
  }
  memset(aSel, 0, (size_t)sl.n);

  if( argc==1 ){
    /* No selectors: import every sheet. */
    memset(aSel, 1, (size_t)sl.n);
  }else{
    for(i=1; i<argc; i++){
      int vt = sqlite3_value_type(argv[i]);
      if( vt==SQLITE_NULL ){
        zErr = sqlite3_mprintf(
            "xlsx_import: sheet selector %d is NULL: sheet name or sheet "
            "number expected", i);
        rc = SQLITE_ERROR;
        goto error_out;
      }
      if( vt==SQLITE_INTEGER || vt==SQLITE_FLOAT ){
        sqlite3_int64 no;
        if( vt==SQLITE_FLOAT ){
          double f = sqlite3_value_double(argv[i]);
          if( f>9.0e18 || f<-9.0e18 || f!=(double)(sqlite3_int64)f ){
            zErr = sqlite3_mprintf(
                "xlsx_import: sheet number must be an integer, not %!.15g", f);
            rc = SQLITE_ERROR;
            goto error_out;
          }
        }
        no = sqlite3_value_int64(argv[i]);
        if( no<1 || no>sl.n ){
          zErr = sqlite3_mprintf(
              "xlsx_import: sheet number %lld is out of range: \"%s\" "
              "contains %d sheet%s (must be between 1 and %d)",
              no, zFile, sl.n, sl.n==1 ? "" : "s", sl.n);
          rc = SQLITE_ERROR;
          goto error_out;
        }
        aSel[no-1] = 1;
      }else{
        const char *zS = (const char*)sqlite3_value_text(argv[i]);
        int j, matched = 0;
        if( zS==0 ){
          zErr = sqlite3_mprintf("xlsx_import: out of memory");
          rc = SQLITE_NOMEM;
          goto error_out;
        }
        /* An exact sheet name takes precedence over a number. */
        for(j=0; j<sl.n; j++){
          if( sl.a[j].zName && strcmp(sl.a[j].zName, zS)==0 ){
            aSel[j] = 1;
            matched = 1;
            break;
          }
        }
        if( !matched && allDigits(zS) ){
          sqlite3_int64 no = strtoll(zS, 0, 10);
          if( no<1 || no>sl.n ){
            zErr = sqlite3_mprintf(
                "xlsx_import: sheet number %lld is out of range: \"%s\" "
                "contains %d sheet%s (must be between 1 and %d)",
                no, zFile, sl.n, sl.n==1 ? "" : "s", sl.n);
            rc = SQLITE_ERROR;
            goto error_out;
          }
          aSel[no-1] = 1;
          matched = 1;
        }
        if( !matched ){
          StrBuf avail;
          int j2;
          strbufInit(&avail);
          for(j2=0; j2<sl.n; j2++){
            strbufPrintf(&avail, "%s\"%s\"", j2 ? ", " : "",
                         sl.a[j2].zName ? sl.a[j2].zName : "?");
          }
          zErr = sqlite3_mprintf(
              "xlsx_import: no such sheet: \"%s\".  Sheets of \"%s\": %s",
              zS, zFile, avail.z ? avail.z : "");
          strbufFree(&avail);
          rc = SQLITE_ERROR;
          goto error_out;
        }
      }
    }
  }

  rc = sqlite3_exec(db, "SAVEPOINT xlsx_import", 0, 0, 0);
  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf("xlsx_import: SAVEPOINT failed: %s",
                           sqlite3_errmsg(db));
    goto error_out;
  }
  for(i=0; i<sl.n; i++){
    if( !aSel[i] ) continue;
    rc = importOneSheet(db, zFile, &sl.a[i], i+1,
                        &azCreated, &nCreated, &nCreatedAlloc, &zErr);
    if( rc!=SQLITE_OK ) goto rollback_out;
    nImported++;
  }
  rc = sqlite3_exec(db, "RELEASE SAVEPOINT xlsx_import", 0, 0, 0);
  if( rc!=SQLITE_OK ){
    zErr = sqlite3_mprintf("xlsx_import: RELEASE failed: %s",
                           sqlite3_errmsg(db));
    goto rollback_out;
  }

  freeSheetList(&sl);
  sqlite3_free(aSel);
  for(i=0; i<nCreated; i++) sqlite3_free(azCreated[i]);
  sqlite3_free(azCreated);
  sqlite3_result_int(ctx, nImported);
  return;

rollback_out:
  sqlite3_exec(db, "ROLLBACK TO SAVEPOINT xlsx_import", 0, 0, 0);
  sqlite3_exec(db, "RELEASE SAVEPOINT xlsx_import", 0, 0, 0);
error_out:
  if( zErr==0 ){
    zErr = sqlite3_mprintf("xlsx_import: %s", sqlite3_errmsg(db));
  }
  sqlite3_result_error(ctx, zErr, -1);
  sqlite3_free(zErr);
  freeSheetList(&sl);
  sqlite3_free(aSel);
  for(i=0; i<nCreated; i++) sqlite3_free(azCreated[i]);
  sqlite3_free(azCreated);
}

/*
** ---------------------------------------------------------------------
** xlsx_import_sheetnames: eponymous-only virtual table.
** ---------------------------------------------------------------------
*/
typedef struct SheetnamesVtab SheetnamesVtab;
struct SheetnamesVtab {
  sqlite3_vtab base;
  sqlite3 *db;
};

typedef struct SheetnamesCur SheetnamesCur;
struct SheetnamesCur {
  sqlite3_vtab_cursor base;
  SheetList sl;            /* sheet list owned by the cursor */
  char *zFile;
  int i;
};

static int snConnect(
  sqlite3 *db,
  void *pAux,
  int argc,
  const char *const*argv,
  sqlite3_vtab **ppVtab,
  char **pzErr
){
  SheetnamesVtab *pNew;
  int rc;
  UNUSED_PARAMETER(pAux);
  UNUSED_PARAMETER(argc);
  UNUSED_PARAMETER(argv);
  UNUSED_PARAMETER(pzErr);
  pNew = (SheetnamesVtab*)sqlite3_malloc(sizeof(SheetnamesVtab));
  if( pNew==0 ) return SQLITE_NOMEM;
  memset(pNew, 0, sizeof(SheetnamesVtab));
  pNew->db = db;
  rc = sqlite3_declare_vtab(db,
      "CREATE TABLE x(\"filename\" TEXT HIDDEN,"
      "\"sheet_num\" INTEGER, \"sheet_name\" TEXT)");
  if( rc!=SQLITE_OK ){
    sqlite3_free(pNew);
    return rc;
  }
  *ppVtab = &pNew->base;
  return SQLITE_OK;
}

static int snDisconnect(sqlite3_vtab *pVtab){
  sqlite3_free(pVtab);
  return SQLITE_OK;
}

static int snBestIndex(sqlite3_vtab *pVtab, sqlite3_index_info *pIdxInfo){
  int i;
  UNUSED_PARAMETER(pVtab);
  for(i=0; i<pIdxInfo->nConstraint; i++){
    const struct sqlite3_index_constraint *pC = &pIdxInfo->aConstraint[i];
    if( pC->usable && pC->iColumn==0 && pC->op==SQLITE_INDEX_CONSTRAINT_EQ ){
      pIdxInfo->aConstraintUsage[i].argvIndex = 1;
      pIdxInfo->aConstraintUsage[i].omit = 1;
      pIdxInfo->estimatedCost = 1000.0;
      pIdxInfo->estimatedRows = 25;
      return SQLITE_OK;
    }
  }
  /* The filename argument is mandatory. */
  return SQLITE_CONSTRAINT;
}

static int snOpen(sqlite3_vtab *pVtab, sqlite3_vtab_cursor **ppCur){
  SheetnamesCur *pCur;
  UNUSED_PARAMETER(pVtab);
  pCur = (SheetnamesCur*)sqlite3_malloc(sizeof(SheetnamesCur));
  if( pCur==0 ) return SQLITE_NOMEM;
  memset(pCur, 0, sizeof(SheetnamesCur));
  *ppCur = &pCur->base;
  return SQLITE_OK;
}

static int snClose(sqlite3_vtab_cursor *cur){
  SheetnamesCur *pCur = (SheetnamesCur*)cur;
  freeSheetList(&pCur->sl);
  sqlite3_free(pCur->zFile);
  sqlite3_free(pCur);
  return SQLITE_OK;
}

static int snFilter(
  sqlite3_vtab_cursor *cur,
  int idxNum,
  const char *idxStr,
  int argc,
  sqlite3_value **argv
){
  SheetnamesCur *pCur = (SheetnamesCur*)cur;
  SheetnamesVtab *pVtab = (SheetnamesVtab*)cur->pVtab;
  char *zErr = 0;
  int rc;
  UNUSED_PARAMETER(idxNum);
  UNUSED_PARAMETER(idxStr);

  freeSheetList(&pCur->sl);
  sqlite3_free(pCur->zFile);
  pCur->zFile = 0;
  pCur->i = 0;

  if( argc<1 || sqlite3_value_type(argv[0])==SQLITE_NULL ){
    pVtab->base.zErrMsg = sqlite3_mprintf(
        "xlsx_import_sheetnames: an XLSX filename argument is required");
    return SQLITE_ERROR;
  }
  pCur->zFile = dupText((const char*)sqlite3_value_text(argv[0]), -1);
  if( pCur->zFile==0 ){
    pVtab->base.zErrMsg = sqlite3_mprintf("out of memory");
    return SQLITE_NOMEM;
  }
  rc = loadSheetList(pVtab->db, pCur->zFile, &pCur->sl, &zErr);
  if( rc!=SQLITE_OK ){
    pVtab->base.zErrMsg = zErr ? zErr : sqlite3_mprintf("out of memory");
    freeSheetList(&pCur->sl);
    return rc;
  }
  return SQLITE_OK;
}

static int snNext(sqlite3_vtab_cursor *cur){
  ((SheetnamesCur*)cur)->i++;
  return SQLITE_OK;
}

static int snEof(sqlite3_vtab_cursor *cur){
  SheetnamesCur *pCur = (SheetnamesCur*)cur;
  return pCur->i >= pCur->sl.n;
}

static int snColumn(sqlite3_vtab_cursor *cur, sqlite3_context *pCtx, int iCol){
  SheetnamesCur *pCur = (SheetnamesCur*)cur;
  switch( iCol ){
    case 0:
      sqlite3_result_text(pCtx, pCur->zFile ? pCur->zFile : "", -1,
                          SQLITE_TRANSIENT);
      break;
    case 1:
      sqlite3_result_int64(pCtx, pCur->i + 1);
      break;
    default: {
      const char *z = "";
      if( pCur->i < pCur->sl.n && pCur->sl.a[pCur->i].zName ){
        z = pCur->sl.a[pCur->i].zName;
      }
      sqlite3_result_text(pCtx, z, -1, SQLITE_TRANSIENT);
      break;
    }
  }
  return SQLITE_OK;
}

static int snRowid(sqlite3_vtab_cursor *cur, sqlite3_int64 *pRowid){
  *pRowid = ((SheetnamesCur*)cur)->i + 1;
  return SQLITE_OK;
}

static sqlite3_module sheetnamesModule = {
  .iVersion    = 1,
  /* xCreate is 0, which makes this an eponymous-only virtual table. */
  .xConnect    = snConnect,
  .xBestIndex  = snBestIndex,
  .xDisconnect = snDisconnect,
  .xOpen       = snOpen,
  .xClose      = snClose,
  .xFilter     = snFilter,
  .xNext       = snNext,
  .xEof        = snEof,
  .xColumn     = snColumn,
  .xRowid      = snRowid
};

/*
** ---------------------------------------------------------------------
** Extension entry point.
** ---------------------------------------------------------------------
*/
#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsximport_init(
  sqlite3 *db,
  char **pzErrMsg,
  const sqlite3_api_routines *pApi
){
  int rc;
  SQLITE_EXTENSION_INIT2(pApi);
  UNUSED_PARAMETER(pzErrMsg);
  rc = sqlite3_create_function(db, "xlsx_import_version", 0,
          SQLITE_UTF8 | SQLITE_DIRECTONLY, 0, versionFunc, 0, 0);
  if( rc==SQLITE_OK ){
    rc = sqlite3_create_function(db, "xlsx_import", -1,
            SQLITE_UTF8 | SQLITE_DIRECTONLY, 0, importFunc, 0, 0);
  }
  if( rc==SQLITE_OK ){
    rc = sqlite3_create_module(db, "xlsx_import_sheetnames",
                               &sheetnamesModule, 0);
  }
  return rc;
}

#ifndef XLSXIMPORT_OMIT_GENERIC_ENTRY
/* Generic entry point alias, so the extension can also be loaded from a
** shared library whose name differs from "xlsximport". */
int sqlite3_extension_init(
  sqlite3 *db,
  char **pzErrMsg,
  const sqlite3_api_routines *pApi
){
  return sqlite3_xlsximport_init(db, pzErrMsg, pApi);
}
#endif
