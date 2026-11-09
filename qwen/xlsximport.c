/*
** xlsximport.c
**
** SQLite extension to import XLSX files.
**
** Requires:
**   - SQLite zipfile extension available in the same database connection.
**   - Expat XML parser.
**
** SQL functions:
**
**   xlsx_import(filename, [sheet_name_or_number, ...])
**     Creates one table for each selected sheet.
**     Table name = sheet name.
**     Column names = first row values.
**     Returns number of sheets imported.
**
**   xlsx_import_sheetnames(filename)
**     Table-valued function returning sheet_num, sheet_name.
**
**   xlsx_import_version()
**     Returns version string.
**
** Build example on Unix-like systems:
**
**   gcc -g -fPIC -shared xlsximport.c -o xlsximport.so \
**       -I/path/to/sqlite3 -lexpat
**
** Load in SQLite:
**
**   SELECT load_extension('zipfile');
**   SELECT load_extension('xlsximport');
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <expat.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>
#include <limits.h>

#ifndef SQLITE_DETERMINISTIC
#define SQLITE_DETERMINISTIC 0x000000800
#endif

#ifndef SQLITE_INDEX_CONSTRAINT_FUNCTION
#define SQLITE_INDEX_CONSTRAINT_FUNCTION 128
#endif

/* --------------------------------------------------------------------- */
/* Small utility helpers                                                  */
/* --------------------------------------------------------------------- */

typedef struct XlsxBuffer {
  char *z;
  int n;
  int nAlloc;
} XlsxBuffer;

static void xlsxBufInit(XlsxBuffer *p){
  memset(p, 0, sizeof(*p));
}

static void xlsxBufFree(XlsxBuffer *p){
  sqlite3_free(p->z);
  p->z = 0;
  p->n = 0;
  p->nAlloc = 0;
}

static void xlsxBufReset(XlsxBuffer *p){
  p->n = 0;
  if( p->z ) p->z[0] = '\0';
}

static int xlsxBufAppend(XlsxBuffer *p, const char *z, int n){
  if( z==0 ) return SQLITE_OK;
  if( n<0 ) n = (int)strlen(z);

  if( p->n + n + 1 > p->nAlloc ){
    int nNew = p->nAlloc ? p->nAlloc : 128;
    char *zNew;
    while( nNew < p->n + n + 1 ) nNew *= 2;
    zNew = (char*)sqlite3_realloc64(p->z, nNew);
    if( zNew==0 ) return SQLITE_NOMEM;
    p->z = zNew;
    p->nAlloc = nNew;
  }

  memcpy(p->z + p->n, z, n);
  p->n += n;
  p->z[p->n] = '\0';
  return SQLITE_OK;
}

static char *xlsxBufTake(XlsxBuffer *p){
  char *z;
  if( p->z==0 ){
    z = (char*)sqlite3_malloc64(1);
    if( z ) z[0] = '\0';
    return z;
  }
  z = (char*)sqlite3_realloc64(p->z, p->n + 1);
  if( z==0 ) z = p->z;
  p->z = 0;
  p->n = 0;
  p->nAlloc = 0;
  return z;
}

static char *xlsxStrdup(const char *z){
  int n;
  char *zOut;
  if( z==0 ) return 0;
  n = (int)strlen(z);
  zOut = (char*)sqlite3_malloc64((sqlite3_int64)n + 1);
  if( zOut ) memcpy(zOut, z, n + 1);
  return zOut;
}

static const char *xlsxLocalName(const char *z){
  const char *p;
  if( z==0 ) return "";
  p = strrchr(z, ':');
  return p ? p + 1 : z;
}

static const char *xlsxAttrValue(const XML_Char **attrs, const char *zLocal){
  int i;
  if( attrs==0 || zLocal==0 ) return 0;
  for(i=0; attrs[i]; i+=2){
    if( strcmp(xlsxLocalName((const char*)attrs[i]), zLocal)==0 ){
      return (const char*)attrs[i+1];
    }
  }
  return 0;
}

/* Convert an Excel cell reference such as "B12" to a 0-based column index. */
static int xlsxRefToCol(const char *z){
  int col = 0;
  if( z==0 ) return -1;

  while( *z && isalpha((unsigned char)*z) ){
    int d = toupper((unsigned char)*z) - 'A';
    if( d < 0 || d >= 26 ) return -1;
    col = col * 26 + d + 1;
    if( col > 100000000 ) return -1;  /* sanity limit */
    z++;
  }

  if( col==0 ) return -1;
  return col - 1;
}

static int xlsxTextIsPositiveInt(const char *z, int *pVal){
  int i = 0;
  int v = 0;
  int has = 0;

  if( z==0 ) return 0;

  while( isspace((unsigned char)z[i]) ) i++;

  while( z[i] >= '0' && z[i] <= '9' ){
    int d = z[i++] - '0';
    if( v > (INT_MAX - d) / 10 ) return 0;
    v = v * 10 + d;
    has = 1;
  }

  while( isspace((unsigned char)z[i]) ) i++;

  if( !has || z[i] != '\0' || v <= 0 ) return 0;
  *pVal = v;
  return 1;
}

static int xlsxIsInt(const char *z, sqlite3_int64 *pVal){
  int i = 0;
  int neg = 0;
  int has = 0;
  sqlite3_uint64 v = 0;
  const sqlite3_uint64 maxPos = 9223372036854775807ULL;
  const sqlite3_uint64 maxNeg = 9223372036854775808ULL;
  sqlite3_uint64 limit;

  if( z==0 ) return 0;

  while( isspace((unsigned char)z[i]) ) i++;

  if( z[i]=='-' || z[i]=='+' ){
    neg = (z[i]=='-');
    i++;
  }

  limit = neg ? maxNeg : maxPos;

  if( !isdigit((unsigned char)z[i]) ) return 0;

  while( isdigit((unsigned char)z[i]) ){
    int d = z[i++] - '0';
    if( v > (limit - d) / 10 ) return 0;
    v = v * 10 + d;
    has = 1;
  }

  while( isspace((unsigned char)z[i]) ) i++;

  if( !has || z[i] != '\0' ) return 0;

  if( neg ){
    if( v == maxNeg ){
      *pVal = -9223372036854775807LL - 1;
    }else{
      *pVal = -(sqlite3_int64)v;
    }
  }else{
    *pVal = (sqlite3_int64)v;
  }

  return 1;
}

static int xlsxIsDouble(const char *z, double *pVal){
  char *end = 0;
  double v;

  if( z==0 || z[0]=='\0' ) return 0;

  v = strtod(z, &end);
  if( end==z ) return 0;

  while( end && isspace((unsigned char)*end) ) end++;
  if( end==0 || *end != '\0' ) return 0;

  *pVal = v;
  return 1;
}

/* Quote an SQL identifier:  some"name -> "some""name" */
static char *xlsxQuoteIdent(const char *z){
  int i, j, n;
  char *out;

  if( z==0 ) z = "";
  n = (int)strlen(z);

  out = (char*)sqlite3_malloc64((sqlite3_int64)n * 2 + 3);
  if( out==0 ) return 0;

  out[0] = '"';
  j = 1;

  for(i=0; i<n; i++){
    char c = z[i];
    if( c == '"' ) out[j++] = '"';
    out[j++] = c;
  }

  out[j++] = '"';
  out[j] = '\0';
  return out;
}

/* --------------------------------------------------------------------- */
/* Dynamic arrays                                                         */
/* --------------------------------------------------------------------- */

typedef struct XlsxStrArray {
  char **az;
  int n;
  int nAlloc;
} XlsxStrArray;

static int xlsxStrArrayAppend(XlsxStrArray *p, char *z){
  if( p->n >= p->nAlloc ){
    int nNew = p->nAlloc ? p->nAlloc * 2 : 8;
    char **azNew = (char**)sqlite3_realloc64(p->az, (sqlite3_int64)nNew * sizeof(char*));
    if( azNew==0 ) return SQLITE_NOMEM;
    p->az = azNew;
    p->nAlloc = nNew;
  }
  p->az[p->n++] = z;
  return SQLITE_OK;
}

static void xlsxStrArrayFree(XlsxStrArray *p){
  int i;
  if( p->az ){
    for(i=0; i<p->n; i++) sqlite3_free(p->az[i]);
    sqlite3_free(p->az);
  }
  memset(p, 0, sizeof(*p));
}

typedef struct XlsxSheet {
  char *zName;
  char *zRId;
} XlsxSheet;

typedef struct XlsxSheetArray {
  XlsxSheet *a;
  int n;
  int nAlloc;
} XlsxSheetArray;

static int xlsxSheetArrayAppend(
  XlsxSheetArray *p,
  const char *zName,
  const char *zRId
){
  char *zN;
  char *zI;

  if( p->n >= p->nAlloc ){
    int nNew = p->nAlloc ? p->nAlloc * 2 : 8;
    XlsxSheet *aNew = (XlsxSheet*)sqlite3_realloc64(p->a, (sqlite3_int64)nNew * sizeof(XlsxSheet));
    if( aNew==0 ) return SQLITE_NOMEM;
    p->a = aNew;
    p->nAlloc = nNew;
  }

  zN = xlsxStrdup(zName ? zName : "");
  if( zN==0 ) return SQLITE_NOMEM;

  zI = xlsxStrdup(zRId ? zRId : "");
  if( zI==0 ){
    sqlite3_free(zN);
    return SQLITE_NOMEM;
  }

  p->a[p->n].zName = zN;
  p->a[p->n].zRId = zI;
  p->n++;
  return SQLITE_OK;
}

static void xlsxSheetArrayFree(XlsxSheetArray *p){
  int i;
  if( p->a ){
    for(i=0; i<p->n; i++){
      sqlite3_free(p->a[i].zName);
      sqlite3_free(p->a[i].zRId);
    }
    sqlite3_free(p->a);
  }
  memset(p, 0, sizeof(*p));
}

typedef struct XlsxRel {
  char *zId;
  char *zTarget;
} XlsxRel;

typedef struct XlsxRelArray {
  XlsxRel *a;
  int n;
  int nAlloc;
} XlsxRelArray;

static int xlsxRelArrayAppend(
  XlsxRelArray *p,
  const char *zId,
  const char *zTarget
){
  char *zI;
  char *zT;

  if( p->n >= p->nAlloc ){
    int nNew = p->nAlloc ? p->nAlloc * 2 : 8;
    XlsxRel *aNew = (XlsxRel*)sqlite3_realloc64(p->a, (sqlite3_int64)nNew * sizeof(XlsxRel));
    if( aNew==0 ) return SQLITE_NOMEM;
    p->a = aNew;
    p->nAlloc = nNew;
  }

  zI = xlsxStrdup(zId ? zId : "");
  if( zI==0 ) return SQLITE_NOMEM;

  zT = xlsxStrdup(zTarget ? zTarget : "");
  if( zT==0 ){
    sqlite3_free(zI);
    return SQLITE_NOMEM;
  }

  p->a[p->n].zId = zI;
  p->a[p->n].zTarget = zT;
  p->n++;
  return SQLITE_OK;
}

static void xlsxRelArrayFree(XlsxRelArray *p){
  int i;
  if( p->a ){
    for(i=0; i<p->n; i++){
      sqlite3_free(p->a[i].zId);
      sqlite3_free(p->a[i].zTarget);
    }
    sqlite3_free(p->a);
  }
  memset(p, 0, sizeof(*p));
}

static const char *xlsxRelFind(XlsxRelArray *pRels, const char *zId){
  int i;
  if( pRels==0 || zId==0 || zId[0]=='\0' ) return 0;
  for(i=0; i<pRels->n; i++){
    if( strcmp(pRels->a[i].zId, zId)==0 ){
      return pRels->a[i].zTarget;
    }
  }
  return 0;
}

/* --------------------------------------------------------------------- */
/* Path helpers for resolving worksheet targets                           */
/* --------------------------------------------------------------------- */

static char *xlsxNormalizePath(const char *z){
  char *zCopy;
  char *p;
  char **azStack = 0;
  int nStack = 0;
  int nAlloc = 0;
  int i;
  XlsxBuffer out;

  if( z==0 ) return 0;

  zCopy = xlsxStrdup(z);
  if( zCopy==0 ) return 0;

  p = zCopy;
  xlsxBufInit(&out);

  while( *p ){
    int len;
    char *seg;

    while( *p == '/' ) p++;
    if( *p == '\0' ) break;

    len = 0;
    while( p[len] && p[len] != '/' ) len++;

    if( len==1 && p[0]=='.' ){
      /* skip */
    }else if( len==2 && p[0]=='.' && p[1]=='.' ){
      if( nStack > 0 ) nStack--;
    }else{
      if( nStack >= nAlloc ){
        int nNew = nAlloc ? nAlloc * 2 : 8;
        char **azNew = (char**)sqlite3_realloc64(azStack, (sqlite3_int64)nNew * sizeof(char*));
        if( azNew==0 ){
          for(i=0; i<nStack; i++) sqlite3_free(azStack[i]);
          sqlite3_free(azStack);
          sqlite3_free(zCopy);
          xlsxBufFree(&out);
          return 0;
        }
        azStack = azNew;
        nAlloc = nNew;
      }

      seg = (char*)sqlite3_malloc64((sqlite3_int64)len + 1);
      if( seg==0 ){
        for(i=0; i<nStack; i++) sqlite3_free(azStack[i]);
        sqlite3_free(azStack);
        sqlite3_free(zCopy);
        xlsxBufFree(&out);
        return 0;
      }

      memcpy(seg, p, len);
      seg[len] = '\0';
      azStack[nStack++] = seg;
    }

    p += len;
  }

  for(i=0; i<nStack; i++){
    if( i > 0 ) xlsxBufAppend(&out, "/", 1);
    xlsxBufAppend(&out, azStack[i], -1);
  }

  for(i=0; i<nStack; i++) sqlite3_free(azStack[i]);
  sqlite3_free(azStack);
  sqlite3_free(zCopy);

  if( out.z==0 ) return xlsxStrdup("");
  return out.z;
}

static char *xlsxResolveTarget(const char *zBase, const char *zTarget){
  XlsxBuffer b;
  char *zPath;

  if( zTarget==0 ) return 0;

  xlsxBufInit(&b);

  if( zTarget[0] == '/' ){
    xlsxBufAppend(&b, zTarget + 1, -1);
  }else{
    if( zBase && zBase[0] ){
      xlsxBufAppend(&b, zBase, -1);
      xlsxBufAppend(&b, "/", 1);
    }
    xlsxBufAppend(&b, zTarget, -1);
  }

  zPath = xlsxNormalizePath(b.z ? b.z : "");
  xlsxBufFree(&b);
  return zPath;
}

/* --------------------------------------------------------------------- */
/* Read one file from the XLSX ZIP archive using SQLite zipfile()         */
/* --------------------------------------------------------------------- */

static int xlsxReadZipEntry(
  sqlite3 *db,
  const char *zZipFile,
  const char *zEntryName,
  unsigned char **ppData,
  int *pnData,
  char **pzErr
){
  sqlite3_stmt *pStmt = 0;
  char *zSql;
  int rc;

  *ppData = 0;
  *pnData = 0;

  zSql = sqlite3_mprintf(
    "SELECT data FROM zipfile(?1) WHERE name = ?2 LIMIT 1"
  );
  if( zSql==0 ) return SQLITE_NOMEM;

  rc = sqlite3_prepare_v2(db, zSql, -1, &pStmt, 0);
  sqlite3_free(zSql);

  if( rc != SQLITE_OK ){
    if( pzErr && *pzErr==0 ){
      *pzErr = sqlite3_mprintf(
        "xlsx_import: cannot query zipfile() for \"%s\": %s",
        zZipFile, sqlite3_errmsg(db)
      );
    }
    return rc;
  }

  sqlite3_bind_text(pStmt, 1, zZipFile, -1, SQLITE_TRANSIENT);
  sqlite3_bind_text(pStmt, 2, zEntryName, -1, SQLITE_TRANSIENT);

  rc = sqlite3_step(pStmt);

  if( rc == SQLITE_ROW ){
    if( sqlite3_column_type(pStmt, 0) == SQLITE_NULL ){
      *ppData = 0;
      *pnData = 0;
      rc = SQLITE_OK;
    }else{
      const unsigned char *pIn = (const unsigned char*)sqlite3_column_blob(pStmt, 0);
      int n = sqlite3_column_bytes(pStmt, 0);
      unsigned char *pOut = (unsigned char*)sqlite3_malloc64((sqlite3_int64)n + 1);

      if( pOut==0 ){
        rc = SQLITE_NOMEM;
      }else{
        if( n > 0 && pIn ) memcpy(pOut, pIn, n);
        pOut[n] = '\0';
        *ppData = pOut;
        *pnData = n;
        rc = SQLITE_OK;
      }
    }
  }else if( rc == SQLITE_DONE ){
    *ppData = 0;
    *pnData = 0;
    rc = SQLITE_OK;
  }else{
    if( pzErr && *pzErr==0 ){
      *pzErr = sqlite3_mprintf(
        "xlsx_import: error reading \"%s\" from \"%s\": %s",
        zEntryName, zZipFile, sqlite3_errmsg(db)
      );
    }
  }

  sqlite3_finalize(pStmt);
  return rc;
}

/* --------------------------------------------------------------------- */
/* Generic Expat parse helper                                             */
/* --------------------------------------------------------------------- */

static void XMLCALL xlsxNoopEndElement(void *userData, const XML_Char *name){
  (void)userData;
  (void)name;
}

static int xlsxParseXml(
  const unsigned char *zData,
  int nData,
  void *pCtx,
  XML_StartElementHandler xStart,
  XML_EndElementHandler xEnd,
  XML_CharacterDataHandler xChar,
  char **pzErr
){
  XML_Parser parser;
  int st;

  parser = XML_ParserCreate(NULL);
  if( parser==0 ){
    if( pzErr && *pzErr==0 ){
      *pzErr = xlsxStrdup("xlsx_import: out of memory creating XML parser");
    }
    return SQLITE_NOMEM;
  }

  XML_SetUserData(parser, pCtx);
  XML_SetElementHandler(parser, xStart, xEnd ? xEnd : xlsxNoopEndElement);
  if( xChar ){
    XML_SetCharacterDataHandler(parser, xChar);
  }

  st = XML_Parse(parser, (const char*)zData, nData, 1);

  if( st == XML_STATUS_ERROR ){
    if( pzErr && *pzErr==0 ){
      *pzErr = sqlite3_mprintf(
        "xlsx_import: XML parse error at line %d: %s",
        (int)XML_GetCurrentLineNumber(parser),
        XML_ErrorString(XML_GetErrorCode(parser))
      );
    }
    XML_ParserFree(parser);
    return SQLITE_ERROR;
  }

  XML_ParserFree(parser);
  return SQLITE_OK;
}

/* --------------------------------------------------------------------- */
/* Parse xl/workbook.xml                                                  */
/* --------------------------------------------------------------------- */

static void XMLCALL xlsxWorkbookStartElement(
  void *userData,
  const XML_Char *name,
  const XML_Char **attrs
){
  XlsxSheetArray *pSheets = (XlsxSheetArray*)userData;

  if( strcmp(xlsxLocalName((const char*)name), "sheet")==0 ){
    const char *zName = xlsxAttrValue(attrs, "name");
    const char *zId   = xlsxAttrValue(attrs, "id");
    xlsxSheetArrayAppend(pSheets, zName, zId);
  }
}

static int xlsxLoadWorkbook(
  sqlite3 *db,
  const char *zFile,
  XlsxSheetArray *pSheets,
  char **pzErr
){
  unsigned char *pData = 0;
  int nData = 0;
  int rc;

  rc = xlsxReadZipEntry(db, zFile, "xl/workbook.xml", &pData, &nData, pzErr);
  if( rc != SQLITE_OK ) return rc;

  if( pData==0 ){
    if( pzErr && *pzErr==0 ){
      *pzErr = sqlite3_mprintf(
        "xlsx_import: cannot find xl/workbook.xml in \"%s\"", zFile
      );
    }
    return SQLITE_ERROR;
  }

  rc = xlsxParseXml(
    pData, nData,
    (void*)pSheets,
    xlsxWorkbookStartElement,
    0,
    0,
    pzErr
  );

  sqlite3_free(pData);
  return rc;
}

/* --------------------------------------------------------------------- */
/* Parse xl/_rels/workbook.xml.rels                                       */
/* --------------------------------------------------------------------- */

static void XMLCALL xlsxRelsStartElement(
  void *userData,
  const XML_Char *name,
  const XML_Char **attrs
){
  XlsxRelArray *pRels = (XlsxRelArray*)userData;

  if( strcmp(xlsxLocalName((const char*)name), "Relationship")==0 ){
    const char *zId     = xlsxAttrValue(attrs, "Id");
    const char *zTarget = xlsxAttrValue(attrs, "Target");
    xlsxRelArrayAppend(pRels, zId, zTarget);
  }
}

static int xlsxLoadRels(
  sqlite3 *db,
  const char *zFile,
  XlsxRelArray *pRels,
  char **pzErr
){
  unsigned char *pData = 0;
  int nData = 0;
  int rc;

  rc = xlsxReadZipEntry(
    db, zFile,
    "xl/_rels/workbook.xml.rels",
    &pData, &nData, pzErr
  );
  if( rc != SQLITE_OK ) return rc;

  if( pData==0 ){
    /* Missing relationships file is tolerated; fallback guessing is used. */
    return SQLITE_OK;
  }

  rc = xlsxParseXml(
    pData, nData,
    (void*)pRels,
    xlsxRelsStartElement,
    0,
    0,
    pzErr
  );

  sqlite3_free(pData);
  return rc;
}

/* --------------------------------------------------------------------- */
/* Parse xl/sharedStrings.xml                                             */
/* --------------------------------------------------------------------- */

typedef struct XlsxSSParse {
  XlsxStrArray *pStrs;
  XlsxBuffer buf;
  int inSI;
  int inT;
  int inPh;
} XlsxSSParse;

static void XMLCALL xlsxSSStartElement(
  void *userData,
  const XML_Char *name,
  const XML_Char **attrs
){
  XlsxSSParse *p = (XlsxSSParse*)userData;
  const char *ln = xlsxLocalName((const char*)name);

  (void)attrs;

  if( strcmp(ln, "si")==0 ){
    p->inSI = 1;
    p->inT = 0;
    p->inPh = 0;
    xlsxBufReset(&p->buf);
  }else if( p->inSI && strcmp(ln, "t")==0 ){
    p->inT = 1;
  }else if( strcmp(ln, "rPh")==0 ){
    p->inPh++;
  }
}

static void XMLCALL xlsxSSEndElement(
  void *userData,
  const XML_Char *name
){
  XlsxSSParse *p = (XlsxSSParse*)userData;
  const char *ln = xlsxLocalName((const char*)name);

  if( strcmp(ln, "si")==0 ){
    char *z = xlsxBufTake(&p->buf);
    if( z ){
      if( xlsxStrArrayAppend(p->pStrs, z) != SQLITE_OK ){
        sqlite3_free(z);
      }
    }
    p->inSI = 0;
    p->inT = 0;
    p->inPh = 0;
  }else if( strcmp(ln, "t")==0 ){
    p->inT = 0;
  }else if( strcmp(ln, "rPh")==0 ){
    if( p->inPh > 0 ) p->inPh--;
  }
}

static void XMLCALL xlsxSSCharacterData(
  void *userData,
  const XML_Char *s,
  int len
){
  XlsxSSParse *p = (XlsxSSParse*)userData;

  if( p->inSI && p->inT && p->inPh==0 ){
    xlsxBufAppend(&p->buf, (const char*)s, len);
  }
}

static int xlsxLoadSharedStrings(
  sqlite3 *db,
  const char *zFile,
  XlsxStrArray *pStrs,
  char **pzErr
){
  unsigned char *pData = 0;
  int nData = 0;
  int rc;
  XlsxSSParse ctx;

  rc = xlsxReadZipEntry(
    db, zFile,
    "xl/sharedStrings.xml",
    &pData, &nData, pzErr
  );
  if( rc != SQLITE_OK ) return rc;

  if( pData==0 ){
    /* No shared strings is fine. */
    return SQLITE_OK;
  }

  memset(&ctx, 0, sizeof(ctx));
  ctx.pStrs = pStrs;
  xlsxBufInit(&ctx.buf);

  rc = xlsxParseXml(
    pData, nData,
    (void*)&ctx,
    xlsxSSStartElement,
    xlsxSSEndElement,
    xlsxSSCharacterData,
    pzErr
  );

  xlsxBufFree(&ctx.buf);
  sqlite3_free(pData);
  return rc;
}

/* --------------------------------------------------------------------- */
/* Sheet parsing / table creation / row insertion                         */
/* --------------------------------------------------------------------- */

#define XLTYPE_NUMERIC 0
#define XLTYPE_SHARED  1
#define XLTYPE_STRING  2
#define XLTYPE_INLINE  3
#define XLTYPE_BOOL    4
#define XLTYPE_ERROR   5
#define XLTYPE_DATE    6

typedef struct XlsxRowCell {
  int iCol;
  int type;
  int hasValue;
  char *zText;
} XlsxRowCell;

typedef struct XlsxSheetParse {
  XML_Parser parser;

  sqlite3 *db;
  const char *zTable;
  XlsxStrArray *pSS;

  int hdrDone;
  int sawRow;

  int nCols;
  int *aiCols;          /* Original Excel column indexes */
  char **azColNames;    /* SQL column names */
  sqlite3_stmt *pInsert;

  int inRow;
  int nextCol;

  int hasContent;
  int inV;
  int inIS;
  int inT;
  int inPh;

  int cellType;
  int curCol;

  XlsxBuffer buf;

  XlsxRowCell *aCells;
  int nCells;
  int nCellsAlloc;

  char *zErrMsg;
} XlsxSheetParse;

static void xlsxSheetAbort(XlsxSheetParse *p, const char *zMsg){
  if( p->zErrMsg==0 ){
    p->zErrMsg = xlsxStrdup(zMsg ? zMsg : "xlsx_import: error");
  }
  if( p->parser ){
    XML_StopParser(p->parser, XML_FALSE);
  }
}

static int xlsxParseCellType(const char *z){
  if( z==0 ) return XLTYPE_NUMERIC;
  if( strcmp(z, "s")==0 ) return XLTYPE_SHARED;
  if( strcmp(z, "str")==0 ) return XLTYPE_STRING;
  if( strcmp(z, "inlineStr")==0 ) return XLTYPE_INLINE;
  if( strcmp(z, "b")==0 ) return XLTYPE_BOOL;
  if( strcmp(z, "e")==0 ) return XLTYPE_ERROR;
  if( strcmp(z, "d")==0 ) return XLTYPE_DATE;
  return XLTYPE_NUMERIC;
}

static void xlsxRowClear(XlsxSheetParse *p){
  int i;
  if( p->aCells ){
    for(i=0; i<p->nCells; i++){
      sqlite3_free(p->aCells[i].zText);
    }
  }
  p->nCells = 0;
}

static int xlsxAddCell(
  XlsxSheetParse *p,
  int iCol,
  int type,
  char *zText,
  int hasValue
){
  if( p->nCells >= p->nCellsAlloc ){
    int nNew = p->nCellsAlloc ? p->nCellsAlloc * 2 : 16;
    XlsxRowCell *aNew = (XlsxRowCell*)sqlite3_realloc64(
      p->aCells, (sqlite3_int64)nNew * sizeof(XlsxRowCell)
    );
    if( aNew==0 ){
      sqlite3_free(zText);
      return SQLITE_NOMEM;
    }
    p->aCells = aNew;
    p->nCellsAlloc = nNew;
  }

  p->aCells[p->nCells].iCol = iCol;
  p->aCells[p->nCells].type = type;
  p->aCells[p->nCells].hasValue = hasValue;
  p->aCells[p->nCells].zText = zText;
  p->nCells++;
  return SQLITE_OK;
}

static XlsxRowCell *xlsxFindCell(XlsxSheetParse *p, int iCol){
  int i;
  for(i=0; i<p->nCells; i++){
    if( p->aCells[i].iCol == iCol ) return &p->aCells[i];
  }
  return 0;
}

static int xlsxCreateTableForColumns(
  sqlite3 *db,
  const char *zTable,
  int nCols,
  char **azCols,
  sqlite3_stmt **ppInsert,
  char **pzErr
){
  char *zTableId = 0;
  char *zSql = 0;
  char *zErr = 0;
  int rc = SQLITE_OK;
  int i;
  XlsxBuffer create;
  XlsxBuffer insert;

  xlsxBufInit(&create);
  xlsxBufInit(&insert);

  if( ppInsert ) *ppInsert = 0;
  if( nCols < 1 ) nCols = 1;

  zTableId = xlsxQuoteIdent(zTable);
  if( zTableId==0 ){
    rc = SQLITE_NOMEM;
    goto cleanup;
  }

  /* Replace any existing table with the same name. */
  zSql = sqlite3_mprintf("DROP TABLE IF EXISTS %s", zTableId);
  if( zSql==0 ){
    rc = SQLITE_NOMEM;
    goto cleanup;
  }

  rc = sqlite3_exec(db, zSql, 0, 0, &zErr);
  sqlite3_free(zSql);
  zSql = 0;

  if( rc != SQLITE_OK ){
    if( pzErr ){
      if( *pzErr ) sqlite3_free(*pzErr);
      *pzErr = zErr ? zErr : xlsxStrdup(sqlite3_errmsg(db));
    }else{
      sqlite3_free(zErr);
    }
    goto cleanup;
  }

  /* CREATE TABLE ... */
  xlsxBufAppend(&create, "CREATE TABLE ", -1);
  xlsxBufAppend(&create, zTableId, -1);
  xlsxBufAppend(&create, "(", -1);

  for(i=0; i<nCols; i++){
    char *zCol;
    const char *zName = (azCols && azCols[i]) ? azCols[i] : "column";

    if( i > 0 ) xlsxBufAppend(&create, ", ", -1);

    zCol = xlsxQuoteIdent(zName);
    if( zCol==0 ){
      rc = SQLITE_NOMEM;
      goto cleanup;
    }

    xlsxBufAppend(&create, zCol, -1);
    sqlite3_free(zCol);
  }

  xlsxBufAppend(&create, ")", -1);

  if( create.z==0 ){
    rc = SQLITE_NOMEM;
    goto cleanup;
  }

  rc = sqlite3_exec(db, create.z, 0, 0, &zErr);
  if( rc != SQLITE_OK ){
    if( pzErr ){
      if( *pzErr ) sqlite3_free(*pzErr);
      *pzErr = zErr ? zErr : xlsxStrdup(sqlite3_errmsg(db));
    }else{
      sqlite3_free(zErr);
    }
    goto cleanup;
  }

  /* Optional INSERT statement for data rows. */
  if( ppInsert ){
    xlsxBufAppend(&insert, "INSERT INTO ", -1);
    xlsxBufAppend(&insert, zTableId, -1);
    xlsxBufAppend(&insert, "(", -1);

    for(i=0; i<nCols; i++){
      char *zCol;
      const char *zName = (azCols && azCols[i]) ? azCols[i] : "column";

      if( i > 0 ) xlsxBufAppend(&insert, ", ", -1);

      zCol = xlsxQuoteIdent(zName);
      if( zCol==0 ){
        rc = SQLITE_NOMEM;
        goto cleanup;
      }

      xlsxBufAppend(&insert, zCol, -1);
      sqlite3_free(zCol);
    }

    xlsxBufAppend(&insert, ") VALUES(", -1);

    for(i=0; i<nCols; i++){
      if( i > 0 ) xlsxBufAppend(&insert, ",", 1);
      xlsxBufAppend(&insert, "?", 1);
    }

    xlsxBufAppend(&insert, ")", -1);

    if( insert.z==0 ){
      rc = SQLITE_NOMEM;
      goto cleanup;
    }

    rc = sqlite3_prepare_v2(db, insert.z, -1, ppInsert, 0);
    if( rc != SQLITE_OK ){
      if( pzErr ){
        if( *pzErr ) sqlite3_free(*pzErr);
        *pzErr = xlsxStrdup(sqlite3_errmsg(db));
      }
      goto cleanup;
    }
  }

cleanup:
  sqlite3_free(zTableId);
  sqlite3_free(zSql);
  xlsxBufFree(&create);
  xlsxBufFree(&insert);
  return rc;
}

static int xlsxBuildHeader(XlsxSheetParse *p){
  int i, j;
  int nUniq = 0;
  int nLimit;
  int rc;
  int max;

  nLimit = sqlite3_limit(p->db, SQLITE_LIMIT_COLUMN, -1);
  if( nLimit < 1 ) nLimit = 2000;

  max = p->nCells > 0 ? p->nCells : 1;

  p->aiCols = (int*)sqlite3_malloc64((sqlite3_int64)max * sizeof(int));
  p->azColNames = (char**)sqlite3_malloc64((sqlite3_int64)max * sizeof(char*));

  if( p->aiCols==0 || p->azColNames==0 ){
    return SQLITE_NOMEM;
  }

  memset(p->azColNames, 0, (size_t)max * sizeof(char*));
  p->nCols = 0;

  /* Build unique column list from header row cells. */
  for(i=0; i<p->nCells; i++){
    XlsxRowCell *c = &p->aCells[i];
    int found = -1;

    for(j=0; j<nUniq; j++){
      if( p->aiCols[j] == c->iCol ){
        found = j;
        break;
      }
    }

    if( found >= 0 ){
      if( c->hasValue && c->zText && c->zText[0] ){
        sqlite3_free(p->azColNames[found]);
        p->azColNames[found] = xlsxStrdup(c->zText);
        if( p->azColNames[found]==0 ) return SQLITE_NOMEM;
      }
    }else{
      p->aiCols[nUniq] = c->iCol;

      if( c->hasValue && c->zText && c->zText[0] ){
        p->azColNames[nUniq] = xlsxStrdup(c->zText);
        if( p->azColNames[nUniq]==0 ) return SQLITE_NOMEM;
      }else{
        p->azColNames[nUniq] = 0;
      }

      nUniq++;
      p->nCols = nUniq;
    }
  }

  /* Sort by Excel column index. */
  for(i=1; i<nUniq; i++){
    int key = p->aiCols[i];
    char *zName = p->azColNames[i];

    for(j=i-1; j>=0 && p->aiCols[j] > key; j--){
      p->aiCols[j+1] = p->aiCols[j];
      p->azColNames[j+1] = p->azColNames[j];
    }

    p->aiCols[j+1] = key;
    p->azColNames[j+1] = zName;
  }

  /* Respect SQLite's column count limit. */
  if( nUniq > nLimit ){
    for(i=nLimit; i<nUniq; i++){
      sqlite3_free(p->azColNames[i]);
      p->azColNames[i] = 0;
    }
    nUniq = nLimit;
  }

  p->nCols = nUniq;

  /* If there were no usable header cells, create one generic column. */
  if( p->nCols == 0 ){
    p->aiCols[0] = 0;
    p->azColNames[0] = xlsxStrdup("column1");
    if( p->azColNames[0]==0 ) return SQLITE_NOMEM;
    p->nCols = 1;
  }else{
    /* Fill in missing names. */
    for(i=0; i<p->nCols; i++){
      if( p->azColNames[i]==0 || p->azColNames[i][0]=='\0' ){
        char z[32];
        sqlite3_snprintf(sizeof(z), z, "column%d", i + 1);
        sqlite3_free(p->azColNames[i]);
        p->azColNames[i] = xlsxStrdup(z);
        if( p->azColNames[i]==0 ) return SQLITE_NOMEM;
      }
    }

    /* Make column names unique. SQLite does not allow duplicate names. */
    for(i=0; i<p->nCols; i++){
      int dup = 1;
      int suffix = 2;
      char *zBase = xlsxStrdup(p->azColNames[i]);
      if( zBase==0 ) return SQLITE_NOMEM;

      while( dup ){
        dup = 0;
        for(j=0; j<i; j++){
          if( strcmp(p->azColNames[j], p->azColNames[i])==0 ){
            dup = 1;
            break;
          }
        }

        if( dup ){
          char *zNew = sqlite3_mprintf("%s_%d", zBase, suffix++);
          if( zNew==0 ){
            sqlite3_free(zBase);
            return SQLITE_NOMEM;
          }
          sqlite3_free(p->azColNames[i]);
          p->azColNames[i] = zNew;
        }
      }

      sqlite3_free(zBase);
    }
  }

  rc = xlsxCreateTableForColumns(
    p->db,
    p->zTable,
    p->nCols,
    p->azColNames,
    &p->pInsert,
    &p->zErrMsg
  );

  if( rc == SQLITE_OK ){
    p->hdrDone = 1;
  }

  return rc;
}

static int xlsxInsertRow(XlsxSheetParse *p){
  int i;
  int rc;

  if( p->pInsert==0 ) return SQLITE_OK;

  rc = sqlite3_reset(p->pInsert);
  if( rc != SQLITE_OK ) return rc;

  sqlite3_clear_bindings(p->pInsert);

  for(i=0; i<p->nCols; i++){
    XlsxRowCell *c = xlsxFindCell(p, p->aiCols[i]);

    if( c==0 || !c->hasValue || c->zText==0 || c->zText[0]=='\0' ){
      sqlite3_bind_null(p->pInsert, i + 1);
      continue;
    }

    switch( c->type ){
      case XLTYPE_BOOL: {
        int v = (c->zText[0]=='1' || sqlite3_stricmp(c->zText, "true")==0);
        sqlite3_bind_int(p->pInsert, i + 1, v);
        break;
      }

      case XLTYPE_NUMERIC: {
        sqlite3_int64 iVal;
        double rVal;

        if( xlsxIsInt(c->zText, &iVal) ){
          sqlite3_bind_int64(p->pInsert, i + 1, iVal);
        }else if( xlsxIsDouble(c->zText, &rVal) ){
          sqlite3_bind_double(p->pInsert, i + 1, rVal);
        }else{
          sqlite3_bind_text(p->pInsert, i + 1, c->zText, -1, SQLITE_TRANSIENT);
        }
        break;
      }

      default: {
        sqlite3_bind_text(p->pInsert, i + 1, c->zText, -1, SQLITE_TRANSIENT);
        break;
      }
    }
  }

  rc = sqlite3_step(p->pInsert);
  if( rc == SQLITE_DONE ) rc = SQLITE_OK;
  return rc;
}

static void XMLCALL xlsxSheetStartElement(
  void *userData,
  const XML_Char *name,
  const XML_Char **attrs
){
  XlsxSheetParse *p = (XlsxSheetParse*)userData;
  const char *ln = xlsxLocalName((const char*)name);

  if( p->zErrMsg ) return;

  if( strcmp(ln, "row")==0 ){
    xlsxRowClear(p);
    p->inRow = 1;
    p->nextCol = 0;
    p->inPh = 0;
    p->sawRow = 1;
    return;
  }

  if( !p->inRow ) return;

  if( strcmp(ln, "c")==0 ){
    const char *zRef = xlsxAttrValue(attrs, "r");
    const char *zType = xlsxAttrValue(attrs, "t");

    p->curCol = zRef ? xlsxRefToCol(zRef) : -1;
    if( p->curCol < 0 ) p->curCol = p->nextCol;

    p->nextCol = p->curCol + 1;
    p->cellType = xlsxParseCellType(zType);
    p->hasContent = 0;
    p->inV = 0;
    p->inIS = 0;
    p->inT = 0;
    p->inPh = 0;

    xlsxBufReset(&p->buf);
    return;
  }

  if( strcmp(ln, "v")==0 ){
    p->inV = 1;
    p->hasContent = 1;
    xlsxBufReset(&p->buf);
    return;
  }

  if( strcmp(ln, "is")==0 ){
    p->inIS = 1;
    p->hasContent = 1;
    xlsxBufReset(&p->buf);
    return;
  }

  if( strcmp(ln, "t")==0 ){
    if( p->inIS ) p->inT = 1;
    return;
  }

  if( strcmp(ln, "rPh")==0 ){
    p->inPh++;
    return;
  }
}

static void XMLCALL xlsxSheetCharacterData(
  void *userData,
  const XML_Char *s,
  int len
){
  XlsxSheetParse *p = (XlsxSheetParse*)userData;

  if( p->zErrMsg ) return;

  if( p->inV || (p->inIS && p->inT && p->inPh==0) ){
    if( xlsxBufAppend(&p->buf, (const char*)s, len) != SQLITE_OK ){
      xlsxSheetAbort(p, "xlsx_import: out of memory");
    }
  }
}

static void XMLCALL xlsxSheetEndElement(
  void *userData,
  const XML_Char *name
){
  XlsxSheetParse *p = (XlsxSheetParse*)userData;
  const char *ln = xlsxLocalName((const char*)name);
  int rc;

  if( p->zErrMsg ) return;

  if( strcmp(ln, "v")==0 ){
    p->inV = 0;
    return;
  }

  if( strcmp(ln, "t")==0 ){
    p->inT = 0;
    return;
  }

  if( strcmp(ln, "rPh")==0 ){
    if( p->inPh > 0 ) p->inPh--;
    return;
  }

  if( strcmp(ln, "is")==0 ){
    p->inIS = 0;
    return;
  }

  if( strcmp(ln, "c")==0 ){
    char *zText = xlsxBufTake(&p->buf);

    if( !p->hasContent ){
      sqlite3_free(zText);
      zText = 0;
    }

    if( zText && p->cellType == XLTYPE_SHARED ){
      int idx = atoi(zText);
      if( idx >= 0 && p->pSS && idx < p->pSS->n ){
        char *zShared = xlsxStrdup(p->pSS->az[idx]);
        sqlite3_free(zText);
        zText = zShared;
      }
    }

    if( xlsxAddCell(p, p->curCol, p->cellType, zText, p->hasContent) != SQLITE_OK ){
      xlsxSheetAbort(p, "xlsx_import: out of memory");
    }

    p->hasContent = 0;
    return;
  }

  if( strcmp(ln, "row")==0 ){
    if( !p->hdrDone ){
      rc = xlsxBuildHeader(p);
    }else{
      rc = xlsxInsertRow(p);
    }

    if( rc != SQLITE_OK ){
      if( p->zErrMsg==0 ){
        p->zErrMsg = xlsxStrdup(sqlite3_errmsg(p->db));
      }
      XML_StopParser(p->parser, XML_FALSE);
    }

    xlsxRowClear(p);
    p->inRow = 0;
    return;
  }
}

static int xlsxCreateSimpleTable(
  sqlite3 *db,
  const char *zTable,
  char **pzErr
){
  char *azCols[1];
  azCols[0] = (char*)"column1";
  return xlsxCreateTableForColumns(db, zTable, 1, azCols, 0, pzErr);
}

static int xlsxImportOneSheet(
  sqlite3 *db,
  const char *zFile,
  XlsxSheet *pSheet,
  int iSheet,
  XlsxRelArray *pRels,
  XlsxStrArray *pSS,
  char **pzErr
){
  int rc = SQLITE_OK;
  char *zPath = 0;
  char *zTable = 0;
  unsigned char *pData = 0;
  int nData = 0;
  XML_Parser parser = 0;
  XlsxSheetParse ctx;

  const char *zTarget = pSheet->zRId ? xlsxRelFind(pRels, pSheet->zRId) : 0;

  if( zTarget ){
    zPath = xlsxResolveTarget("xl", zTarget);
  }else{
    char *zGuess = sqlite3_mprintf("worksheets/sheet%d.xml", iSheet + 1);
    if( zGuess ){
      zPath = xlsxResolveTarget("xl", zGuess);
      sqlite3_free(zGuess);
    }
  }

  if( zPath==0 ) return SQLITE_NOMEM;

  rc = xlsxReadZipEntry(db, zFile, zPath, &pData, &nData, pzErr);
  if( rc != SQLITE_OK ){
    sqlite3_free(zPath);
    return rc;
  }

  if( pData==0 ){
    if( pzErr && *pzErr==0 ){
      *pzErr = sqlite3_mprintf(
        "xlsx_import: cannot find worksheet \"%s\" in \"%s\"",
        zPath, zFile
      );
    }
    sqlite3_free(zPath);
    return SQLITE_ERROR;
  }

  if( pSheet->zName && pSheet->zName[0] ){
    zTable = xlsxStrdup(pSheet->zName);
  }else{
    zTable = sqlite3_mprintf("sheet%d", iSheet + 1);
  }

  if( zTable==0 ){
    sqlite3_free(pData);
    sqlite3_free(zPath);
    return SQLITE_NOMEM;
  }

  memset(&ctx, 0, sizeof(ctx));
  ctx.db = db;
  ctx.zTable = zTable;
  ctx.pSS = pSS;
  xlsxBufInit(&ctx.buf);

  parser = XML_ParserCreate(NULL);
  if( parser==0 ){
    rc = SQLITE_NOMEM;
    goto import_cleanup;
  }

  ctx.parser = parser;

  XML_SetUserData(parser, &ctx);
  XML_SetElementHandler(parser, xlsxSheetStartElement, xlsxSheetEndElement);
  XML_SetCharacterDataHandler(parser, xlsxSheetCharacterData);

  if( XML_Parse(parser, (const char*)pData, nData, 1) == XML_STATUS_ERROR ){
    if( ctx.zErrMsg==0 ){
      if( pzErr && *pzErr==0 ){
        *pzErr = sqlite3_mprintf(
          "xlsx_import: XML parse error in \"%s\" at line %d: %s",
          zPath,
          (int)XML_GetCurrentLineNumber(parser),
          XML_ErrorString(XML_GetErrorCode(parser))
        );
      }
    }
    rc = SQLITE_ERROR;
  }else{
    if( ctx.zErrMsg ){
      rc = SQLITE_ERROR;
    }else if( !ctx.hdrDone ){
      /* No rows at all: create a default one-column table. */
      rc = xlsxCreateSimpleTable(db, zTable, pzErr);
    }
  }

import_cleanup:
  if( parser ) XML_ParserFree(parser);

  if( ctx.zErrMsg ){
    if( pzErr && *pzErr==0 ){
      *pzErr = xlsxStrdup(ctx.zErrMsg);
    }
    sqlite3_free(ctx.zErrMsg);
    ctx.zErrMsg = 0;
  }

  xlsxRowClear(&ctx);
  sqlite3_free(ctx.aCells);

  if( ctx.pInsert ) sqlite3_finalize(ctx.pInsert);

  if( ctx.azColNames ){
    int i;
    for(i=0; i<ctx.nCols; i++){
      sqlite3_free(ctx.azColNames[i]);
    }
    sqlite3_free(ctx.azColNames);
  }

  sqlite3_free(ctx.aiCols);
  xlsxBufFree(&ctx.buf);

  sqlite3_free(zTable);
  sqlite3_free(pData);
  sqlite3_free(zPath);

  return rc;
}

/* --------------------------------------------------------------------- */
/* SQL scalar function: xlsx_import()                                     */
/* --------------------------------------------------------------------- */

static void xlsxImportFunc(
  sqlite3_context *context,
  int argc,
  sqlite3_value **argv
){
  sqlite3 *db = sqlite3_context_db_handle(context);
  const char *zFile;

  XlsxSheetArray sheets = {0,0,0};
  XlsxRelArray rels = {0,0,0};
  XlsxStrArray ss = {0,0,0};

  unsigned char *aSel = 0;
  char *zErr = 0;

  int rc = SQLITE_OK;
  int nImported = 0;
  int i;

  if( argc < 1 ){
    sqlite3_result_error(
      context,
      "xlsx_import(): first argument must be an XLSX filename",
      -1
    );
    return;
  }

  zFile = (const char*)sqlite3_value_text(argv[0]);
  if( zFile==0 ){
    sqlite3_result_error(
      context,
      "xlsx_import(): filename must not be NULL",
      -1
    );
    return;
  }

  rc = xlsxLoadWorkbook(db, zFile, &sheets, &zErr);
  if( rc != SQLITE_OK ) goto error;

  if( sheets.n == 0 && argc == 1 ){
    sqlite3_result_int(context, 0);
    goto cleanup;
  }

  aSel = (unsigned char*)sqlite3_malloc64(sheets.n > 0 ? sheets.n : 1);
  if( aSel==0 ){
    rc = SQLITE_NOMEM;
    goto error;
  }
  memset(aSel, 0, sheets.n > 0 ? sheets.n : 1);

  if( argc == 1 ){
    for(i=0; i<sheets.n; i++) aSel[i] = 1;
  }else{
    int haveSel = 0;

    for(i=1; i<argc; i++){
      int eType = sqlite3_value_type(argv[i]);

      if( eType == SQLITE_NULL ) continue;

      if( eType == SQLITE_INTEGER || eType == SQLITE_FLOAT ){
        int iSheet = sqlite3_value_int(argv[i]);

        if( iSheet < 1 || iSheet > sheets.n ){
          if( zErr==0 ){
            zErr = sqlite3_mprintf(
              "xlsx_import(): sheet number %d out of range (1..%d)",
              iSheet, sheets.n
            );
          }
          rc = SQLITE_ERROR;
          goto error;
        }

        aSel[iSheet - 1] = 1;
        haveSel = 1;
      }else{
        const char *zName = (const char*)sqlite3_value_text(argv[i]);
        int j;
        int found = 0;

        if( zName ){
          /* Exact match first. */
          for(j=0; j<sheets.n; j++){
            if( strcmp(zName, sheets.a[j].zName)==0 ){
              aSel[j] = 1;
              found = 1;
            }
          }

          /* Then case-insensitive match. */
          if( !found ){
            for(j=0; j<sheets.n; j++){
              if( sqlite3_stricmp(zName, sheets.a[j].zName)==0 ){
                aSel[j] = 1;
                found = 1;
              }
            }
          }

          /* If still not found, allow numeric sheet number passed as text. */
          if( !found ){
            int iSheet;
            if( xlsxTextIsPositiveInt(zName, &iSheet) && iSheet <= sheets.n ){
              aSel[iSheet - 1] = 1;
              found = 1;
            }
          }
        }

        if( !found ){
          if( zErr==0 ){
            zErr = sqlite3_mprintf(
              "xlsx_import(): sheet \"%s\" not found",
              zName ? zName : ""
            );
          }
          rc = SQLITE_ERROR;
          goto error;
        }

        haveSel = 1;
      }
    }

    if( !haveSel ){
      sqlite3_result_int(context, 0);
      goto cleanup;
    }
  }

  rc = xlsxLoadRels(db, zFile, &rels, &zErr);
  if( rc != SQLITE_OK ) goto error;

  rc = xlsxLoadSharedStrings(db, zFile, &ss, &zErr);
  if( rc != SQLITE_OK ) goto error;

  rc = sqlite3_exec(db, "SAVEPOINT xlsximport;", 0, 0, &zErr);
  if( rc != SQLITE_OK ) goto error;

  for(i=0; i<sheets.n; i++){
    if( !aSel[i] ) continue;

    rc = xlsxImportOneSheet(
      db, zFile,
      &sheets.a[i], i,
      &rels, &ss,
      &zErr
    );

    if( rc != SQLITE_OK ) break;
    nImported++;
  }

  if( rc == SQLITE_OK ){
    rc = sqlite3_exec(db, "RELEASE xlsximport;", 0, 0, &zErr);
  }else{
    char *zErr2 = 0;
    sqlite3_exec(
      db,
      "ROLLBACK TO xlsximport; RELEASE xlsximport;",
      0, 0, &zErr2
    );
    sqlite3_free(zErr2);
  }

  if( rc != SQLITE_OK ) goto error;

  sqlite3_result_int(context, nImported);
  goto cleanup;

error:
  if( zErr==0 ){
    zErr = xlsxStrdup(sqlite3_errmsg(db));
  }
  if( zErr ){
    sqlite3_result_error(context, zErr, -1);
  }else{
    sqlite3_result_error(context, "xlsx_import(): out of memory", -1);
  }

cleanup:
  sqlite3_free(aSel);
  xlsxSheetArrayFree(&sheets);
  xlsxRelArrayFree(&rels);
  xlsxStrArrayFree(&ss);
  sqlite3_free(zErr);
}

/* --------------------------------------------------------------------- */
/* SQL scalar function: xlsx_import_version()                             */
/* --------------------------------------------------------------------- */

static void xlsxVersionFunc(
  sqlite3_context *context,
  int argc,
  sqlite3_value **argv
){
  (void)argc;
  (void)argv;
  sqlite3_result_text(
    context,
    "2026-11-03 Qwen3.8-Max (Thinking)",
    -1,
    SQLITE_STATIC
  );
}

/* --------------------------------------------------------------------- */
/* Table-valued function: xlsx_import_sheetnames(filename)                */
/* --------------------------------------------------------------------- */

typedef struct XlsxNamesVtab {
  sqlite3_vtab base;
  sqlite3 *db;
} XlsxNamesVtab;

typedef struct XlsxNamesCursor {
  sqlite3_vtab_cursor base;
  char *zFile;
  char **azNames;
  int nNames;
  int i;
} XlsxNamesCursor;

static void xlsxVtabSetError(sqlite3_vtab *pVtab, char *zMsg){
  if( pVtab->zErrMsg ) sqlite3_free(pVtab->zErrMsg);
  pVtab->zErrMsg = zMsg;
}

static int xlsxNamesConnect(
  sqlite3 *db,
  void *pAux,
  int argc,
  const char *const *argv,
  sqlite3_vtab **ppVTab,
  char **pzErr
){
  XlsxNamesVtab *p;
  int rc;

  (void)pAux;
  (void)argc;
  (void)argv;
  (void)pzErr;

  p = (XlsxNamesVtab*)sqlite3_malloc64(sizeof(XlsxNamesVtab));
  if( p==0 ) return SQLITE_NOMEM;

  memset(p, 0, sizeof(*p));
  p->db = db;

  *ppVTab = &p->base;

  rc = sqlite3_declare_vtab(
    db,
    "CREATE TABLE x(sheet_num INTEGER, sheet_name TEXT, filename TEXT HIDDEN)"
  );

  if( rc != SQLITE_OK ){
    sqlite3_free(p);
    *ppVTab = 0;
  }

  return rc;
}

static int xlsxNamesDisconnect(sqlite3_vtab *pVTab){
  sqlite3_free(pVTab);
  return SQLITE_OK;
}

static int xlsxNamesBestIndex(
  sqlite3_vtab *pVTab,
  sqlite3_index_info *info
){
  int i;
  int iFile = -1;

  for(i=0; i<info->nConstraint; i++){
    int op;

    if( !info->aConstraint[i].usable ) continue;
    if( info->aConstraint[i].iColumn != 2 ) continue;  /* filename HIDDEN */

    op = info->aConstraint[i].op;

    if( op == SQLITE_INDEX_CONSTRAINT_EQ ||
        op >= SQLITE_INDEX_CONSTRAINT_FUNCTION ){
      iFile = i;
      break;
    }
  }

  if( iFile < 0 ){
    xlsxVtabSetError(
      pVTab,
      sqlite3_mprintf("xlsx_import_sheetnames() requires a filename argument")
    );
    return SQLITE_CONSTRAINT;
  }

  info->aConstraintUsage[iFile].argvIndex = 1;
  info->aConstraintUsage[iFile].omit = 1;

  info->idxNum = 1;
  info->estimatedCost = 1.0;
  info->estimatedRows = 100;

  return SQLITE_OK;
}

static int xlsxNamesOpen(
  sqlite3_vtab *pVTab,
  sqlite3_vtab_cursor **ppCursor
){
  XlsxNamesCursor *pCur;

  (void)pVTab;

  pCur = (XlsxNamesCursor*)sqlite3_malloc64(sizeof(XlsxNamesCursor));
  if( pCur==0 ) return SQLITE_NOMEM;

  memset(pCur, 0, sizeof(*pCur));
  pCur->base.pVtab = pVTab;

  *ppCursor = &pCur->base;
  return SQLITE_OK;
}

static void xlsxNamesCursorClear(XlsxNamesCursor *p){
  int i;

  if( p->azNames ){
    for(i=0; i<p->nNames; i++){
      sqlite3_free(p->azNames[i]);
    }
    sqlite3_free(p->azNames);
  }

  sqlite3_free(p->zFile);

  p->zFile = 0;
  p->azNames = 0;
  p->nNames = 0;
  p->i = 0;
}

static int xlsxNamesClose(sqlite3_vtab_cursor *pCursor){
  XlsxNamesCursor *p = (XlsxNamesCursor*)pCursor;
  xlsxNamesCursorClear(p);
  sqlite3_free(p);
  return SQLITE_OK;
}

static int xlsxNamesFilter(
  sqlite3_vtab_cursor *pCursor,
  int idxNum,
  const char *idxStr,
  int argc,
  sqlite3_value **argv
){
  XlsxNamesCursor *p = (XlsxNamesCursor*)pCursor;
  XlsxNamesVtab *pVtab = (XlsxNamesVtab*)pCursor->pVtab;
  XlsxSheetArray sheets = {0,0,0};
  const char *zFile;
  char *zErr = 0;
  int rc;
  int i;

  (void)idxStr;

  xlsxNamesCursorClear(p);

  if( idxNum != 1 || argc < 1 ){
    xlsxVtabSetError(
      pCursor->pVtab,
      sqlite3_mprintf("xlsx_import_sheetnames() requires a filename argument")
    );
    return SQLITE_ERROR;
  }

  zFile = (const char*)sqlite3_value_text(argv[0]);
  if( zFile==0 ){
    xlsxVtabSetError(
      pCursor->pVtab,
      sqlite3_mprintf("xlsx_import_sheetnames(): filename must not be NULL")
    );
    return SQLITE_ERROR;
  }

  p->zFile = xlsxStrdup(zFile);
  if( p->zFile==0 ) return SQLITE_NOMEM;

  rc = xlsxLoadWorkbook(pVtab->db, zFile, &sheets, &zErr);
  if( rc != SQLITE_OK ){
    xlsxVtabSetError(
      pCursor->pVtab,
      zErr ? zErr : sqlite3_mprintf("xlsx_import_sheetnames(): cannot read \"%s\"", zFile)
    );
    xlsxSheetArrayFree(&sheets);
    return rc;
  }

  if( sheets.n > 0 ){
    p->azNames = (char**)sqlite3_malloc64((sqlite3_int64)sheets.n * sizeof(char*));
    if( p->azNames==0 ){
      xlsxSheetArrayFree(&sheets);
      return SQLITE_NOMEM;
    }

    memset(p->azNames, 0, (size_t)sheets.n * sizeof(char*));

    for(i=0; i<sheets.n; i++){
      p->azNames[i] = xlsxStrdup(sheets.a[i].zName);
      if( p->azNames[i]==0 ){
        int j;
        for(j=0; j<i; j++) sqlite3_free(p->azNames[j]);
        sqlite3_free(p->azNames);
        p->azNames = 0;
        xlsxSheetArrayFree(&sheets);
        return SQLITE_NOMEM;
      }
    }

    p->nNames = sheets.n;
  }

  xlsxSheetArrayFree(&sheets);
  p->i = 0;
  return SQLITE_OK;
}

static int xlsxNamesNext(sqlite3_vtab_cursor *pCursor){
  XlsxNamesCursor *p = (XlsxNamesCursor*)pCursor;
  p->i++;
  return SQLITE_OK;
}

static int xlsxNamesEof(sqlite3_vtab_cursor *pCursor){
  XlsxNamesCursor *p = (XlsxNamesCursor*)pCursor;
  return p->i >= p->nNames;
}

static int xlsxNamesColumn(
  sqlite3_vtab_cursor *pCursor,
  sqlite3_context *context,
  int iColumn
){
  XlsxNamesCursor *p = (XlsxNamesCursor*)pCursor;

  if( p->i < 0 || p->i >= p->nNames ){
    sqlite3_result_null(context);
    return SQLITE_OK;
  }

  if( iColumn == 0 ){
    /* sheet_num: 1-based sheet number in workbook.xml order */
    sqlite3_result_int(context, p->i + 1);
  }else if( iColumn == 1 ){
    /* sheet_name */
    sqlite3_result_text(
      context,
      p->azNames[p->i] ? p->azNames[p->i] : "",
      -1,
      SQLITE_TRANSIENT
    );
  }else if( iColumn == 2 ){
    /* hidden filename column */
    sqlite3_result_text(
      context,
      p->zFile ? p->zFile : "",
      -1,
      SQLITE_TRANSIENT
    );
  }else{
    sqlite3_result_null(context);
  }

  return SQLITE_OK;
}

static int xlsxNamesRowid(
  sqlite3_vtab_cursor *pCursor,
  sqlite3_int64 *pRowid
){
  XlsxNamesCursor *p = (XlsxNamesCursor*)pCursor;
  *pRowid = p->i + 1;
  return SQLITE_OK;
}

static sqlite3_module xlsxSheetNamesModule;

/* --------------------------------------------------------------------- */
/* Extension entry point                                                  */
/* --------------------------------------------------------------------- */

#ifdef _WIN32
#define XLSXIMPORT_EXPORT __declspec(dllexport)
#else
#define XLSXIMPORT_EXPORT
#endif

XLSXIMPORT_EXPORT int sqlite3_xlsximport_init(
  sqlite3 *db,
  char **pzErrMsg,
  const sqlite3_api_routines *pApi
){
  int rc;

  SQLITE_EXTENSION_INIT2(pApi);

  /* Initialize the eponymous virtual table module.
  ** Doing this at run time avoids -Wmissing-field-initializers warnings
  ** and remains compatible with newer sqlite3_module fields that are
  ** left zeroed.
  */
  xlsxSheetNamesModule.iVersion     = 0;
  xlsxSheetNamesModule.xCreate      = 0;
  xlsxSheetNamesModule.xConnect     = xlsxNamesConnect;
  xlsxSheetNamesModule.xBestIndex   = xlsxNamesBestIndex;
  xlsxSheetNamesModule.xDisconnect  = xlsxNamesDisconnect;
  xlsxSheetNamesModule.xDestroy     = 0;
  xlsxSheetNamesModule.xOpen        = xlsxNamesOpen;
  xlsxSheetNamesModule.xClose       = xlsxNamesClose;
  xlsxSheetNamesModule.xFilter      = xlsxNamesFilter;
  xlsxSheetNamesModule.xNext        = xlsxNamesNext;
  xlsxSheetNamesModule.xEof         = xlsxNamesEof;
  xlsxSheetNamesModule.xColumn      = xlsxNamesColumn;
  xlsxSheetNamesModule.xRowid       = xlsxNamesRowid;

  rc = sqlite3_create_function(
    db,
    "xlsx_import",
    -1,
    SQLITE_UTF8,
    0,
    xlsxImportFunc,
    0,
    0
  );

  if( rc == SQLITE_OK ){
    rc = sqlite3_create_function(
      db,
      "xlsx_import_version",
      0,
      SQLITE_UTF8 | SQLITE_DETERMINISTIC,
      0,
      xlsxVersionFunc,
      0,
      0
    );
  }

  if( rc == SQLITE_OK ){
    rc = sqlite3_create_module(
      db,
      "xlsx_import_sheetnames",
      &xlsxSheetNamesModule,
      0
    );
  }

  if( rc != SQLITE_OK && pzErrMsg ){
    *pzErrMsg = sqlite3_mprintf(
      "xlsximport: initialization failed: %s",
      sqlite3_errmsg(db)
    );
  }

  return rc;
}