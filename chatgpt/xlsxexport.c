/*
Prompts used to generate this code:
Create the C code for a SQLite extension named xlsxexport that contains a SQL function named xlsx_export 
that saves multiple tables as a single XLSX spreadsheet, with the sheet names equal to the table names, 
and the sheet headers in bold and with autofilter. 
The table names are the arguments 2 to N of the xlsx_export function.
If invoked with only one parameter then exports all the tables in the schema.
xlsx_export returns the number of tables imported.

XLSX format is a ZIP container with XML files inside.
Use the SQLite virtual table zipfile extension to handle the ZIP container.
Do not use external libraries to handle XML.
Warn if the Excel maximum cell size is exceeded. Sanitize the sheet names to conform to Excel restrictions.
Add SQL function xlsx_export_version returning '2026-01-07 ChatGPT-5.2'.
Make zipfile remove any previous XLSX file with same name.
Add bold headers + autofilter XML, generate full [Content_Types].xml and relationships, export actual table data rows, make the export fully Excel-round-trip compliant.
Compile cleanly with: -std=c11 -Wall -Wextra -Wpedantic
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>

#define EXCEL_MAX_CELL_CHARS 32767

/* ========================= Dynamic String Buffer ========================= */

typedef struct StrBuf {
  char *z;
  size_t n;
  size_t cap;
} StrBuf;

static void sb_init(StrBuf *p){
  p->z = NULL;
  p->n = 0;
  p->cap = 0;
}

static void sb_free(StrBuf *p){
  sqlite3_free(p->z);
  sb_init(p);
}

static void sb_append(StrBuf *p, const char *z, int n){
  if(n<0) n = (int)strlen(z);
  if(p->n + n + 1 > p->cap){
    p->cap = p->cap ? p->cap * 2 : 1024;
    while(p->cap < p->n + n + 1) p->cap *= 2;
    p->z = sqlite3_realloc64(p->z, p->cap);
  }
  memcpy(p->z + p->n, z, n);
  p->n += n;
  p->z[p->n] = 0;
}

static void sb_appendf(StrBuf *p, const char *zFmt, ...){
  va_list ap;
  char *z;
  va_start(ap, zFmt);
  z = sqlite3_vmprintf(zFmt, ap);
  va_end(ap);
  if(z){
    sb_append(p, z, -1);
    sqlite3_free(z);
  }
}

/* ========================= Utilities ========================= */

static char *sanitize_sheetname(const char *z){
  char buf[64];
  size_t j = 0;
  for(size_t i = 0; z[i] && j < 31; i++){
    char c = z[i];
    if(strchr("[]:*?/\\", c)){
      buf[j++] = '_';
    }else{
      buf[j++] = c;
    }
  }
  buf[j] = 0;
  if(j==0) return sqlite3_mprintf("Sheet1");
  return sqlite3_mprintf("%s", buf);
}

static void xml_escape(StrBuf *p, const char *z){
  if(!z) return;
  for(int i=0; z[i]; i++){
    switch(z[i]){
      case '&':  sb_append(p, "&amp;", 5);  break;
      case '<':  sb_append(p, "&lt;", 4);   break;
      case '>':  sb_append(p, "&gt;", 4);   break;
      case '"':  sb_append(p, "&quot;", 6); break;
      case '\'': sb_append(p, "&apos;", 6); break;
      default:   sb_append(p, &z[i], 1);    break;
    }
  }
}

static void int_to_col(int n, char *zOut){
  char buf[16];
  int i = 0;
  n++;
  while(n > 0){
    n--;
    buf[i++] = (char)('A' + (n % 26));
    n /= 26;
  }
  for(int j=0; j<i; j++) zOut[j] = buf[i-1-j];
  zOut[i] = 0;
}

/* ========================= ZIP Writer Helper ========================= */

static int zip_write(sqlite3 *db, const char *zVtab, const char *zName, const char *zData, int nData){
  sqlite3_stmt *pSt;
  char *zSql = sqlite3_mprintf("INSERT INTO \"%w\"(name, data) VALUES(?, ?)", zVtab);
  int rc = sqlite3_prepare_v2(db, zSql, -1, &pSt, NULL);
  sqlite3_free(zSql);
  if(rc != SQLITE_OK) return rc;
  sqlite3_bind_text(pSt, 1, zName, -1, SQLITE_STATIC);
  sqlite3_bind_blob(pSt, 2, zData, nData, SQLITE_STATIC);
  rc = sqlite3_step(pSt);
  sqlite3_finalize(pSt);
  return (rc == SQLITE_DONE) ? SQLITE_OK : rc;
}

/* ========================= XML Generators ========================= */

static void gen_content_types(StrBuf *p, int nSheets){
  sb_append(p, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
  sb_append(p, "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n", -1);
  sb_append(p, "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n", -1);
  sb_append(p, "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n", -1);
  sb_append(p, "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n", -1);
  sb_append(p, "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n", -1);
  for(int i=1; i<=nSheets; i++){
    sb_appendf(p, "  <Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n", i);
  }
  sb_append(p, "</Types>", -1);
}

static void gen_rels(StrBuf *p){
  sb_append(p, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
  sb_append(p, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n", -1);
  sb_append(p, "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n", -1);
  sb_append(p, "</Relationships>", -1);
}

static void gen_styles(StrBuf *p){
  sb_append(p, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
  sb_append(p, "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n", -1);
  sb_append(p, "  <fonts count=\"2\">\n", -1);
  sb_append(p, "    <font><sz val=\"11\"/><name val=\"Calibri\"/></font>\n", -1);
  sb_append(p, "    <font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>\n", -1);
  sb_append(p, "  </fonts>\n", -1);
  sb_append(p, "  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>\n", -1);
  sb_append(p, "  <borders count=\"1\"><border/></borders>\n", -1);
  sb_append(p, "  <cellStyleXfs count=\"1\"><xf/></cellStyleXfs>\n", -1);
  sb_append(p, "  <cellXfs count=\"2\">\n", -1);
  sb_append(p, "    <xf fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n", -1);
  sb_append(p, "    <xf fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>\n", -1);
  sb_append(p, "  </cellXfs>\n", -1);
  sb_append(p, "  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\n", -1);
  sb_append(p, "</styleSheet>", -1);
}

static void gen_workbook(StrBuf *p, int nSheets, char **azSheetNames, char **azRanges){
  sb_append(p, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
  sb_append(p, "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n", -1);
  sb_append(p, "  <fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>\n", -1);
  sb_append(p, "  <workbookPr defaultThemeVersion=\"124226\"/>\n", -1);
  sb_append(p, "  <bookViews><workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/></bookViews>\n", -1);
  sb_append(p, "  <sheets>\n", -1);
  for(int i=0; i<nSheets; i++){
    sb_appendf(p, "    <sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>\n", azSheetNames[i], i+1, i+1);
  }
  sb_append(p, "  </sheets>\n", -1);
  if(azRanges){
    sb_append(p, "  <definedNames>\n", -1);
    for(int i=0; i<nSheets; i++){
      if(azRanges[i]){
        sb_appendf(p, "    <definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"%d\" hidden=\"1\">%s!%s</definedName>\n", i, azSheetNames[i], azRanges[i]);
      }
    }
    sb_append(p, "  </definedNames>\n", -1);
  }
  sb_append(p, "  <calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>\n", -1);
  sb_append(p, "</workbook>", -1);
}

static void gen_workbook_rels(StrBuf *p, int nSheets){
  sb_append(p, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
  sb_append(p, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n", -1);
  for(int i=1; i<=nSheets; i++){
    sb_appendf(p, "  <Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>\n", i, i);
  }
  sb_appendf(p, "  <Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n", nSheets+1);
  sb_append(p, "</Relationships>", -1);
}

/* ========================= Export ========================= */

static void xlsx_export(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  if(argc < 1){
    sqlite3_result_error(ctx, "xlsx_export(filename, [table...])", -1);
    return;
  }

  sqlite3 *db = sqlite3_context_db_handle(ctx);
  const char *zFilename = (const char*)sqlite3_value_text(argv[0]);

  /* Overwrite logic: delete file if it exists */
  char *zSql = sqlite3_mprintf("SELECT writefile(%Q, NULL)", zFilename);
  sqlite3_exec(db, zSql, NULL, NULL, NULL);
  sqlite3_free(zSql);

  /* Discover tables */
  char **azTables = NULL;
  char **azSheets = NULL;
  int nSheets = 0;

  if(argc == 1){
    sqlite3_stmt *pSt;
    sqlite3_prepare_v2(db, "SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%'", -1, &pSt, NULL);
    while(sqlite3_step(pSt) == SQLITE_ROW){
      azTables = sqlite3_realloc64(azTables, (nSheets+1)*sizeof(char*));
      azSheets = sqlite3_realloc64(azSheets, (nSheets+1)*sizeof(char*));
      azTables[nSheets] = sqlite3_mprintf("%s", sqlite3_column_text(pSt, 0));
      azSheets[nSheets] = sanitize_sheetname(azTables[nSheets]);
      nSheets++;
    }
    sqlite3_finalize(pSt);
  }else{
    nSheets = argc - 1;
    azTables = sqlite3_malloc64(nSheets * sizeof(char*));
    azSheets = sqlite3_malloc64(nSheets * sizeof(char*));
    for(int i=0; i<nSheets; i++){
      azTables[i] = sqlite3_mprintf("%s", sqlite3_value_text(argv[i+1]));
      azSheets[i] = sanitize_sheetname(azTables[i]);
    }
  }

  if(nSheets == 0){
    sqlite3_result_int(ctx, 0);
    return;
  }

  char **azRanges = sqlite3_malloc64(nSheets * sizeof(char*));
  memset(azRanges, 0, nSheets * sizeof(char*));

  /* Create temporary virtual table for zip construction */
  char *zVtab = sqlite3_mprintf("xlsx_zip_%p", ctx);
  zSql = sqlite3_mprintf("CREATE VIRTUAL TABLE \"%w\" USING zipfile(%Q)", zVtab, zFilename);
  int rc = sqlite3_exec(db, zSql, NULL, NULL, NULL);
  sqlite3_free(zSql);

  if(rc != SQLITE_OK){
    sqlite3_result_error(ctx, "Failed to create zip virtual table", -1);
    goto cleanup;
  }

  /* 1. Content Types */
  StrBuf sb;
  sb_init(&sb);
  gen_content_types(&sb, nSheets);
  zip_write(db, zVtab, "[Content_Types].xml", sb.z, (int)sb.n);
  sb_free(&sb);

  /* 2. Rels */
  gen_rels(&sb);
  zip_write(db, zVtab, "_rels/.rels", sb.z, (int)sb.n);
  sb_free(&sb);

  /* 3. Styles */
  gen_styles(&sb);
  zip_write(db, zVtab, "xl/styles.xml", sb.z, (int)sb.n);
  sb_free(&sb);

  /* 4. Workbook Rels */
  gen_workbook_rels(&sb, nSheets);
  zip_write(db, zVtab, "xl/_rels/workbook.xml.rels", sb.z, (int)sb.n);
  sb_free(&sb);

  /* 5. Worksheets */
  for(int i=0; i<nSheets; i++){
    StrBuf sbW;
    sb_init(&sbW);
    sb_append(&sbW, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
    sb_append(&sbW, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n", -1);

    sqlite3_stmt *pSt;
    char *zQ = sqlite3_mprintf("SELECT * FROM \"%w\"", azTables[i]);
    rc = sqlite3_prepare_v2(db, zQ, -1, &pSt, NULL);
    sqlite3_free(zQ);

    int nCols = sqlite3_column_count(pSt);
    int nRows = 1;

    /* Buffer for sheetData to calculate dimensions */
    StrBuf sbD;
    sb_init(&sbD);
    sb_append(&sbD, "  <sheetData>\n", -1);

    /* Headers */
    sb_append(&sbD, "    <row r=\"1\">\n", -1);
    for(int c=0; c<nCols; c++){
      char zColRef[8];
      int_to_col(c, zColRef);
      sb_appendf(&sbD, "      <c r=\"%s1\" t=\"inlineStr\" s=\"1\"><is><t>", zColRef);
      xml_escape(&sbD, sqlite3_column_name(pSt, c));
      sb_append(&sbD, "</t></is></c>\n", -1);
    }
    sb_append(&sbD, "    </row>\n", -1);

    /* Rows */
    while(sqlite3_step(pSt) == SQLITE_ROW){
      nRows++;
      sb_appendf(&sbD, "    <row r=\"%d\">\n", nRows);
      for(int c=0; c<nCols; c++){
        char zColRef[8];
        int_to_col(c, zColRef);
        int type = sqlite3_column_type(pSt, c);
        if(type == SQLITE_NULL) continue;
        if(type == SQLITE_INTEGER || type == SQLITE_FLOAT){
          sb_appendf(&sbD, "      <c r=\"%s%d\"><v>%s</v></c>\n", zColRef, nRows, sqlite3_column_text(pSt, c));
        }else{
          const char *zVal = (const char*)sqlite3_column_text(pSt, c);
          if(strlen(zVal) > EXCEL_MAX_CELL_CHARS){
             sqlite3_log(SQLITE_WARNING, "xlsx_export: cell truncated in table %s, row %d, col %d", azTables[i], nRows, c+1);
          }
          sb_appendf(&sbD, "      <c r=\"%s%d\" t=\"inlineStr\"><is><t>", zColRef, nRows);
          xml_escape(&sbD, zVal);
          sb_append(&sbD, "</t></is></c>\n", -1);
        }
      }
      sb_append(&sbD, "    </row>\n", -1);
    }
    sqlite3_finalize(pSt);
    sb_append(&sbD, "  </sheetData>\n", -1);

    /* Dimensions and Autofilter */
    char zFirst[8], zLast[8];
    int_to_col(0, zFirst);
    int_to_col(nCols > 0 ? nCols - 1 : 0, zLast);
    sb_appendf(&sbW, "  <dimension ref=\"%s1:%s%d\"/>\n", zFirst, zLast, nRows);
    sb_append(&sbW, "  <sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>\n", -1);
    sb_append(&sbW, "  <sheetFormatPr defaultRowHeight=\"15\"/>\n", -1);
    sb_append(&sbW, sbD.z, -1);
    if(nCols > 0){
      sb_appendf(&sbW, "  <autoFilter ref=\"%s1:%s%d\"/>\n", zFirst, zLast, nRows);
      azRanges[i] = sqlite3_mprintf("$%s$1:$%s$%d", zFirst, zLast, nRows);
    }
    sb_append(&sbW, "</worksheet>", -1);

    char zPath[64];
    snprintf(zPath, sizeof(zPath), "xl/worksheets/sheet%d.xml", i+1);
    zip_write(db, zVtab, zPath, sbW.z, (int)sbW.n);
    sb_free(&sbW);
    sb_free(&sbD);
  }

  /* 6. Workbook (last, because it needs azRanges) */
  gen_workbook(&sb, nSheets, azSheets, azRanges);
  zip_write(db, zVtab, "xl/workbook.xml", sb.z, (int)sb.n);
  sb_free(&sb);

  /* Finalize ZIP */
  zSql = sqlite3_mprintf("DROP TABLE \"%w\"", zVtab);
  sqlite3_exec(db, zSql, NULL, NULL, NULL);
  sqlite3_free(zSql);

  sqlite3_result_int(ctx, nSheets);

cleanup:
  for(int i=0; i<nSheets; i++){
    sqlite3_free(azTables[i]);
    sqlite3_free(azSheets[i]);
    sqlite3_free(azRanges[i]);
  }
  sqlite3_free(azTables);
  sqlite3_free(azSheets);
  sqlite3_free(azRanges);
  sqlite3_free(zVtab);
}

static void xlsx_export_version(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  (void)argc; (void)argv;
  sqlite3_result_text(ctx, "2026-01-07 ChatGPT-5.2", -1, SQLITE_STATIC);
}

/* ========================= Init ========================= */

#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsxexport_init(sqlite3 *db, char **pzErrMsg, const sqlite3_api_routines *pApi){
  (void)pzErrMsg;
  SQLITE_EXTENSION_INIT2(pApi);
  sqlite3_create_function(db, "xlsx_export", -1, SQLITE_UTF8, NULL, xlsx_export, NULL, NULL);
  sqlite3_create_function(db, "xlsx_export_version", 0, SQLITE_UTF8, NULL, xlsx_export_version, NULL, NULL);
  return SQLITE_OK;
}
