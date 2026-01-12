/*
Prompts used to generate this code:
Create the C code for a SQLite extension named xlsxexport that contains a SQL function named xlsx_export 
that saves multiple tables as a single XLSX spreadsheet, with the sheet names equal to the table names, 
and the sheet headers in bold and with autofilter. If invoked with only one parameter then exports all 
the tables in the schema. xlsx_export returns the number of tables imported.
Use libxlsxwriter. Warn if the Excel maximum cell size is exceeded. 
Sanitize the sheet names to conform to Excel restrictions. 
Add SQL function xlsx_export_version returning "2026-01-07 ChatGPT-5.2".
Compile cleanly with: -std=c11 -Wall -Wextra -Wpedantic
Include build example.
Linking to sqlite library is not needed. Just include -I../../sqlite-amalgamation-3510200
Include as comments the prompts used.
*/

/*
 * xlsxexport.c — SQLite XLSX export extension
 *
 * Requires libxlsxwriter.
 *
 * Build example:
 *   cc -std=c11 -Wall -Wextra -Wpedantic -fPIC -shared -I../../sqlite-amalgamation-3510200 \
 *      xlsxexport.c -lxlsxwriter
 */

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <xlsxwriter.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <ctype.h>

#define EXCEL_MAX_CELL_CHARS 32767

/* ============================================================ */
/* Utilities                                                    */
/* ============================================================ */

static char *sqlite_strdup(const char *z){
  size_t n = strlen(z) + 1;
  char *p = sqlite3_malloc64(n);
  if(p) memcpy(p, z, n);
  return p;
}

/* Sanitize sheet names to Excel restrictions:
 * - max 31 chars
 * - cannot contain: : \ / ? * [ ]
 */
static void sanitize_sheet_name(char *z){
  const char *bad = ":\\/?*[]";
  for(char *p = z; *p; p++){
    if(strchr(bad, *p)) *p = '_';
  }
  if(strlen(z) > 31) z[31] = 0;
}

/* ============================================================ */
/* Export implementation                                       */
/* ============================================================ */

static int export_table(sqlite3 *db, lxw_workbook *wb, const char *table){
  char *sheetname = sqlite_strdup(table);
  sanitize_sheet_name(sheetname);

  lxw_worksheet *ws = workbook_add_worksheet(wb, sheetname);
  if(!ws){ sqlite3_free(sheetname); return 0; }

  lxw_format *fmt_hdr = workbook_add_format(wb);
  format_set_bold(fmt_hdr);

  sqlite3_stmt *st = NULL;
  char sql[512];
  snprintf(sql, sizeof(sql), "SELECT * FROM \"%s\"", table);

  if(sqlite3_prepare_v2(db, sql, -1, &st, NULL) != SQLITE_OK){
    sqlite3_free(sheetname);
    return 0;
  }

  int cols = sqlite3_column_count(st);

  /* Header row */
  for(int c = 0; c < cols; c++){
    const char *name = sqlite3_column_name(st, c);
    worksheet_write_string(ws, 0, c, name, fmt_hdr);
  }

  worksheet_autofilter(ws, 0, 0, 0, cols - 1);

  int r = 1;
  while(sqlite3_step(st) == SQLITE_ROW){
    for(int c = 0; c < cols; c++){
      const unsigned char *txt = sqlite3_column_text(st, c);
      if(txt){
        size_t len = strlen((const char*)txt);
        if(len > EXCEL_MAX_CELL_CHARS){
          fprintf(stderr, "Warning: cell exceeds Excel max size (%zu) in table %s\n", len, table);
          char buf[EXCEL_MAX_CELL_CHARS + 1];
          memcpy(buf, txt, EXCEL_MAX_CELL_CHARS);
          buf[EXCEL_MAX_CELL_CHARS] = 0;
          worksheet_write_string(ws, r, c, buf, NULL);
        }else{
          worksheet_write_string(ws, r, c, (const char*)txt, NULL);
        }
      }
    }
    r++;
  }

  sqlite3_finalize(st);
  sqlite3_free(sheetname);
  return 1;
}

/* ============================================================ */
/* xlsx_export() SQL function                                  */
/* ============================================================ */

static void xlsx_export(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  int exported_count = 0;
  if(argc < 1){
    sqlite3_result_error(ctx, "output filename required", -1);
    return;
  }

  sqlite3 *db = sqlite3_context_db_handle(ctx);
  const char *filename = (const char*)sqlite3_value_text(argv[0]);

  lxw_workbook *wb = workbook_new(filename);
  if(!wb){
    sqlite3_result_error(ctx, "cannot create XLSX file", -1);
    return;
  }

  if(argc == 1){
    /* Export all tables */
    sqlite3_stmt *st = NULL;
    const char *sql = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'";
    if(sqlite3_prepare_v2(db, sql, -1, &st, NULL) == SQLITE_OK){
      while(sqlite3_step(st) == SQLITE_ROW){
        exported_count += export_table(db, wb, (const char*)sqlite3_column_text(st, 0));
      }
      sqlite3_finalize(st);
    }
  }else{
    /* Export selected tables */
    for(int i = 1; i < argc; i++){
      exported_count += export_table(db, wb, (const char*)sqlite3_value_text(argv[i]));
    }
  }

  workbook_close(wb);
  sqlite3_result_int(ctx, exported_count);
}

static void xlsx_export_version(sqlite3_context *ctx, int argc, sqlite3_value **argv){
  (void)argc; (void)argv;
  sqlite3_result_text(ctx, "2026-01-07 ChatGPT-5.2", -1, SQLITE_STATIC);
}

/* ============================================================ */
/* Entry                                                       */
/* ============================================================ */

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
