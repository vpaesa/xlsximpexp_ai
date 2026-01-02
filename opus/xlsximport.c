/*
xlsximport.c - SQLite extension to import XLSX files

Uses the SQLite zipfile extension to read XLSX archives and expat for XML
parsing.
Two SQL functions defined
xlsx_import() creates one table for each sheet in the XLSX file, with table name
equal to sheet name, and column names equal to the values in the first row of
the sheet. xlsx_import_version() returns the version string.

Usage:
SELECT xlsx_import('filename.xlsx');
SELECT xlsx_import_version();
**
** ============================================================================
** REQUIREMENTS / DESIGN NOTES
** ============================================================================
**
** This extension uses the SQLite zipfile extension to open XLSX files and
** gather the following content:
**   - xl/sharedStrings.xml
**   - xl/worksheets/sheet1.xml to xl/worksheets/sheetN.xml
**   - xl/workbook.xml
**
** The name of each sheet is stored in xl/workbook.xml.
** The individual sheets are kept in xl/worksheets/sheet1.xml to sheetN.xml.
**
** To save on space, Microsoft stores all character literal values in one
** common xl/sharedStrings.xml dictionary file. The individual cell value
** found for a string in the actual sheet.xml file is just an index into
** this dictionary.
**
** Microsoft does not store empty cells or rows in xl/worksheets/sheet*.xml,
** so any gaps between values must be handled by this code.
**
** COLUMN NAME CONVERSION (Base-26):
** To figure out the number of skipped columns, we need to calculate the
** distance between cells like "AB67" and "C67". The way columns are named
** (A through Z, then AA through AZ, then BA through BZ, etc.) suggests a
** base-26 system. We use a simple conversion method from base-26 to decimal
** and then use subtraction to find empty cells between columns.
**
** XML STRUCTURE NOTES:
**   - xl/sharedStrings.xml has "sst:uniqueCount" with count of unique strings
**   - xl/worksheets/sheet*.xml has "dimension:ref" with enclosing cell range
**   - Inline strings use <c t="inlineStr"><is><t>value</t></is></c>
**   - Shared strings use <c t="s"><v>index</v></c>
**   - Numeric values use <c><v>number</v></c> (no type attribute)
**
** XML parsing is done using the expat library.
**
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <ctype.h>
#include <expat.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

/*
** ============================================================================
** Utility Functions
** ============================================================================
*/

/*
** Convert a column name (A, B, ..., Z, AA, AB, ...) to a 1-based column number.
** A=1, B=2, ..., Z=26, AA=27, AB=28, ...
*/
static int col_to_num(const char *col) {
  int num = 0;
  while (*col && isalpha((unsigned char)*col)) {
    num = num * 26 + (toupper((unsigned char)*col) - 'A' + 1);
    col++;
  }
  return num;
}

/*
** Parse a cell reference like "AB67" into column number and row number.
** Returns the column number (1-based) and sets *row to the row number
* (1-based).
*/
static int parse_cell_ref(const char *ref, int *row) {
  int col = 0;
  const char *p = ref;

  /* Parse column letters */
  while (*p && isalpha((unsigned char)*p)) {
    col = col * 26 + (toupper((unsigned char)*p) - 'A' + 1);
    p++;
  }

  /* Parse row number */
  if (row) {
    *row = atoi(p);
  }

  return col;
}

/*
** ============================================================================
** Shared Strings Parser
** ============================================================================
*/

typedef struct {
  char **strings;  /* Array of string values */
  int count;       /* Number of strings */
  int capacity;    /* Allocated capacity */
  int in_t;        /* Currently inside <t> element */
  char *current;   /* Current string being built */
  int current_len; /* Length of current string */
  int current_cap; /* Capacity of current string buffer */
} SharedStrings;

static void ss_init(SharedStrings *ss) { memset(ss, 0, sizeof(*ss)); }

static void ss_free(SharedStrings *ss) {
  for (int i = 0; i < ss->count; i++) {
    free(ss->strings[i]);
  }
  free(ss->strings);
  free(ss->current);
  memset(ss, 0, sizeof(*ss));
}

static void ss_add_string(SharedStrings *ss, const char *str) {
  if (ss->count >= ss->capacity) {
    int new_cap = ss->capacity ? ss->capacity * 2 : 64;
    char **new_strings = realloc(ss->strings, new_cap * sizeof(char *));
    if (!new_strings)
      return;
    ss->strings = new_strings;
    ss->capacity = new_cap;
  }
  ss->strings[ss->count++] = strdup(str ? str : "");
}

static void ss_append_text(SharedStrings *ss, const char *text, int len) {
  if (ss->current_len + len >= ss->current_cap) {
    int new_cap = ss->current_cap ? ss->current_cap * 2 : 256;
    while (new_cap <= ss->current_len + len)
      new_cap *= 2;
    char *new_current = realloc(ss->current, new_cap);
    if (!new_current)
      return;
    ss->current = new_current;
    ss->current_cap = new_cap;
  }
  memcpy(ss->current + ss->current_len, text, len);
  ss->current_len += len;
  ss->current[ss->current_len] = '\0';
}

static void XMLCALL ss_start_element(void *userData, const XML_Char *name,
                                     const XML_Char **atts) {
  SharedStrings *ss = (SharedStrings *)userData;
  (void)atts;

  if (strcmp(name, "t") == 0) {
    ss->in_t = 1;
    ss->current_len = 0;
    if (ss->current)
      ss->current[0] = '\0';
  }
}

static void XMLCALL ss_end_element(void *userData, const XML_Char *name) {
  SharedStrings *ss = (SharedStrings *)userData;

  if (strcmp(name, "si") == 0) {
    /* End of string item - add accumulated text */
    ss_add_string(ss, ss->current ? ss->current : "");
    ss->current_len = 0;
    if (ss->current)
      ss->current[0] = '\0';
  } else if (strcmp(name, "t") == 0) {
    ss->in_t = 0;
  }
}

static void XMLCALL ss_char_data(void *userData, const XML_Char *s, int len) {
  SharedStrings *ss = (SharedStrings *)userData;

  if (ss->in_t) {
    ss_append_text(ss, s, len);
  }
}

static int parse_shared_strings(const char *xml, int xml_len,
                                SharedStrings *ss) {
  XML_Parser parser = XML_ParserCreate(NULL);
  if (!parser)
    return -1;

  ss_init(ss);

  XML_SetUserData(parser, ss);
  XML_SetElementHandler(parser, ss_start_element, ss_end_element);
  XML_SetCharacterDataHandler(parser, ss_char_data);

  if (XML_Parse(parser, xml, xml_len, 1) == XML_STATUS_ERROR) {
    XML_ParserFree(parser);
    ss_free(ss);
    return -1;
  }

  XML_ParserFree(parser);
  return 0;
}

/*
** ============================================================================
** Workbook Parser (Sheet Names)
** ============================================================================
*/

typedef struct {
  char *name;  /* Sheet name */
  int sheetId; /* Sheet ID */
} SheetInfo;

typedef struct {
  SheetInfo *sheets; /* Array of sheet info */
  int count;         /* Number of sheets */
  int capacity;      /* Allocated capacity */
} Workbook;

static void wb_init(Workbook *wb) { memset(wb, 0, sizeof(*wb)); }

static void wb_free(Workbook *wb) {
  for (int i = 0; i < wb->count; i++) {
    free(wb->sheets[i].name);
  }
  free(wb->sheets);
  memset(wb, 0, sizeof(*wb));
}

static void wb_add_sheet(Workbook *wb, const char *name, int sheetId) {
  if (wb->count >= wb->capacity) {
    int new_cap = wb->capacity ? wb->capacity * 2 : 8;
    SheetInfo *new_sheets = realloc(wb->sheets, new_cap * sizeof(SheetInfo));
    if (!new_sheets)
      return;
    wb->sheets = new_sheets;
    wb->capacity = new_cap;
  }
  wb->sheets[wb->count].name = strdup(name ? name : "");
  wb->sheets[wb->count].sheetId = sheetId;
  wb->count++;
}

static void XMLCALL wb_start_element(void *userData, const XML_Char *name,
                                     const XML_Char **atts) {
  Workbook *wb = (Workbook *)userData;

  if (strcmp(name, "sheet") == 0) {
    const char *sheet_name = NULL;
    int sheetId = 0;

    for (int i = 0; atts[i]; i += 2) {
      if (strcmp(atts[i], "name") == 0) {
        sheet_name = atts[i + 1];
      } else if (strcmp(atts[i], "sheetId") == 0) {
        sheetId = atoi(atts[i + 1]);
      }
    }

    if (sheet_name) {
      wb_add_sheet(wb, sheet_name, sheetId);
    }
  }
}

static void XMLCALL wb_end_element(void *userData, const XML_Char *name) {
  (void)userData;
  (void)name;
}

static int parse_workbook(const char *xml, int xml_len, Workbook *wb) {
  XML_Parser parser = XML_ParserCreate(NULL);
  if (!parser)
    return -1;

  wb_init(wb);

  XML_SetUserData(parser, wb);
  XML_SetElementHandler(parser, wb_start_element, wb_end_element);

  if (XML_Parse(parser, xml, xml_len, 1) == XML_STATUS_ERROR) {
    XML_ParserFree(parser);
    wb_free(wb);
    return -1;
  }

  XML_ParserFree(parser);
  return 0;
}

/*
** ============================================================================
** Worksheet Parser
** ============================================================================
*/

typedef struct {
  char *value; /* Cell value (string or numeric as string) */
  int is_null; /* Whether this cell is empty/null */
} CellValue;

typedef struct {
  CellValue *cells; /* Row of cells */
  int count;        /* Number of cells in row */
  int capacity;     /* Allocated capacity */
} Row;

typedef struct {
  Row *rows;    /* Array of rows */
  int count;    /* Number of rows */
  int capacity; /* Allocated capacity */
  int max_col;  /* Maximum column number seen */
} Worksheet;

typedef struct {
  Worksheet *ws;     /* Worksheet being built */
  SharedStrings *ss; /* Shared strings reference */

  /* Current cell state */
  int cur_row;   /* Current row number (1-based) */
  int cur_col;   /* Current column number (1-based) */
  char cur_type; /* Cell type: 's'=shared string, 'n'=number, 'i'=inline,
                    'b'=boolean */
  int in_v;      /* Inside <v> element */
  int in_t;      /* Inside <t> element (for inline strings) */
  int in_is;     /* Inside <is> element (inline string container) */
  char *text;    /* Accumulated text */
  int text_len;  /* Length of accumulated text */
  int text_cap;  /* Capacity of text buffer */
} WorksheetParser;

static void ws_init(Worksheet *ws) { memset(ws, 0, sizeof(*ws)); }

static void ws_free(Worksheet *ws) {
  for (int i = 0; i < ws->count; i++) {
    for (int j = 0; j < ws->rows[i].count; j++) {
      free(ws->rows[i].cells[j].value);
    }
    free(ws->rows[i].cells);
  }
  free(ws->rows);
  memset(ws, 0, sizeof(*ws));
}

static Row *ws_get_row(Worksheet *ws, int row_num) {
  /* Ensure we have enough rows (row_num is 1-based) */
  while (ws->count < row_num) {
    if (ws->count >= ws->capacity) {
      int new_cap = ws->capacity ? ws->capacity * 2 : 64;
      Row *new_rows = realloc(ws->rows, new_cap * sizeof(Row));
      if (!new_rows)
        return NULL;
      ws->rows = new_rows;
      ws->capacity = new_cap;
    }
    memset(&ws->rows[ws->count], 0, sizeof(Row));
    ws->count++;
  }
  return &ws->rows[row_num - 1];
}

static void ws_set_cell(Worksheet *ws, int row_num, int col_num,
                        const char *value) {
  Row *row = ws_get_row(ws, row_num);
  if (!row)
    return;

  /* Ensure we have enough columns (col_num is 1-based) */
  while (row->count < col_num) {
    if (row->count >= row->capacity) {
      int new_cap = row->capacity ? row->capacity * 2 : 16;
      CellValue *new_cells = realloc(row->cells, new_cap * sizeof(CellValue));
      if (!new_cells)
        return;
      row->cells = new_cells;
      row->capacity = new_cap;
    }
    row->cells[row->count].value = NULL;
    row->cells[row->count].is_null = 1;
    row->count++;
  }

  /* Set the cell value */
  free(row->cells[col_num - 1].value);
  row->cells[col_num - 1].value = value ? strdup(value) : NULL;
  row->cells[col_num - 1].is_null = (value == NULL);

  /* Track max column */
  if (col_num > ws->max_col) {
    ws->max_col = col_num;
  }
}

static void wsp_init(WorksheetParser *wsp, Worksheet *ws, SharedStrings *ss) {
  memset(wsp, 0, sizeof(*wsp));
  wsp->ws = ws;
  wsp->ss = ss;
  wsp->cur_type = 'n'; /* Default to number */
}

static void wsp_free(WorksheetParser *wsp) { free(wsp->text); }

static void wsp_append_text(WorksheetParser *wsp, const char *s, int len) {
  if (wsp->text_len + len >= wsp->text_cap) {
    int new_cap = wsp->text_cap ? wsp->text_cap * 2 : 256;
    while (new_cap <= wsp->text_len + len)
      new_cap *= 2;
    char *new_text = realloc(wsp->text, new_cap);
    if (!new_text)
      return;
    wsp->text = new_text;
    wsp->text_cap = new_cap;
  }
  memcpy(wsp->text + wsp->text_len, s, len);
  wsp->text_len += len;
  wsp->text[wsp->text_len] = '\0';
}

static void XMLCALL ws_start_element(void *userData, const XML_Char *name,
                                     const XML_Char **atts) {
  WorksheetParser *wsp = (WorksheetParser *)userData;

  if (strcmp(name, "c") == 0) {
    /* Cell element */
    wsp->cur_type = 'n'; /* Default to number */
    wsp->cur_row = 0;
    wsp->cur_col = 0;

    for (int i = 0; atts[i]; i += 2) {
      if (strcmp(atts[i], "r") == 0) {
        /* Parse cell reference */
        wsp->cur_col = parse_cell_ref(atts[i + 1], &wsp->cur_row);
      } else if (strcmp(atts[i], "t") == 0) {
        /* Cell type */
        if (strcmp(atts[i + 1], "s") == 0) {
          wsp->cur_type = 's'; /* Shared string */
        } else if (strcmp(atts[i + 1], "inlineStr") == 0) {
          wsp->cur_type = 'i'; /* Inline string */
        } else if (strcmp(atts[i + 1], "b") == 0) {
          wsp->cur_type = 'b'; /* Boolean */
        } else if (strcmp(atts[i + 1], "str") == 0) {
          wsp->cur_type = 'f'; /* Formula string result */
        }
      }
    }

    /* Reset text buffer */
    wsp->text_len = 0;
    if (wsp->text)
      wsp->text[0] = '\0';
  } else if (strcmp(name, "v") == 0) {
    wsp->in_v = 1;
    wsp->text_len = 0;
    if (wsp->text)
      wsp->text[0] = '\0';
  } else if (strcmp(name, "is") == 0) {
    wsp->in_is = 1;
  } else if (strcmp(name, "t") == 0 && wsp->in_is) {
    wsp->in_t = 1;
    wsp->text_len = 0;
    if (wsp->text)
      wsp->text[0] = '\0';
  }
}

static void XMLCALL ws_end_element(void *userData, const XML_Char *name) {
  WorksheetParser *wsp = (WorksheetParser *)userData;

  if (strcmp(name, "c") == 0) {
    /* End of cell - store the value */
    if (wsp->cur_row > 0 && wsp->cur_col > 0) {
      const char *value = NULL;

      if (wsp->cur_type == 's' && wsp->text && wsp->ss) {
        /* Shared string - look up by index */
        int idx = atoi(wsp->text);
        if (idx >= 0 && idx < wsp->ss->count) {
          value = wsp->ss->strings[idx];
        }
      } else if (wsp->cur_type == 'i') {
        /* Inline string - use accumulated text */
        value = wsp->text;
      } else if (wsp->text && wsp->text_len > 0) {
        /* Number or other - use as-is */
        value = wsp->text;
      }

      ws_set_cell(wsp->ws, wsp->cur_row, wsp->cur_col, value);
    }
  } else if (strcmp(name, "v") == 0) {
    wsp->in_v = 0;
  } else if (strcmp(name, "is") == 0) {
    wsp->in_is = 0;
  } else if (strcmp(name, "t") == 0 && wsp->in_is) {
    wsp->in_t = 0;
  }
}

static void XMLCALL ws_char_data(void *userData, const XML_Char *s, int len) {
  WorksheetParser *wsp = (WorksheetParser *)userData;

  if (wsp->in_v || wsp->in_t) {
    wsp_append_text(wsp, s, len);
  }
}

static int parse_worksheet(const char *xml, int xml_len, SharedStrings *ss,
                           Worksheet *ws) {
  XML_Parser parser = XML_ParserCreate(NULL);
  if (!parser)
    return -1;

  ws_init(ws);

  WorksheetParser wsp;
  wsp_init(&wsp, ws, ss);

  XML_SetUserData(parser, &wsp);
  XML_SetElementHandler(parser, ws_start_element, ws_end_element);
  XML_SetCharacterDataHandler(parser, ws_char_data);

  int result = 0;
  if (XML_Parse(parser, xml, xml_len, 1) == XML_STATUS_ERROR) {
    result = -1;
  }

  wsp_free(&wsp);
  XML_ParserFree(parser);

  if (result != 0) {
    ws_free(ws);
  }

  return result;
}

/*
** ============================================================================
** Table Name Escaping
** ============================================================================
*/

/*
** Escape a sheet name to make it a valid SQLite identifier.
** Replaces problematic characters and wraps in quotes if necessary.
*/
static char *escape_identifier(const char *name) {
  if (!name || !*name) {
    return strdup("\"unnamed\"");
  }

  /* Calculate needed size (worst case: every char needs escaping) */
  int len = (int)strlen(name);
  char *escaped = malloc(len * 2 + 3); /* Extra for quotes and null */
  if (!escaped)
    return NULL;

  char *p = escaped;
  *p++ = '"';

  for (const char *s = name; *s; s++) {
    if (*s == '"') {
      *p++ = '"'; /* Double the quote */
    }
    *p++ = *s;
  }

  *p++ = '"';
  *p = '\0';

  return escaped;
}

/*
** ============================================================================
** Main Import Function
** ============================================================================
*/

/*
** Read a file from the XLSX archive using the zipfile extension.
*/
static int read_zip_entry(sqlite3 *db, const char *xlsx_path,
                          const char *entry_name, char **data, int *data_len) {
  char *sql = sqlite3_mprintf("SELECT data FROM zipfile(%Q) WHERE name = %Q",
                              xlsx_path, entry_name);
  if (!sql)
    return SQLITE_NOMEM;

  sqlite3_stmt *stmt = NULL;
  int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
  sqlite3_free(sql);

  if (rc != SQLITE_OK) {
    return rc;
  }

  rc = sqlite3_step(stmt);
  if (rc == SQLITE_ROW) {
    const void *blob = sqlite3_column_blob(stmt, 0);
    int blob_len = sqlite3_column_bytes(stmt, 0);

    *data = malloc(blob_len + 1);
    if (*data) {
      memcpy(*data, blob, blob_len);
      (*data)[blob_len] = '\0';
      *data_len = blob_len;
      rc = SQLITE_OK;
    } else {
      rc = SQLITE_NOMEM;
    }
  } else if (rc == SQLITE_DONE) {
    /* Entry not found - not an error, just return empty */
    *data = NULL;
    *data_len = 0;
    rc = SQLITE_OK;
  }

  sqlite3_finalize(stmt);
  return rc;
}

/*
** Create a table from a worksheet.
*/
static int create_table_from_worksheet(sqlite3 *db, const char *table_name,
                                       Worksheet *ws, char **pzErrMsg) {
  if (ws->count == 0 || ws->max_col == 0) {
    /* Empty worksheet */
    return SQLITE_OK;
  }

  /* Build column names from first row */
  Row *first_row = &ws->rows[0];

  /* Start building CREATE TABLE statement */
  sqlite3_str *sql = sqlite3_str_new(db);

  char *escaped_table = escape_identifier(table_name);
  sqlite3_str_appendf(sql, "CREATE TABLE IF NOT EXISTS %s (", escaped_table);
  free(escaped_table);

  for (int col = 0; col < ws->max_col; col++) {
    if (col > 0)
      sqlite3_str_appendall(sql, ", ");

    const char *col_name = NULL;
    if (col < first_row->count && first_row->cells[col].value) {
      col_name = first_row->cells[col].value;
    }

    if (col_name && *col_name) {
      char *escaped_col = escape_identifier(col_name);
      sqlite3_str_appendall(sql, escaped_col);
      free(escaped_col);
    } else {
      /* Generate default column name */
      sqlite3_str_appendf(sql, "\"col%d\"", col + 1);
    }
  }

  sqlite3_str_appendall(sql, ")");

  char *create_sql = sqlite3_str_finish(sql);
  if (!create_sql) {
    return SQLITE_NOMEM;
  }

  int rc = sqlite3_exec(db, create_sql, NULL, NULL, pzErrMsg);
  sqlite3_free(create_sql);

  if (rc != SQLITE_OK) {
    return rc;
  }

  /* Build INSERT statement */
  sql = sqlite3_str_new(db);
  escaped_table = escape_identifier(table_name);
  sqlite3_str_appendf(sql, "INSERT INTO %s VALUES (", escaped_table);
  free(escaped_table);

  for (int col = 0; col < ws->max_col; col++) {
    if (col > 0)
      sqlite3_str_appendall(sql, ", ");
    sqlite3_str_appendall(sql, "?");
  }
  sqlite3_str_appendall(sql, ")");

  char *insert_sql = sqlite3_str_finish(sql);
  if (!insert_sql) {
    return SQLITE_NOMEM;
  }

  sqlite3_stmt *stmt = NULL;
  rc = sqlite3_prepare_v2(db, insert_sql, -1, &stmt, NULL);
  sqlite3_free(insert_sql);

  if (rc != SQLITE_OK) {
    return rc;
  }

  /* Insert data rows (skip first row which is headers) */
  for (int row_idx = 1; row_idx < ws->count; row_idx++) {
    Row *row = &ws->rows[row_idx];

    /* Check if row has any data */
    int has_data = 0;
    for (int col = 0; col < row->count; col++) {
      if (!row->cells[col].is_null) {
        has_data = 1;
        break;
      }
    }
    if (!has_data && row->count == 0)
      continue;

    sqlite3_reset(stmt);

    for (int col = 0; col < ws->max_col; col++) {
      if (col < row->count && !row->cells[col].is_null &&
          row->cells[col].value) {
        sqlite3_bind_text(stmt, col + 1, row->cells[col].value, -1,
                          SQLITE_TRANSIENT);
      } else {
        sqlite3_bind_null(stmt, col + 1);
      }
    }

    rc = sqlite3_step(stmt);
    if (rc != SQLITE_DONE) {
      sqlite3_finalize(stmt);
      return rc;
    }
  }

  sqlite3_finalize(stmt);
  return SQLITE_OK;
}

/*
** xlsx_import(filename) - Import all sheets from an XLSX file as tables.
*/
static void xlsx_import_func(sqlite3_context *ctx, int argc,
                             sqlite3_value **argv) {
  if (argc < 1) {
    sqlite3_result_error(ctx, "xlsx_import requires a filename argument", -1);
    return;
  }

  const char *filename = (const char *)sqlite3_value_text(argv[0]);
  if (!filename) {
    sqlite3_result_error(ctx, "Invalid filename", -1);
    return;
  }

  sqlite3 *db = sqlite3_context_db_handle(ctx);
  int rc;
  char *errmsg = NULL;

  /* Read shared strings */
  SharedStrings ss;
  ss_init(&ss);

  char *ss_data = NULL;
  int ss_len = 0;
  rc = read_zip_entry(db, filename, "xl/sharedStrings.xml", &ss_data, &ss_len);
  if (rc != SQLITE_OK) {
    sqlite3_result_error(
        ctx, "Failed to read XLSX file (is zipfile extension loaded?)", -1);
    return;
  }

  if (ss_data && ss_len > 0) {
    if (parse_shared_strings(ss_data, ss_len, &ss) != 0) {
      free(ss_data);
      sqlite3_result_error(ctx, "Failed to parse shared strings", -1);
      return;
    }
  }
  free(ss_data);

  /* Read workbook to get sheet names */
  Workbook wb;
  wb_init(&wb);

  char *wb_data = NULL;
  int wb_len = 0;
  rc = read_zip_entry(db, filename, "xl/workbook.xml", &wb_data, &wb_len);
  if (rc != SQLITE_OK || !wb_data) {
    ss_free(&ss);
    sqlite3_result_error(ctx, "Failed to read workbook", -1);
    return;
  }

  if (parse_workbook(wb_data, wb_len, &wb) != 0) {
    free(wb_data);
    ss_free(&ss);
    sqlite3_result_error(ctx, "Failed to parse workbook", -1);
    return;
  }
  free(wb_data);

  /* Process each sheet */
  int tables_created = 0;
  for (int i = 0; i < wb.count; i++) {
    char sheet_path[64];
    snprintf(sheet_path, sizeof(sheet_path), "xl/worksheets/sheet%d.xml",
             i + 1);

    char *sheet_data = NULL;
    int sheet_len = 0;
    rc = read_zip_entry(db, filename, sheet_path, &sheet_data, &sheet_len);

    if (rc != SQLITE_OK || !sheet_data || sheet_len == 0) {
      free(sheet_data);
      continue;
    }

    Worksheet ws;
    if (parse_worksheet(sheet_data, sheet_len, &ss, &ws) != 0) {
      free(sheet_data);
      continue;
    }
    free(sheet_data);

    rc = create_table_from_worksheet(db, wb.sheets[i].name, &ws, &errmsg);
    ws_free(&ws);

    if (rc != SQLITE_OK) {
      wb_free(&wb);
      ss_free(&ss);
      if (errmsg) {
        sqlite3_result_error(ctx, errmsg, -1);
        sqlite3_free(errmsg);
      } else {
        sqlite3_result_error(ctx, "Failed to create table", -1);
      }
      return;
    }

    tables_created++;
  }

  wb_free(&wb);
  ss_free(&ss);

  sqlite3_result_int(ctx, tables_created);
}

/*
** xlsx_import_version() - Return the version string.
*/
static void xlsx_import_version_func(sqlite3_context *ctx, int argc,
                                     sqlite3_value **argv) {
  (void)argc;
  (void)argv;
  sqlite3_result_text(ctx, "2025-12-30 Claude Opus 4.5 (Thinking)", -1,
                      SQLITE_STATIC);
}

/*
** ============================================================================
** Extension Entry Point
** ============================================================================
*/

#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsximport_init(sqlite3 *db, char **pzErrMsg,
                            const sqlite3_api_routines *pApi) {
  SQLITE_EXTENSION_INIT2(pApi);
  (void)pzErrMsg;

  int rc = sqlite3_create_function(db, "xlsx_import", 1,
                                   SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL,
                                   xlsx_import_func, NULL, NULL);
  if (rc != SQLITE_OK)
    return rc;

  rc = sqlite3_create_function(db, "xlsx_import_version", 0,
                               SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL,
                               xlsx_import_version_func, NULL, NULL);

  return rc;
}
