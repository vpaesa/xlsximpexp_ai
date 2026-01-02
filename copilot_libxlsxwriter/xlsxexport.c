/*
User prompts included as comments:

Create the C code for a SQLite extension named xlsxexport using the 
libXLSXwriter library that contains a SQL function named xlsx_export 
that saves multiple tables as a single XLSX spreadsheet, with the 
sheet names equal to the table names, and the sheet headers in bold 
and with autofilter. 
The table names are the arguments 2 to N of the xlsx_export function.
Sanitize the sheet names to conform to Excel restrictions.
Add SQL function xlsx_export_version returning "2025-12-30 Copilot Think Deeper (GPT 5.1?)".
Include as comments the prompts used.

*/

/*
Build notes:
- Requires libxlsxwriter and sqlite3 development headers/libraries.
- Example compile (Linux):
    gcc -fPIC -shared -o xlsxexport.so xlsxexport.c -lxlsxwriter -lsqlite3
- Load into SQLite:
    .load ./xlsxexport
- Usage:
    SELECT xlsx_export('out.xlsx', 'table1', 'table2', 'table3');
    SELECT xlsx_export_version();
*/

#define _GNU_SOURCE
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "sqlite3ext.h"
SQLITE_EXTENSION_INIT1

#include <xlsxwriter.h>

/* Excel sheet name restrictions:
   - Max length 31 characters
   - Cannot contain: : \ / ? * [ ]
   - Cannot begin or end with single quote '
   - Cannot be empty
   - Must be unique within workbook
   We'll sanitize by removing invalid chars, trimming whitespace, truncating to 31 chars,
   replacing empty names with "SheetN", and ensuring uniqueness by appending suffixes.
*/

/* Sanitize a single sheet name. Caller must free returned string. */
static char *sanitize_sheet_name(const char *name_in, int index) {
    if (!name_in) name_in = "";
    size_t len = strlen(name_in);
    /* Copy and remove invalid characters */
    char *tmp = malloc(len + 1);
    if (!tmp) return NULL;
    size_t w = 0;
    for (size_t i = 0; i < len; ++i) {
        char c = name_in[i];
        /* Skip invalid characters */
        if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']') {
            continue;
        }
        /* Remove control characters */
        if ((unsigned char)c < 0x20) continue;
        tmp[w++] = c;
    }
    tmp[w] = '\0';

    /* Trim leading/trailing whitespace */
    char *start = tmp;
    while (*start && isspace((unsigned char)*start)) start++;
    char *end = tmp + strlen(tmp);
    while (end > start && isspace((unsigned char)*(end - 1))) end--;
    size_t out_len = (size_t)(end - start);

    /* Remove leading or trailing single quote */
    while (out_len > 0 && *start == '\'') { start++; out_len--; }
    while (out_len > 0 && start[out_len - 1] == '\'') out_len--;

    /* Truncate to 31 characters */
    const size_t MAX_SHEET = 31;
    size_t final_len = out_len > MAX_SHEET ? MAX_SHEET : out_len;

    char *out = malloc(final_len + 32); /* extra for suffix if needed */
    if (!out) { free(tmp); return NULL; }
    memcpy(out, start, final_len);
    out[final_len] = '\0';

    /* If empty, generate default name */
    if (final_len == 0) {
        snprintf(out, final_len + 32, "Sheet%d", index + 1);
    }

    free(tmp);
    return out;
}

/* Ensure uniqueness by appending (n) if necessary. Returns a newly allocated string. */
static char *ensure_unique_name(char **existing, int existing_count, const char *candidate) {
    for (int i = 0; i < existing_count; ++i) {
        if (existing[i] && strcmp(existing[i], candidate) == 0) {
            int suffix = 1;
            char buf[64];
            while (1) {
                snprintf(buf, sizeof(buf), "%s(%d)", candidate, suffix);
                int found = 0;
                for (int j = 0; j < existing_count; ++j) {
                    if (existing[j] && strcmp(existing[j], buf) == 0) { found = 1; break; }
                }
                if (!found) {
                    return strdup(buf);
                }
                suffix++;
            }
        }
    }
    return strdup(candidate);
}

/* Helper: list all user tables (not used in this variant because table names come from args),
   but kept for completeness. Returns array of malloc'd strings and count via out_count. Caller frees.
*/
static char **list_all_tables(sqlite3 *db, int *out_count, char **err_msg) {
    const char *sql = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name;";
    sqlite3_stmt *stmt = NULL;
    int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    if (rc != SQLITE_OK) {
        if (err_msg) *err_msg = sqlite3_mprintf("Failed to prepare sqlite_master query: %s", sqlite3_errmsg(db));
        return NULL;
    }
    int capacity = 16;
    char **arr = malloc(sizeof(char*) * capacity);
    int count = 0;
    while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
        const unsigned char *name = sqlite3_column_text(stmt, 0);
        if (name) {
            if (count >= capacity) {
                capacity *= 2;
                arr = realloc(arr, sizeof(char*) * capacity);
            }
            arr[count++] = strdup((const char*)name);
        }
    }
    if (rc != SQLITE_DONE) {
        if (err_msg) *err_msg = sqlite3_mprintf("Error iterating sqlite_master: %s", sqlite3_errmsg(db));
        for (int i = 0; i < count; ++i) free(arr[i]);
        free(arr);
        sqlite3_finalize(stmt);
        return NULL;
    }
    sqlite3_finalize(stmt);
    *out_count = count;
    return arr;
}

/* Main function: xlsx_export(filename TEXT, table1 TEXT, table2 TEXT, ...)
   - filename: output XLSX file path (required)
   - table names: one or more table names (arguments 2..N). Each must be non-NULL.
   Returns INTEGER 0 on success, or raises an SQLite error on failure.
*/
static void xlsx_export_func(sqlite3_context *context, int argc, sqlite3_value **argv) {
    sqlite3 *db = sqlite3_context_db_handle(context);

    if (argc < 2) {
        sqlite3_result_error(context, "xlsx_export: requires at least filename and one table name", -1);
        return;
    }

    if (sqlite3_value_type(argv[0]) == SQLITE_NULL) {
        sqlite3_result_error(context, "xlsx_export: filename must not be NULL", -1);
        return;
    }
    const char *filename = (const char*)sqlite3_value_text(argv[0]);

    /* Collect table names from argv[1]..argv[argc-1] */
    int table_count = argc - 1;
    char **tables = malloc(sizeof(char*) * table_count);
    if (!tables) {
        sqlite3_result_error(context, "xlsx_export: memory allocation failed", -1);
        return;
    }
    for (int i = 0; i < table_count; ++i) {
        sqlite3_value *v = argv[i + 1];
        if (!v || sqlite3_value_type(v) == SQLITE_NULL) {
            /* free allocated names so far */
            for (int j = 0; j < i; ++j) free(tables[j]);
            free(tables);
            sqlite3_result_error(context, "xlsx_export: table names (arguments 2..N) must not be NULL", -1);
            return;
        }
        const char *tname = (const char*)sqlite3_value_text(v);
        tables[i] = strdup(tname ? tname : "");
        if (!tables[i]) {
            for (int j = 0; j < i; ++j) free(tables[j]);
            free(tables);
            sqlite3_result_error(context, "xlsx_export: memory allocation failed", -1);
            return;
        }
    }

    /* Sanitize and ensure unique sheet names */
    char **sheet_names = malloc(sizeof(char*) * table_count);
    if (!sheet_names) {
        for (int i = 0; i < table_count; ++i) free(tables[i]);
        free(tables);
        sqlite3_result_error(context, "xlsx_export: memory allocation failed", -1);
        return;
    }
    for (int i = 0; i < table_count; ++i) sheet_names[i] = NULL;
    for (int i = 0; i < table_count; ++i) {
        char *san = sanitize_sheet_name(tables[i], i);
        if (!san) {
            sqlite3_result_error(context, "xlsx_export: memory allocation failed", -1);
            for (int j = 0; j < table_count; ++j) { free(tables[j]); if (sheet_names[j]) free(sheet_names[j]); }
            free(tables); free(sheet_names);
            return;
        }
        char *unique = ensure_unique_name(sheet_names, i, san);
        free(san);
        sheet_names[i] = unique;
    }

    /* Create workbook */
    lxw_workbook  *workbook = workbook_new(filename);
    if (!workbook) {
        sqlite3_result_error(context, "xlsx_export: failed to create workbook", -1);
        for (int i = 0; i < table_count; ++i) { free(tables[i]); free(sheet_names[i]); }
        free(tables); free(sheet_names);
        return;
    }

    /* Create bold format for headers */
    lxw_format *fmt_bold = workbook_add_format(workbook);
    format_set_bold(fmt_bold);

    /* Iterate tables and write sheets */
    for (int t = 0; t < table_count; ++t) {
        const char *tbl = tables[t];
        const char *sheet = sheet_names[t];

        lxw_worksheet *worksheet = workbook_add_worksheet(workbook, sheet);
        if (!worksheet) {
            workbook_close(workbook);
            sqlite3_result_error(context, "xlsx_export: failed to add worksheet", -1);
            for (int i = 0; i < table_count; ++i) { free(tables[i]); free(sheet_names[i]); }
            free(tables); free(sheet_names);
            return;
        }

        /* Prepare SELECT * FROM "table" */
        char *sql = NULL;
        if (asprintf(&sql, "SELECT * FROM \"%s\";", tbl) == -1) sql = NULL;
        if (!sql) {
            workbook_close(workbook);
            sqlite3_result_error(context, "xlsx_export: memory allocation failed", -1);
            for (int i = 0; i < table_count; ++i) { free(tables[i]); free(sheet_names[i]); }
            free(tables); free(sheet_names);
            return;
        }

        sqlite3_stmt *stmt = NULL;
        int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
        free(sql);
        if (rc != SQLITE_OK) {
            workbook_close(workbook);
            sqlite3_result_error(context, sqlite3_errmsg(db), -1);
            for (int i = 0; i < table_count; ++i) { free(tables[i]); free(sheet_names[i]); }
            free(tables); free(sheet_names);
            return;
        }

        /* Write header row */
        int col_count = sqlite3_column_count(stmt);
        for (int c = 0; c < col_count; ++c) {
            const char *colname = sqlite3_column_name(stmt, c);
            if (!colname) colname = "";
            worksheet_write_string(worksheet, 0, c, colname, fmt_bold);
        }

        /* Write data rows */
        int row = 1;
        while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
            for (int c = 0; c < col_count; ++c) {
                int col_type = sqlite3_column_type(stmt, c);
                if (col_type == SQLITE_INTEGER) {
                    sqlite3_int64 v = sqlite3_column_int64(stmt, c);
                    worksheet_write_number(worksheet, row, c, (double)v, NULL);
                } else if (col_type == SQLITE_FLOAT) {
                    double v = sqlite3_column_double(stmt, c);
                    worksheet_write_number(worksheet, row, c, v, NULL);
                } else if (col_type == SQLITE_NULL) {
                    worksheet_write_blank(worksheet, row, c, NULL);
                } else {
                    const unsigned char *text = sqlite3_column_text(stmt, c);
                    if (text) {
                        worksheet_write_string(worksheet, row, c, (const char*)text, NULL);
                    } else {
                        worksheet_write_blank(worksheet, row, c, NULL);
                    }
                }
            }
            row++;
        }
        if (rc != SQLITE_DONE) {
            sqlite3_finalize(stmt);
            workbook_close(workbook);
            sqlite3_result_error(context, sqlite3_errmsg(db), -1);
            for (int i = 0; i < table_count; ++i) { free(tables[i]); free(sheet_names[i]); }
            free(tables); free(sheet_names);
            return;
        }
        sqlite3_finalize(stmt);

        /* Apply autofilter from header row (0) to last data row (row-1) and last column (col_count-1) */
        if (col_count > 0 && row > 1) {
            worksheet_autofilter(worksheet, 0, 0, row - 1, col_count - 1);
        }
    }

    /* Close workbook */
    lxw_error err = workbook_close(workbook);
    if (err != LXW_NO_ERROR) {
        sqlite3_result_error(context, "xlsx_export: failed to write workbook file", -1);
        for (int i = 0; i < table_count; ++i) { free(tables[i]); free(sheet_names[i]); }
        free(tables); free(sheet_names);
        return;
    }

    /* Cleanup */
    for (int i = 0; i < table_count; ++i) {
        free(tables[i]);
        free(sheet_names[i]);
    }
    free(tables);
    free(sheet_names);

    /* Return success (0) */
    sqlite3_result_int(context, 0);
}

/* xlsx_export_version() -> TEXT */
static void xlsx_export_version_func(sqlite3_context *context, int argc, sqlite3_value **argv) {
    (void)argc; (void)argv;
    sqlite3_result_text(context, "2025-12-30 Copilot Think Deeper (GPT 5.1?)", -1, SQLITE_STATIC);
}

/* Extension entry point */
#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsxexport_init(sqlite3 *db, char **pzErrMsg, const sqlite3_api_routines *pApi) {
    (void)pzErrMsg;
    SQLITE_EXTENSION_INIT2(pApi);

    int rc = SQLITE_OK;
    /* Register xlsx_export with variable arguments (-1). Arguments: filename, table1, table2, ... */
    rc = sqlite3_create_function(db, "xlsx_export", -1, SQLITE_UTF8, NULL, xlsx_export_func, NULL, NULL);
    if (rc != SQLITE_OK) {
        if (pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_export: %s", sqlite3_errmsg(db));
        return rc;
    }
    rc = sqlite3_create_function(db, "xlsx_export_version", 0, SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL, xlsx_export_version_func, NULL, NULL);
    if (rc != SQLITE_OK) {
        if (pzErrMsg) *pzErrMsg = sqlite3_mprintf("Failed to register xlsx_export_version: %s", sqlite3_errmsg(db));
        return rc;
    }
    return SQLITE_OK;
}

