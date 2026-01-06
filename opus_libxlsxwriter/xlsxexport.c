/*
PROMPTS USED:
Create the C code for a SQLite extension named xlsxexport using the 
libXLSXwriter library that contains a SQL function named xlsx_export 
that saves multiple tables as a single XLSX spreadsheet, with the 
sheet names equal to the table names, and the sheet headers in bold 
and with autofilter. If invoked with only one parameter then exports all the tables in the schema
Include as comments the commands to cross compile 
on Cygwin statically the extension and libXLSXwriter. Include as 
comments the prompts used.

CROSS-COMPILATION ON CYGWIN (static build):

1. Build libxlsxwriter statically:
    cd libxlsxwriter
    make USE_SYSTEM_MINIZIP=0 USE_STANDARD_TMPFILE=1 CFLAGS="-fPIC"

2. Compile the extension (adjust paths as needed):
    x86_64-w64-mingw32-gcc -shared -o xlsxexport.dll xlsxexport.c \
        -I/path/to/libxlsxwriter/include \
        -I/path/to/sqlite3 \
        -L/path/to/libxlsxwriter/lib \
        -l:libxlsxwriter.a -lz \
        -static-libgcc -static

3. Alternative single command with inline paths:
    x86_64-w64-mingw32-gcc -shared -fPIC -o xlsxexport.dll xlsxexport.c \
        -I./libxlsxwriter/include \
        -L./libxlsxwriter/lib \
        -Wl,-Bstatic -lxlsxwriter -lz \
        -Wl,-Bdynamic \
        -DSQLITE_CORE

USAGE IN SQLITE:
    .load xlsxexport
    SELECT xlsx_export('output.xlsx');  -- Export all tables in the schema
    SELECT xlsx_export('output.xlsx', 'table1', 'table2', 'table3');
    -- or with a single table:
    SELECT xlsx_export('output.xlsx', 'mytable');
*/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "sqlite3ext.h"
#include "xlsxwriter.h"

SQLITE_EXTENSION_INIT1

/*
** Helper function to export a single table to a worksheet.
** Returns 0 on success, non-zero on error.
*/
static int export_table_to_sheet(
    sqlite3 *db,
    lxw_workbook *workbook,
    const char *table_name,
    char **err_msg
) {
    lxw_worksheet *worksheet;
    sqlite3_stmt *stmt = NULL;
    char *sql = NULL;
    int rc;
    int col_count;
    lxw_row_t row_num;
    lxw_col_t col_num;
    lxw_format *header_format;
    
    /* Create a worksheet with the table name */
    worksheet = workbook_add_worksheet(workbook, table_name);
    if (!worksheet) {
        *err_msg = sqlite3_mprintf("Failed to create worksheet for table '%s'", table_name);
        return 1;
    }
    
    /* Create a bold format for headers */
    header_format = workbook_add_format(workbook);
    format_set_bold(header_format);
    
    /* Build the SELECT query */
    sql = sqlite3_mprintf("SELECT * FROM \"%w\"", table_name);
    if (!sql) {
        *err_msg = sqlite3_mprintf("Out of memory");
        return 1;
    }
    
    /* Prepare the statement */
    rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    sqlite3_free(sql);
    
    if (rc != SQLITE_OK) {
        *err_msg = sqlite3_mprintf("Failed to prepare query for table '%s': %s",
                                    table_name, sqlite3_errmsg(db));
        return 1;
    }
    
    /* Get column count */
    col_count = sqlite3_column_count(stmt);
    
    /* Write header row with column names */
    for (col_num = 0; col_num < col_count; col_num++) {
        const char *col_name = sqlite3_column_name(stmt, col_num);
        worksheet_write_string(worksheet, 0, col_num, col_name, header_format);
    }
    
    /* Set autofilter on the header row */
    if (col_count > 0) {
        worksheet_autofilter(worksheet, 0, 0, 0, col_count - 1);
    }
    
    /* Write data rows */
    row_num = 1;
    while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
        for (col_num = 0; col_num < col_count; col_num++) {
            int col_type = sqlite3_column_type(stmt, col_num);
            
            switch (col_type) {
                case SQLITE_INTEGER:
                    worksheet_write_number(worksheet, row_num, col_num,
                                           (double)sqlite3_column_int64(stmt, col_num), NULL);
                    break;
                    
                case SQLITE_FLOAT:
                    worksheet_write_number(worksheet, row_num, col_num,
                                           sqlite3_column_double(stmt, col_num), NULL);
                    break;
                    
                case SQLITE_TEXT:
                    worksheet_write_string(worksheet, row_num, col_num,
                                           (const char *)sqlite3_column_text(stmt, col_num), NULL);
                    break;
                    
                case SQLITE_BLOB:
                    /* Write BLOBs as hex strings */
                    {
                        int blob_size = sqlite3_column_bytes(stmt, col_num);
                        const unsigned char *blob = sqlite3_column_blob(stmt, col_num);
                        char *hex = sqlite3_malloc(blob_size * 2 + 1);
                        if (hex) {
                            int i;
                            for (i = 0; i < blob_size; i++) {
                                sprintf(hex + i * 2, "%02X", blob[i]);
                            }
                            hex[blob_size * 2] = '\0';
                            worksheet_write_string(worksheet, row_num, col_num, hex, NULL);
                            sqlite3_free(hex);
                        }
                    }
                    break;
                    
                case SQLITE_NULL:
                default:
                    /* Leave cell empty for NULL values */
                    break;
            }
        }
        row_num++;
    }
    
    sqlite3_finalize(stmt);
    
    if (rc != SQLITE_DONE) {
        *err_msg = sqlite3_mprintf("Error reading table '%s': %s",
                                    table_name, sqlite3_errmsg(db));
        return 1;
    }
    
    /* Extend autofilter to cover all data rows */
    if (col_count > 0 && row_num > 1) {
        worksheet_autofilter(worksheet, 0, 0, row_num - 1, col_count - 1);
    }
    
    return 0;
}

/*
** SQL function: xlsx_export(filename [, table1, table2, ...])
**
** Exports tables to an XLSX file.
** The first argument is the output filename.
** If no table names are provided, all tables in the schema are exported.
** If table names are provided, only those tables are exported.
** Each table becomes a separate worksheet with the table name as sheet name.
** Headers are bold and have autofilter enabled.
**
** Returns the filename on success, or raises an error on failure.
*/
static void xlsx_export_func(
    sqlite3_context *context,
    int argc,
    sqlite3_value **argv
) {
    sqlite3 *db;
    lxw_workbook *workbook;
    lxw_workbook_options options = {.constant_memory = LXW_FALSE};
    const char *filename;
    char *err_msg = NULL;
    int i;
    lxw_error error;
    
    /* Need at least the filename */
    if (argc < 1) {
        sqlite3_result_error(context, 
            "xlsx_export requires at least 1 argument: filename", -1);
        return;
    }
    
    /* Get the output filename */
    if (sqlite3_value_type(argv[0]) != SQLITE_TEXT) {
        sqlite3_result_error(context, "First argument must be the output filename", -1);
        return;
    }
    filename = (const char *)sqlite3_value_text(argv[0]);
    
    /* Get database connection */
    db = sqlite3_context_db_handle(context);
    
    /* Create the workbook */
    workbook = workbook_new_opt(filename, &options);
    if (!workbook) {
        sqlite3_result_error(context, "Failed to create workbook", -1);
        return;
    }
    
    if (argc == 1) {
        /* No table names provided - export all tables from schema */
        sqlite3_stmt *stmt = NULL;
        int rc;
        
        /* Query sqlite_master for all user tables (exclude sqlite_ internal tables) */
        rc = sqlite3_prepare_v2(db, 
            "SELECT name FROM sqlite_master WHERE type='table' "
            "AND name NOT LIKE 'sqlite_%' ORDER BY rowid", 
            -1, &stmt, NULL);
        
        if (rc != SQLITE_OK) {
            sqlite3_result_error(context, "Failed to query schema for table names", -1);
            workbook_close(workbook);
            return;
        }
        
        while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
            const char *table_name = (const char *)sqlite3_column_text(stmt, 0);
            
            if (export_table_to_sheet(db, workbook, table_name, &err_msg) != 0) {
                sqlite3_result_error(context, err_msg, -1);
                sqlite3_free(err_msg);
                sqlite3_finalize(stmt);
                workbook_close(workbook);
                return;
            }
        }
        
        sqlite3_finalize(stmt);
        
        if (rc != SQLITE_DONE) {
            sqlite3_result_error(context, "Error enumerating tables", -1);
            workbook_close(workbook);
            return;
        }
    } else {
        /* Export specified tables */
        for (i = 1; i < argc; i++) {
            const char *table_name;
            
            if (sqlite3_value_type(argv[i]) != SQLITE_TEXT) {
                sqlite3_result_error(context, "Table names must be strings", -1);
                workbook_close(workbook);
                return;
            }
            
            table_name = (const char *)sqlite3_value_text(argv[i]);
            
            if (export_table_to_sheet(db, workbook, table_name, &err_msg) != 0) {
                sqlite3_result_error(context, err_msg, -1);
                sqlite3_free(err_msg);
                workbook_close(workbook);
                return;
            }
        }
    }
    
    /* Close the workbook and write the file */
    error = workbook_close(workbook);
    if (error != LXW_NO_ERROR) {
        char *msg = sqlite3_mprintf("Error closing workbook: %s", lxw_strerror(error));
        sqlite3_result_error(context, msg, -1);
        sqlite3_free(msg);
        return;
    }
    
    /* Return the filename on success */
    sqlite3_result_text(context, filename, -1, SQLITE_TRANSIENT);
}

/*
** SQL function: xlsx_export_version()
**
** Returns the version string.
*/
static void xlsx_export_version(
    sqlite3_context *context,
    int argc,
    sqlite3_value **argv
) {
    (void)argc;
    (void)argv;
    sqlite3_result_text(context, "2025-12-30 Claude Opus 4.5 (Thinking)", -1, SQLITE_STATIC);
}

/*
** Extension entry point
*/
#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsxexport_init(
    sqlite3 *db,
    char **pzErrMsg,
    const sqlite3_api_routines *pApi
) {
    int rc;
    SQLITE_EXTENSION_INIT2(pApi);
    
    (void)pzErrMsg;  /* Unused parameter */
    
    /* Register the xlsx_export function */
    rc = sqlite3_create_function(
        db,
        "xlsx_export",      /* Function name */
        -1,                 /* Variable number of arguments */
        SQLITE_UTF8 | SQLITE_DETERMINISTIC,
        NULL,               /* User data */
        xlsx_export_func,   /* Function implementation */
        NULL,               /* Step (for aggregate functions) */
        NULL                /* Final (for aggregate functions) */
    );

    if (rc == SQLITE_OK) {
        rc = sqlite3_create_function(
            db,
            "xlsx_export_version",
            0,
            SQLITE_UTF8 | SQLITE_DETERMINISTIC,
            NULL,
            xlsx_export_version,
            NULL,
            NULL
        );
    }
    
    return rc;
}
