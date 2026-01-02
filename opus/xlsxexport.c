/*
** PROMPTS USED:
**
** "create the C code for a SQLite extension named xlsxexport that 
** contains a SQL function named xlsx_export that saves multiple tables 
** as a single XLSX spreadsheet, with the sheet names equal to the table 
** names, and the sheet headers in bold and with autofilter. The XLSX file 
** is a ZIP archive with XML files inside. Use the zipfile extension to 
** create this ZIP archive. Do not use any external library to write XML 
** files. Include as comments the prompts used."
**
** "warn if the Excel maximum cell size is exceeded"
**
** "sanitize the sheet names to conform to Excel restrictions"
**
** USAGE IN SQLITE:
**    .load zipfile
**    .load xlsxexport
**    SELECT xlsx_export('output.xlsx', 'table1', 'table2', 'table3');
**    -- or with a single table:
**    SELECT xlsx_export('output.xlsx', 'mytable');
**
** NOTES:
**    - Requires the zipfile extension to be loaded first
**    - Generates XLSX-compliant XML files manually
**    - Headers are bold (using styles.xml) with autofilter enabled
**    - Warns if cell content exceeds Excel's 32,767 character limit
**    - Sheet names are sanitized (max 31 chars, no \ / ? * [ ] :, no "History")
*/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "sqlite3ext.h"

SQLITE_EXTENSION_INIT1

/* Excel maximum cell size (32,767 characters) */
#define EXCEL_MAX_CELL_SIZE 32767

/* Warning tracking structure */
typedef struct ExportWarnings {
    int cells_truncated;
    int first_truncated_row;
    int first_truncated_col;
    const char *first_truncated_table;
} ExportWarnings;

/* String buffer for dynamic string building */
typedef struct StrBuf {
    char *str;
    size_t len;
    size_t cap;
} StrBuf;

static void strbuf_init(StrBuf *sb) {
    sb->str = NULL;
    sb->len = 0;
    sb->cap = 0;
}

static void strbuf_free(StrBuf *sb) {
    sqlite3_free(sb->str);
    sb->str = NULL;
    sb->len = 0;
    sb->cap = 0;
}

static int strbuf_append(StrBuf *sb, const char *s) {
    size_t slen = strlen(s);
    if (sb->len + slen + 1 > sb->cap) {
        size_t newcap = (sb->cap == 0) ? 4096 : sb->cap * 2;
        while (newcap < sb->len + slen + 1) newcap *= 2;
        char *newstr = sqlite3_realloc(sb->str, newcap);
        if (!newstr) return 1;
        sb->str = newstr;
        sb->cap = newcap;
    }
    memcpy(sb->str + sb->len, s, slen + 1);
    sb->len += slen;
    return 0;
}

static int strbuf_appendf(StrBuf *sb, const char *fmt, ...) {
    va_list ap;
    char *s;
    va_start(ap, fmt);
    s = sqlite3_vmprintf(fmt, ap);
    va_end(ap);
    if (!s) return 1;
    int rc = strbuf_append(sb, s);
    sqlite3_free(s);
    return rc;
}

/* Convert column number to Excel column letter (0=A, 1=B, ..., 25=Z, 26=AA, ...) */
static void col_to_letter(int col, char *buf) {
    char temp[16];
    int i = 0;
    col++;  /* 1-based for calculation */
    while (col > 0) {
        col--;
        temp[i++] = 'A' + (col % 26);
        col /= 26;
    }
    /* Reverse */
    int j;
    for (j = 0; j < i; j++) {
        buf[j] = temp[i - 1 - j];
    }
    buf[i] = '\0';
}

/* Escape XML special characters */
static char *xml_escape(const char *s) {
    StrBuf sb;
    strbuf_init(&sb);
    const char *p;
    for (p = s; *p; p++) {
        switch (*p) {
            case '&':  strbuf_append(&sb, "&amp;"); break;
            case '<':  strbuf_append(&sb, "&lt;"); break;
            case '>':  strbuf_append(&sb, "&gt;"); break;
            case '"':  strbuf_append(&sb, "&quot;"); break;
            case '\'': strbuf_append(&sb, "&apos;"); break;
            default:
                if ((unsigned char)*p < 32 && *p != '\t' && *p != '\n' && *p != '\r') {
                    /* Skip invalid XML characters */
                } else {
                    char c[2] = {*p, 0};
                    strbuf_append(&sb, c);
                }
                break;
        }
    }
    return sb.str ? sb.str : sqlite3_mprintf("");
}

/* Excel sheet name maximum length */
#define EXCEL_MAX_SHEET_NAME_LEN 31

/*
** Sanitize a sheet name to conform to Excel restrictions:
** - Maximum 31 characters
** - Cannot contain: \ / ? * [ ] :
** - Cannot be empty
** - Cannot start or end with apostrophe (')
** - Cannot be "History" (reserved by Excel)
** Returns a newly allocated string that must be freed with sqlite3_free().
*/
static char *sanitize_sheet_name(const char *name) {
    char *result;
    int i, j, len;
    
    if (!name || !*name) {
        return sqlite3_mprintf("Sheet");
    }
    
    len = (int)strlen(name);
    if (len > EXCEL_MAX_SHEET_NAME_LEN) {
        len = EXCEL_MAX_SHEET_NAME_LEN;
    }
    
    result = sqlite3_malloc(len + 1);
    if (!result) return NULL;
    
    /* Copy characters, replacing invalid ones with underscore */
    for (i = 0, j = 0; i < len && name[i]; i++) {
        char c = name[i];
        /* Check for invalid characters: \ / ? * [ ] : */
        if (c == '\\' || c == '/' || c == '?' || c == '*' || 
            c == '[' || c == ']' || c == ':') {
            result[j++] = '_';
        } else {
            result[j++] = c;
        }
    }
    result[j] = '\0';
    
    /* Handle leading apostrophe */
    if (result[0] == '\'') {
        result[0] = '_';
    }
    
    /* Handle trailing apostrophe */
    if (j > 0 && result[j - 1] == '\'') {
        result[j - 1] = '_';
    }
    
    /* If result is empty after sanitization, use default name */
    if (result[0] == '\0') {
        sqlite3_free(result);
        return sqlite3_mprintf("Sheet");
    }
    
    /* "History" is a reserved sheet name in Excel */
    if ((j == 7) && 
        (result[0] == 'H' || result[0] == 'h') &&
        (result[1] == 'I' || result[1] == 'i') &&
        (result[2] == 'S' || result[2] == 's') &&
        (result[3] == 'T' || result[3] == 't') &&
        (result[4] == 'O' || result[4] == 'o') &&
        (result[5] == 'R' || result[5] == 'r') &&
        (result[6] == 'Y' || result[6] == 'y')) {
        sqlite3_free(result);
        return sqlite3_mprintf("History_");
    }
    
    return result;
}

/* Generate [Content_Types].xml */
static char *gen_content_types(int sheet_count) {
    StrBuf sb;
    strbuf_init(&sb);
    
    strbuf_append(&sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    
    int i;
    for (i = 1; i <= sheet_count; i++) {
        strbuf_appendf(&sb, 
            "<Override PartName=\"/xl/worksheets/sheet%d.xml\" "
            "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>", i);
    }
    
    strbuf_append(&sb, "</Types>");
    return sb.str;
}

/* Generate _rels/.rels */
static char *gen_rels(void) {
    return sqlite3_mprintf(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>");
}

/* Generate xl/_rels/workbook.xml.rels */
static char *gen_workbook_rels(int sheet_count) {
    StrBuf sb;
    strbuf_init(&sb);
    
    strbuf_append(&sb, 
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
    
    int i;
    for (i = 1; i <= sheet_count; i++) {
        strbuf_appendf(&sb,
            "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>",
            i, i);
    }
    
    strbuf_append(&sb, "</Relationships>");
    return sb.str;
}

/* Generate xl/workbook.xml */
static char *gen_workbook(const char **sheet_names, int sheet_count) {
    StrBuf sb;
    strbuf_init(&sb);
    
    strbuf_append(&sb,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets>");
    
    int i;
    for (i = 0; i < sheet_count; i++) {
        char *sanitized_name = sanitize_sheet_name(sheet_names[i]);
        char *escaped_name = xml_escape(sanitized_name ? sanitized_name : "Sheet");
        strbuf_appendf(&sb, "<sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>",
            escaped_name, i + 1, i + 1);
        sqlite3_free(escaped_name);
        sqlite3_free(sanitized_name);
    }
    
    strbuf_append(&sb, "</sheets></workbook>");
    return sb.str;
}

/* Generate xl/styles.xml with bold style for headers */
static char *gen_styles(void) {
    return sqlite3_mprintf(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"2\">"
        "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
        "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>"
        "</fonts>"
        "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>"
        "<borders count=\"1\"><border/></borders>"
        "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>"
        "<cellXfs count=\"2\">"
        "<xf fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "<xf fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
        "</cellXfs>"
        "</styleSheet>");
}

/* Generate xl/worksheets/sheetN.xml for a table */
static char *gen_worksheet(sqlite3 *db, const char *table_name, char **err_msg, 
                           ExportWarnings *warnings) {
    StrBuf sb;
    strbuf_init(&sb);
    sqlite3_stmt *stmt = NULL;
    char *sql = NULL;
    int rc;
    int col_count;
    int row_num;
    int col;
    char col_letter[16];
    int last_row = 1;
    
    /* Build SELECT query */
    sql = sqlite3_mprintf("SELECT * FROM \"%w\"", table_name);
    if (!sql) {
        *err_msg = sqlite3_mprintf("Out of memory");
        return NULL;
    }
    
    rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    sqlite3_free(sql);
    
    if (rc != SQLITE_OK) {
        *err_msg = sqlite3_mprintf("Failed to prepare query for table '%s': %s",
            table_name, sqlite3_errmsg(db));
        return NULL;
    }
    
    col_count = sqlite3_column_count(stmt);
    
    /* Start worksheet XML */
    strbuf_append(&sb,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetData>");
    
    /* Write header row with column names (style 1 = bold) */
    strbuf_append(&sb, "<row r=\"1\">");
    for (col = 0; col < col_count; col++) {
        const char *col_name = sqlite3_column_name(stmt, col);
        char *escaped = xml_escape(col_name);
        col_to_letter(col, col_letter);
        strbuf_appendf(&sb, "<c r=\"%s1\" t=\"inlineStr\" s=\"1\"><is><t>%s</t></is></c>",
            col_letter, escaped);
        sqlite3_free(escaped);
    }
    strbuf_append(&sb, "</row>");
    
    /* Write data rows */
    row_num = 2;
    while ((rc = sqlite3_step(stmt)) == SQLITE_ROW) {
        strbuf_appendf(&sb, "<row r=\"%d\">", row_num);
        
        for (col = 0; col < col_count; col++) {
            int col_type = sqlite3_column_type(stmt, col);
            col_to_letter(col, col_letter);
            
            switch (col_type) {
                case SQLITE_INTEGER:
                    strbuf_appendf(&sb, "<c r=\"%s%d\"><v>%lld</v></c>",
                        col_letter, row_num, sqlite3_column_int64(stmt, col));
                    break;
                    
                case SQLITE_FLOAT:
                    strbuf_appendf(&sb, "<c r=\"%s%d\"><v>%.15g</v></c>",
                        col_letter, row_num, sqlite3_column_double(stmt, col));
                    break;
                    
                case SQLITE_TEXT: {
                    const char *text = (const char *)sqlite3_column_text(stmt, col);
                    int text_len = sqlite3_column_bytes(stmt, col);
                    char *text_to_escape = NULL;
                    
                    /* Check for Excel cell size limit */
                    if (text_len > EXCEL_MAX_CELL_SIZE) {
                        /* Truncate the text */
                        text_to_escape = sqlite3_malloc(EXCEL_MAX_CELL_SIZE + 1);
                        if (text_to_escape) {
                            memcpy(text_to_escape, text, EXCEL_MAX_CELL_SIZE);
                            text_to_escape[EXCEL_MAX_CELL_SIZE] = '\0';
                        }
                        /* Track warning */
                        warnings->cells_truncated++;
                        if (warnings->cells_truncated == 1) {
                            warnings->first_truncated_row = row_num;
                            warnings->first_truncated_col = col + 1;
                            warnings->first_truncated_table = table_name;
                        }
                    }
                    
                    char *escaped = xml_escape(text_to_escape ? text_to_escape : text);
                    strbuf_appendf(&sb, "<c r=\"%s%d\" t=\"inlineStr\"><is><t>%s</t></is></c>",
                        col_letter, row_num, escaped);
                    sqlite3_free(escaped);
                    sqlite3_free(text_to_escape);
                    break;
                }
                    
                case SQLITE_BLOB: {
                    /* Write BLOBs as hex strings */
                    int blob_size = sqlite3_column_bytes(stmt, col);
                    int hex_len = blob_size * 2;
                    int truncated = 0;
                    
                    /* Check for Excel cell size limit (hex is 2x blob size) */
                    if (hex_len > EXCEL_MAX_CELL_SIZE) {
                        blob_size = EXCEL_MAX_CELL_SIZE / 2;
                        hex_len = blob_size * 2;
                        truncated = 1;
                        /* Track warning */
                        warnings->cells_truncated++;
                        if (warnings->cells_truncated == 1) {
                            warnings->first_truncated_row = row_num;
                            warnings->first_truncated_col = col + 1;
                            warnings->first_truncated_table = table_name;
                        }
                    }
                    
                    const unsigned char *blob = sqlite3_column_blob(stmt, col);
                    char *hex = sqlite3_malloc(hex_len + 1);
                    if (hex) {
                        int i;
                        for (i = 0; i < blob_size; i++) {
                            sprintf(hex + i * 2, "%02X", blob[i]);
                        }
                        hex[hex_len] = '\0';
                        strbuf_appendf(&sb, "<c r=\"%s%d\" t=\"inlineStr\"><is><t>%s</t></is></c>",
                            col_letter, row_num, hex);
                        sqlite3_free(hex);
                    }
                    (void)truncated; /* Used for tracking only */
                    break;
                }
                    
                case SQLITE_NULL:
                default:
                    /* Empty cell for NULL */
                    break;
            }
        }
        
        strbuf_append(&sb, "</row>");
        row_num++;
    }
    last_row = row_num - 1;
    
    sqlite3_finalize(stmt);
    
    if (rc != SQLITE_DONE) {
        *err_msg = sqlite3_mprintf("Error reading table '%s': %s",
            table_name, sqlite3_errmsg(db));
        strbuf_free(&sb);
        return NULL;
    }
    
    strbuf_append(&sb, "</sheetData>");
    
    /* Add autofilter for columns A through last column, rows 1 through last_row */
    if (col_count > 0) {
        char last_col_letter[16];
        col_to_letter(col_count - 1, last_col_letter);
        strbuf_appendf(&sb, "<autoFilter ref=\"A1:%s%d\"/>", last_col_letter, last_row);
    }
    
    strbuf_append(&sb, "</worksheet>");
    
    return sb.str;
}

/*
** SQL function: xlsx_export(filename, table1, table2, ...)
**
** Exports one or more tables to an XLSX file using the zipfile extension.
** Returns the filename on success.
*/
static void xlsx_export_func(
    sqlite3_context *context,
    int argc,
    sqlite3_value **argv
) {
    sqlite3 *db;
    const char *filename;
    char *err_msg = NULL;
    int i;
    int sheet_count;
    const char **sheet_names = NULL;
    char **sheet_contents = NULL;
    char *content_types = NULL;
    char *rels = NULL;
    char *workbook_rels = NULL;
    char *workbook = NULL;
    char *styles = NULL;
    sqlite3_stmt *stmt = NULL;
    char *sql = NULL;
    int rc;
    ExportWarnings warnings = {0, 0, 0, NULL};
    
    /* Need at least filename and one table name */
    if (argc < 2) {
        sqlite3_result_error(context, 
            "xlsx_export requires at least 2 arguments: filename and table name(s)", -1);
        return;
    }
    
    /* Get the output filename */
    if (sqlite3_value_type(argv[0]) != SQLITE_TEXT) {
        sqlite3_result_error(context, "First argument must be the output filename", -1);
        return;
    }
    filename = (const char *)sqlite3_value_text(argv[0]);
    
    db = sqlite3_context_db_handle(context);
    sheet_count = argc - 1;
    
    /* Collect sheet names */
    sheet_names = sqlite3_malloc(sizeof(char *) * sheet_count);
    sheet_contents = sqlite3_malloc(sizeof(char *) * sheet_count);
    if (!sheet_names || !sheet_contents) {
        sqlite3_result_error(context, "Out of memory", -1);
        goto cleanup;
    }
    memset(sheet_contents, 0, sizeof(char *) * sheet_count);
    
    for (i = 0; i < sheet_count; i++) {
        if (sqlite3_value_type(argv[i + 1]) != SQLITE_TEXT) {
            sqlite3_result_error(context, "Table names must be strings", -1);
            goto cleanup;
        }
        sheet_names[i] = (const char *)sqlite3_value_text(argv[i + 1]);
    }
    
    /* Generate all worksheet contents */
    for (i = 0; i < sheet_count; i++) {
        sheet_contents[i] = gen_worksheet(db, sheet_names[i], &err_msg, &warnings);
        if (!sheet_contents[i]) {
            sqlite3_result_error(context, err_msg, -1);
            sqlite3_free(err_msg);
            goto cleanup;
        }
    }
    
    /* Generate other XML files */
    content_types = gen_content_types(sheet_count);
    rels = gen_rels();
    workbook_rels = gen_workbook_rels(sheet_count);
    workbook = gen_workbook(sheet_names, sheet_count);
    styles = gen_styles();
    
    if (!content_types || !rels || !workbook_rels || !workbook || !styles) {
        sqlite3_result_error(context, "Failed to generate XML content", -1);
        goto cleanup;
    }
    
    /* Delete existing file first using writefile if it exists */
    sql = sqlite3_mprintf("SELECT writefile(%Q, zeroblob(0))", filename);
    sqlite3_exec(db, sql, NULL, NULL, NULL);
    sqlite3_free(sql);
    
    /* Now we need to use zipfile() to create the archive */
    /* First, delete the file if it exists */
    sql = sqlite3_mprintf("SELECT writefile(%Q, NULL)", filename);
    sqlite3_exec(db, sql, NULL, NULL, NULL);
    sqlite3_free(sql);
    
    /* Build the complete ZIP using zipfile() Aggregate Function */
    {
        StrBuf insert_sql;
        strbuf_init(&insert_sql);
        strbuf_append(&insert_sql, "WITH contents(name, data) AS (\n");
        
        /* [Content_Types].xml */
        strbuf_append(&insert_sql, "VALUES('[Content_Types].xml', ?) UNION ALL \n");
        
        /* _rels/.rels */
        strbuf_append(&insert_sql, "VALUES('_rels/.rels', ?) UNION ALL \n");
        
        /* xl/_rels/workbook.xml.rels */
        strbuf_append(&insert_sql, "VALUES('xl/_rels/workbook.xml.rels', ?) UNION ALL \n");
        
        /* xl/workbook.xml */
        strbuf_append(&insert_sql, "VALUES('xl/workbook.xml', ?) UNION ALL \n");
        
        /* xl/worksheets/sheetN.xml for each sheet */
        for (i = 0; i < sheet_count; i++) {
            strbuf_appendf(&insert_sql, "VALUES('xl/worksheets/sheet%d.xml', ?) UNION ALL \n", i + 1);
        }
        
        /* xl/styles.xml */
        strbuf_append(&insert_sql, "VALUES('xl/styles.xml', ?)) \n");

        strbuf_append(&insert_sql, "SELECT writefile(?, (SELECT zipfile(name, data) FROM contents))");

        fprintf(stderr, "zipfile_insert_via_vtab() %s\n", insert_sql.str);
        rc = sqlite3_prepare_v2(db, insert_sql.str, -1, &stmt, NULL);
        strbuf_free(&insert_sql);
        
        if (rc != SQLITE_OK) {
            sqlite3_result_error(context, 
                "Failed to prepare zipfile statement. Is the zipfile extension loaded?", -1);
            goto cleanup;
        }
        
        /* Bind all parameters */
        int param = 1;
        sqlite3_bind_text(stmt, param++, content_types, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, param++, rels, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, param++, workbook_rels, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, param++, workbook, -1, SQLITE_STATIC);
        for (i = 0; i < sheet_count; i++) {
            sqlite3_bind_text(stmt, param++, sheet_contents[i], -1, SQLITE_STATIC);
        }
        sqlite3_bind_text(stmt, param++, styles, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, param++, filename, -1, SQLITE_STATIC);
        
        rc = sqlite3_step(stmt);
        sqlite3_finalize(stmt);
        stmt = NULL;
        
        if (rc != SQLITE_ROW && rc != SQLITE_DONE) {
            err_msg = sqlite3_mprintf("Failed to create ZIP file: %s", sqlite3_errmsg(db));
            sqlite3_result_error(context, err_msg, -1);
            sqlite3_free(err_msg);
            goto cleanup;
        }
    }
    
    /* Return the filename on success, with warning if cells were truncated */
    if (warnings.cells_truncated > 0) {
        char *result = sqlite3_mprintf(
            "%s (WARNING: %d cell(s) exceeded Excel's %d character limit and were truncated. "
            "First occurrence: table '%s', row %d, column %d)",
            filename, warnings.cells_truncated, EXCEL_MAX_CELL_SIZE,
            warnings.first_truncated_table, warnings.first_truncated_row,
            warnings.first_truncated_col);
        sqlite3_result_text(context, result, -1, sqlite3_free);
    } else {
        sqlite3_result_text(context, filename, -1, SQLITE_TRANSIENT);
    }
    
cleanup:
    if (stmt) sqlite3_finalize(stmt);
    if (sheet_names) sqlite3_free(sheet_names);
    if (sheet_contents) {
        for (i = 0; i < sheet_count; i++) {
            sqlite3_free(sheet_contents[i]);
        }
        sqlite3_free(sheet_contents);
    }
    sqlite3_free(content_types);
    sqlite3_free(rels);
    sqlite3_free(workbook_rels);
    sqlite3_free(workbook);
    sqlite3_free(styles);
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
        SQLITE_UTF8,
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
