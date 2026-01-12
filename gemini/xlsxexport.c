/*
PROMPTS USED:
Create in lsxexport.c the C code for a SQLite extension named xlsxexport that 
contains a SQL function named xlsx_export that saves multiple tables 
as a single XLSX spreadsheet, with the sheet names equal to the table 
names, and the sheet headers in bold and with autofilter.
xlsx_export returns the number of tables exported. 
If invoked with only one parameter then exports all the tables in the schema
The XLSX file is a ZIP archive with XML files inside. 
Use the SQLite zipfile extension to create this ZIP archive. 
Do not use any external library to write XML files. 
Warn if the Excel maximum cell size is exceeded.
Sanitize the sheet names to conform to Excel restrictions.
Add SQL function xlsx_export_version returning "2026-01-07 Gemini 3 Pro (High)".
Include as comments the prompts used.

USAGE IN SQLITE:
    .load xlsxexport
    SELECT xlsx_export('output.xlsx');  -- Export all tables in the schema
    SELECT xlsx_export('output.xlsx', 'table1', 'table2', 'table3');
    -- or with a single table:
    SELECT xlsx_export('output.xlsx', 'mytable');
    SELECT xlsx_export_version();
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <stdarg.h>

/* Wrapper around snprintf to warn on truncation */
static int warn_snprintf(char *str, size_t size, const char *format, ...) {
    va_list args;
    int ret;
    va_start(args, format);
    ret = vsnprintf(str, size, format, args);
    va_end(args);
    if (ret >= 0 && (size_t)ret >= size) {
        fprintf(stderr, "snprintf truncation: needed %d, available %d\n", ret, (int)size);
    }
    return ret;
}

/* Limits */
#define MAX_CELL_LEN 32767

/* Buffer for string construction */
typedef struct {
    char *data;
    size_t len;
    size_t cap;
} StrBuf;

static void strbuf_init(StrBuf *sb) {
    sb->data = NULL;
    sb->len = 0;
    sb->cap = 0;
}

static void strbuf_append(StrBuf *sb, const char *s, int len) {
    if (len < 0) len = (int)strlen(s);
    if (sb->len + len + 1 > sb->cap) {
        size_t new_cap = sb->cap ? sb->cap * 2 : 1024;
        while (new_cap < sb->len + len + 1) new_cap *= 2;
        sb->data = sqlite3_realloc(sb->data, new_cap);
        sb->cap = new_cap;
    }
    memcpy(sb->data + sb->len, s, len);
    sb->len += len;
    sb->data[sb->len] = '\0';
}

static void strbuf_free(StrBuf *sb) {
    sqlite3_free(sb->data);
    strbuf_init(sb);
}

static void strbuf_append_xml_escaped(StrBuf *sb, const char *s) {
    while (*s) {
        switch (*s) {
            case '<': strbuf_append(sb, "&lt;", 4); break;
            case '>': strbuf_append(sb, "&gt;", 4); break;
            case '&': strbuf_append(sb, "&amp;", 5); break;
            case '"': strbuf_append(sb, "&quot;", 6); break;
            case '\'': strbuf_append(sb, "&apos;", 6); break;
            default: {
                char c = *s;
                strbuf_append(sb, &c, 1);
            }
        }
        s++;
    }
}

/* Sanitize sheet name */
static char *sanitize_sheet_name(const char *name) {
    char *sanitized = sqlite3_mprintf("%s", name);
    int len = 0;
    int i, j;
    
    /* Remove invalid chars: \ / ? * [ ] : */
    for (i = 0, j = 0; sanitized[i]; i++) {
        char c = sanitized[i];
        if (c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']' || c == ':') {
            continue; /* Skip */
        }
        sanitized[j++] = c;
    }
    sanitized[j] = '\0';
    len = j;
    
    /* Truncate to 31 chars */
    if (len > 31) sanitized[31] = '\0';
    
    return sanitized;
}

/* Helper to convert int to column name (A, B, ... AA, AB) */
static void int_to_col(int n, char *buf) {
    char temp[16];
    int i = 0;
    n++; /* 1-based */
    while (n > 0) {
        n--;
        temp[i++] = (char)('A' + (n % 26));
        n /= 26;
    }
    int j;
    for (j = 0; j < i; j++) {
        buf[j] = temp[i - 1 - j];
    }
    buf[i] = '\0';
}

/* Helper to write file to zipfile table */
static int write_to_zip(sqlite3 *db, const char *zip_table, const char *filename, const char *data, int len) {
    sqlite3_stmt *stmt;
    char *sql = sqlite3_mprintf("INSERT INTO \"%w\"(name, data) VALUES(?, ?)", zip_table);
    int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    sqlite3_free(sql);
    
    if (rc != SQLITE_OK) return rc;
    
    sqlite3_bind_text(stmt, 1, filename, -1, SQLITE_STATIC);
    sqlite3_bind_blob(stmt, 2, data, len, SQLITE_STATIC); /* data is not modified, so STATIC is safe if data persists or assume transient? Text copy. */
    /* safe to use TRANSIENT if we free data later, or STATIC if it's string literal/managed buffer */
    /* For safety let's use SQLITE_TRANSIENT for dynamic buffers passed here */
    
    rc = sqlite3_step(stmt);
    sqlite3_finalize(stmt);
    return (rc == SQLITE_DONE) ? SQLITE_OK : rc;
}

/* Generate [Content_Types].xml */
static void generate_content_types(StrBuf *sb, int num_sheets) {
    strbuf_append(sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
    strbuf_append(sb, "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n", -1);
    strbuf_append(sb, "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n", -1);
    strbuf_append(sb, "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n", -1);
    strbuf_append(sb, "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n", -1);
    strbuf_append(sb, "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n", -1);
    
    int i; char buf[512];
    for (i = 1; i <= num_sheets; i++) {
        warn_snprintf(buf, sizeof(buf), "  <Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n", i);
        strbuf_append(sb, buf, -1);
    }
    strbuf_append(sb, "</Types>", -1);
}

/* Generate _rels/.rels */
static void generate_rels(StrBuf *sb) {
    strbuf_append(sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
    strbuf_append(sb, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n", -1);
    strbuf_append(sb, "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n", -1);
    strbuf_append(sb, "</Relationships>", -1);
}

/* Generate xl/styles.xml (with bold font) */
static void generate_styles(StrBuf *sb) {
    strbuf_append(sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
    strbuf_append(sb, "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n", -1);
    strbuf_append(sb, "  <fonts count=\"2\">\n", -1);
    strbuf_append(sb, "    <font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>\n", -1);
    strbuf_append(sb, "    <font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>\n", -1);
    strbuf_append(sb, "  </fonts>\n", -1);
    strbuf_append(sb, "  <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\n", -1);
    strbuf_append(sb, "  <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n", -1);
    strbuf_append(sb, "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n", -1);
    strbuf_append(sb, "  <cellXfs count=\"2\">\n", -1);
    strbuf_append(sb, "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n", -1); /* Normal */
    strbuf_append(sb, "    <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>\n", -1); /* Bold */
    strbuf_append(sb, "  </cellXfs>\n", -1);
    strbuf_append(sb, "  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\n", -1);
    strbuf_append(sb, "</styleSheet>", -1);
}

/* Generate xl/workbook.xml */
static void generate_workbook(StrBuf *sb, int num_sheets, char **sheet_names, char **ranges) {
    strbuf_append(sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
    strbuf_append(sb, "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n", -1);
    strbuf_append(sb, "  <fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>\n", -1);
    strbuf_append(sb, "  <workbookPr defaultThemeVersion=\"124226\"/>\n", -1);
    strbuf_append(sb, "  <bookViews><workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/></bookViews>\n", -1);
    strbuf_append(sb, "  <sheets>\n", -1);
    int i; char buf[512];
    for (i = 0; i < num_sheets; i++) {
        /* Id points to rId found in xl/_rels/workbook.xml.rels. rId start at 1. sheetId start at 1. */
        warn_snprintf(buf, sizeof(buf), "    <sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>\n", sheet_names[i], i+1, i+1);
        strbuf_append(sb, buf, -1);
    }
    strbuf_append(sb, "  </sheets>\n", -1);
    if (ranges) {
        strbuf_append(sb, "  <definedNames>\n", -1);
        for (i = 0; i < num_sheets; i++) {
            if (ranges[i]) {
                warn_snprintf(buf, sizeof(buf), "    <definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"%d\" hidden=\"1\">%s!%s</definedName>\n", i, sheet_names[i], ranges[i]);
                strbuf_append(sb, buf, -1);
            }
        }
        strbuf_append(sb, "  </definedNames>\n", -1);
    }
    strbuf_append(sb, "  <calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>\n", -1);
    strbuf_append(sb, "</workbook>", -1);
}

/* Generate xl/_rels/workbook.xml.rels */
static void generate_workbook_rels(StrBuf *sb, int num_sheets) {
    strbuf_append(sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
    strbuf_append(sb, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n", -1);
    int i; char buf[512];
    for (i = 0; i < num_sheets; i++) {
        warn_snprintf(buf, sizeof(buf), "  <Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>\n", i+1, i+1);
        strbuf_append(sb, buf, -1);
    }
    warn_snprintf(buf, sizeof(buf), "  <Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n", num_sheets + 1);
    strbuf_append(sb, buf, -1);
    
    strbuf_append(sb, "</Relationships>", -1);
}

/* Main Export Function */
static void xlsx_export(sqlite3_context *context, int argc, sqlite3_value **argv) {
    if (argc < 1) {
        sqlite3_result_error(context, "Usage: xlsx_export(filename [, table_name1, ...])", -1);
        return;
    }

    const char *filename = (const char *)sqlite3_value_text(argv[0]);
    sqlite3 *db = sqlite3_context_db_handle(context);
    
    char zip_table_name[64];
    /* Create a unique temp table name */
    warn_snprintf(zip_table_name, sizeof(zip_table_name), "temp_xlsx_zip_%p", (void*)filename);
    
    /* Drop if exists */
    char *sql = sqlite3_mprintf("DROP TABLE IF EXISTS \"%w\"", zip_table_name);
    sqlite3_exec(db, sql, NULL, NULL, NULL);
    sqlite3_free(sql);
    
    /* Create zipfile virtual table */
    sql = sqlite3_mprintf("CREATE VIRTUAL TABLE \"%w\" USING zipfile('%q')", zip_table_name, filename);
    int rc = sqlite3_exec(db, sql, NULL, NULL, NULL);
    sqlite3_free(sql);
    
    if (rc != SQLITE_OK) {
        sqlite3_result_error(context, "Failed to create internal zipfile table. Is zipfile extension loaded?", -1);
        return;
    }
    
    int num_found = 0;
    char **sheet_names = NULL;
    char **table_names = NULL;
    int t;
    
    if (argc == 1) {
        /* Export all tables */
        sqlite3_stmt *stmt;
        rc = sqlite3_prepare_v2(db, "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name NOT LIKE 'temp_xlsx_zip_%' ORDER BY name", -1, &stmt, NULL);
        if (rc == SQLITE_OK) {
            while (sqlite3_step(stmt) == SQLITE_ROW) {
                const char *tbl = (const char *)sqlite3_column_text(stmt, 0);
                table_names = sqlite3_realloc(table_names, (num_found + 1) * sizeof(char*));
                sheet_names = sqlite3_realloc(sheet_names, (num_found + 1) * sizeof(char*));
                table_names[num_found] = sqlite3_mprintf("%s", tbl);
                sheet_names[num_found] = sanitize_sheet_name(tbl);
                num_found++;
            }
            sqlite3_finalize(stmt);
        }
    } else {
        /* Export specified tables */
        for (t = 0; t < argc - 1; t++) {
            const char *tbl = (const char *)sqlite3_value_text(argv[t+1]);
            /* Check if table exists */
            sqlite3_stmt *stmt;
            char *check_sql = sqlite3_mprintf("SELECT 1 FROM sqlite_master WHERE type IN ('table','view') AND name='%q'", tbl);
            rc = sqlite3_prepare_v2(db, check_sql, -1, &stmt, NULL);
            sqlite3_free(check_sql);
            if (rc == SQLITE_OK) {
                if (sqlite3_step(stmt) == SQLITE_ROW) {
                    table_names = sqlite3_realloc(table_names, (num_found + 1) * sizeof(char*));
                    sheet_names = sqlite3_realloc(sheet_names, (num_found + 1) * sizeof(char*));
                    table_names[num_found] = sqlite3_mprintf("%s", tbl);
                    sheet_names[num_found] = sanitize_sheet_name(tbl);
                    num_found++;
                }
                sqlite3_finalize(stmt);
            }
        }
    }
    
    if (num_found == 0) {
        /* Nothing to export */
        sql = sqlite3_mprintf("DROP TABLE \"%w\"", zip_table_name);
        sqlite3_exec(db, sql, NULL, NULL, NULL);
        sqlite3_free(sql);
        sqlite3_result_int(context, 0);
        return;
    }
    
    int num_sheets = num_found;
    char **ranges = sqlite3_malloc(num_sheets * sizeof(char*));
    memset(ranges, 0, num_sheets * sizeof(char*));
    
    /* Process structure bits */
    StrBuf sb;
    
    /* 1. [Content_Types].xml */
    strbuf_init(&sb);
    generate_content_types(&sb, num_sheets);
    write_to_zip(db, zip_table_name, "[Content_Types].xml", sb.data, (int)sb.len);
    strbuf_free(&sb);
    
    /* 2. _rels/.rels */
    strbuf_init(&sb);
    generate_rels(&sb);
    write_to_zip(db, zip_table_name, "_rels/.rels", sb.data, (int)sb.len);
    strbuf_free(&sb);
    
    /* 3. xl/styles.xml */
    strbuf_init(&sb);
    generate_styles(&sb);
    write_to_zip(db, zip_table_name, "xl/styles.xml", sb.data, (int)sb.len);
    strbuf_free(&sb);
    
    /* 4. xl/workbook.xml & sheet names processing */
    /* sheet_names already populated */
    
    /* Wait: workbook.xml needs ranges, we must process sheets first! */
    /* Let's delay workbook.xml until after sheets are processed. */
    
    /* 5. xl/_rels/workbook.xml.rels */
    strbuf_init(&sb);
    generate_workbook_rels(&sb, num_sheets);
    write_to_zip(db, zip_table_name, "xl/_rels/workbook.xml.rels", sb.data, (int)sb.len);
    strbuf_free(&sb);
    
    /* 6. Process each table -> sheet */
    for (t = 0; t < num_sheets; t++) {
        const char *tbl_in = table_names[t];
        StrBuf sb_data;
        strbuf_init(&sb_data);
        strbuf_append(&sb_data, "  <sheetData>\n", -1);
        
        /* Get Columns */
        sqlite3_stmt *stmt;
        sql = sqlite3_mprintf("PRAGMA table_info(\"%w\")", tbl_in);
        rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
        sqlite3_free(sql);
        
        int col_count = 0;
        char **cols = NULL;
        
        if (rc == SQLITE_OK) {
             /* Row 1: Header */
             strbuf_append(&sb_data, "    <row r=\"1\">\n", -1);
             while (sqlite3_step(stmt) == SQLITE_ROW) {
                 const char *cname = (const char *)sqlite3_column_text(stmt, 1);
                 cols = sqlite3_realloc(cols, (col_count + 1) * sizeof(char*));
                 cols[col_count] = sqlite3_mprintf("%s", cname);
                 
                 char col_ref[16];
                 int_to_col(col_count, col_ref);
                 
                 char cell_open[64];
                 /* s=1 style applies bold */
                 warn_snprintf(cell_open, sizeof(cell_open), "      <c r=\"%s1\" t=\"inlineStr\" s=\"1\"><is><t>", col_ref);
                 strbuf_append(&sb_data, cell_open, -1);
                 strbuf_append_xml_escaped(&sb_data, cname);
                 strbuf_append(&sb_data, "</t></is></c>\n", -1);
                 col_count++;
             }
             strbuf_append(&sb_data, "    </row>\n", -1);
             sqlite3_finalize(stmt);
        } else {
             /* Table might not exist or error - should have been caught in pre-validation */
             strbuf_free(&sb_data);
             continue;
        }
        
        /* Get Data */
        sql = sqlite3_mprintf("SELECT * FROM \"%w\"", tbl_in);
        rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
        sqlite3_free(sql);
        
        int r_idx = 2; /* 1-based, header was 1 */
        if (rc == SQLITE_OK) {
            while (sqlite3_step(stmt) == SQLITE_ROW) {
                char row_open[32];
                warn_snprintf(row_open, sizeof(row_open), "    <row r=\"%d\">\n", r_idx);
                strbuf_append(&sb_data, row_open, -1);
                
                int c;
                for (c = 0; c < col_count; c++) {
                    char col_ref[16];
                    int_to_col(c, col_ref);
                    
                    if (sqlite3_column_type(stmt, c) == SQLITE_NULL) continue;
                    
                    int type = sqlite3_column_type(stmt, c);
                    if (type == SQLITE_INTEGER || type == SQLITE_FLOAT) {
                        char cell_open[64];
                        warn_snprintf(cell_open, sizeof(cell_open), "      <c r=\"%s%d\" t=\"n\"><v>", col_ref, r_idx);
                        strbuf_append(&sb_data, cell_open, -1);
                        const char *val = (const char *)sqlite3_column_text(stmt, c);
                        strbuf_append(&sb_data, val, -1);
                        strbuf_append(&sb_data, "</v></c>\n", -1);
                    } else {
                        const char *val = (const char *)sqlite3_column_text(stmt, c);
                        if (val) {
                            int len = (int)strlen(val);
                            if (len > MAX_CELL_LEN) {
                                char *msg = sqlite3_mprintf("Warning: Cell at %s%d exceeds %d characters (size: %d).", col_ref, r_idx, MAX_CELL_LEN, len);
                                sqlite3_log(SQLITE_WARNING, "%s", msg);
                                sqlite3_free(msg);
                            }
                            char cell_open[64];
                            warn_snprintf(cell_open, sizeof(cell_open), "      <c r=\"%s%d\" t=\"inlineStr\"><is><t>", col_ref, r_idx);
                            strbuf_append(&sb_data, cell_open, -1);
                            strbuf_append_xml_escaped(&sb_data, val);
                            strbuf_append(&sb_data, "</t></is></c>\n", -1);
                        }
                    }
                }
                strbuf_append(&sb_data, "    </row>\n", -1);
                r_idx++;
            }
            sqlite3_finalize(stmt);
        }
        
        strbuf_append(&sb_data, "  </sheetData>\n", -1);
        
        /* Assemble final XML */
        strbuf_init(&sb);
        strbuf_append(&sb, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", -1);
        strbuf_append(&sb, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n", -1);

        char first[16], last[16];
        int_to_col(0, first);
        int_to_col(col_count > 0 ? col_count - 1 : 0, last);
        
        /* Add Dimension tag */
        char dim[128];
        warn_snprintf(dim, sizeof(dim), "  <dimension ref=\"%s1:%s%d\"/>\n", first, last, r_idx > 1 ? r_idx - 1 : 1);
        strbuf_append(&sb, dim, -1);

        /* Add sheetViews */
        strbuf_append(&sb, "  <sheetViews>\n    <sheetView tabSelected=\"1\" workbookViewId=\"0\"/>\n  </sheetViews>\n", -1);
        
        /* Add sheetFormatPr */
        strbuf_append(&sb, "  <sheetFormatPr defaultRowHeight=\"15\"/>\n", -1);
        
        /* Add collected sheet data */
        if (sb_data.data) strbuf_append(&sb, sb_data.data, -1);
        
        /* Handle Autofilter */
        if (col_count > 0) {
            char af[128];
            /* Autofilter on row-based range */
            warn_snprintf(af, sizeof(af), "  <autoFilter ref=\"%s1:%s%d\"/>\n", first, last, r_idx > 1 ? r_idx - 1 : 1);
            strbuf_append(&sb, af, -1);
            /* Store range for workbook.xml definedNames */
            ranges[t] = sqlite3_mprintf("$%s$1:$%s$%d", first, last, r_idx > 1 ? r_idx - 1 : 1);
        }
        
        strbuf_append(&sb, "</worksheet>", -1);
        
        char path[64];
        warn_snprintf(path, sizeof(path), "xl/worksheets/sheet%d.xml", t+1);
        write_to_zip(db, zip_table_name, path, sb.data, (int)sb.len);
        
        strbuf_free(&sb);
        strbuf_free(&sb_data);
        
        /* Cleanup cols */
        int z;
        for (z = 0; z < col_count; z++) sqlite3_free(cols[z]);
        sqlite3_free(cols);
    }
    
    /* 4. xl/workbook.xml (Now we have ranges) */
    strbuf_init(&sb);
    generate_workbook(&sb, num_sheets, sheet_names, ranges);
    write_to_zip(db, zip_table_name, "xl/workbook.xml", sb.data, (int)sb.len);
    strbuf_free(&sb);

    /* Cleanup sheet info */
    for (t = 0; t < num_sheets; t++) {
        sqlite3_free(sheet_names[t]);
        sqlite3_free(table_names[t]);
        sqlite3_free(ranges[t]);
    }
    sqlite3_free(sheet_names);
    sqlite3_free(table_names);
    sqlite3_free(ranges);
    
    /* Close zip (DROP TABLE) */
    sql = sqlite3_mprintf("DROP TABLE \"%w\"", zip_table_name);
    sqlite3_exec(db, sql, NULL, NULL, NULL);
    sqlite3_free(sql);
    
    sqlite3_result_int(context, num_sheets);
}

/* Version Function */
static void xlsx_export_version(sqlite3_context *context, int argc, sqlite3_value **argv) {
    (void)argc; (void)argv;
    sqlite3_result_text(context, "2026-01-07 Gemini 3 Pro (High)", -1, SQLITE_STATIC);
}

/* Init */
#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsxexport_init(sqlite3 *db, char **pzErrMsg, const sqlite3_api_routines *pApi) {
    int rc = SQLITE_OK;
    SQLITE_EXTENSION_INIT2(pApi);
    (void)pzErrMsg;
    
    rc = sqlite3_create_function(db, "xlsx_export", -1, SQLITE_UTF8, NULL, xlsx_export, NULL, NULL);
    if (rc == SQLITE_OK) {
        rc = sqlite3_create_function(db, "xlsx_export_version", 0, SQLITE_UTF8, NULL, xlsx_export_version, NULL, NULL);
    }
    return rc;
}
