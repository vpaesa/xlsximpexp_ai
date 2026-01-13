/*
User prompt history (embedded for traceability):

1. "is manadatory that xl/sharedStrings.xml has \"sst:uniqueCount\" with count of unique strings"
2. "Create the C code for a SQLite extension named xlsxexport ..."
3. "Include build example (Linux native an crosscompiling for Windows) and usage in code"
4. "Include OpenXLSX build example (Linux native and crosscompiling for Windows) in code"
5. "full CMakeLists.txt for xlsxexport"
6. "linking to sqlite library is not needed. Just include -I../../sqlite-amalgamation-3510200"
7. "add a superbuild that fetches SQLite amalgamation + OpenXLSX"
8. "assume project folder name is chatgpt_libopenxlsx"
9. "assume project name is xlsximpex_ai and subdirectory is chatgpt_libopenxlsx"
10. "Add all user prompts as comments to xlsxexport.cpp"
11. "include the full xlsx_export implementation (not just the version function), add a README.md"
*/

/* Usage:
.load ./xlsxexport.so

-- Export all tables
SELECT xlsx_export('database.xlsx');

-- Export specific tables
SELECT xlsx_export('subset.xlsx', 'users', 'orders', 'audit_log');

-- Check version
SELECT xlsx_export_version();

*/
/* Linux Build:
sudo apt install cmake g++ zlib1g-dev
git clone https://github.com/troldal/OpenXLSX.git
cd OpenXLSX
cmake -S . -B build-static \
  -DCMAKE_BUILD_TYPE=Release \
  -DBUILD_SHARED_LIBS=OFF \
  -DOPENXLSX_BUILD_TESTS=OFF \
  -DOPENXLSX_BUILD_SAMPLES=OFF

cmake --build build-static
sudo cmake --install build-static

g++ -shared -fPIC xlsxexport.cpp \
    -o xlsxexport.so \
    -std=c11 -Wall -Wextra -Wpedantic \
    -I/usr/local/include -I../../sqlite-amalgamation-3510200 \
    -L/usr/local/lib \
    -lOpenXLSX

*/
/* Linux crosscompiling for Windows 64:
sudo apt install zlib1g-dev-mingw-w64
cmake -S OpenXLSX -B build-win \
  -DCMAKE_TOOLCHAIN_FILE=mingw-x64.cmake \
  -DCMAKE_BUILD_TYPE=Release \
  -DBUILD_SHARED_LIBS=OFF \
  -DOPENXLSX_BUILD_TESTS=OFF \
  -DOPENXLSX_BUILD_SAMPLES=OFF \
  -DZLIB_LIBRARY=/usr/x86_64-w64-mingw32/lib/libz.a \
  -DZLIB_INCLUDE_DIR=/usr/x86_64-w64-mingw32/include
x86_64-w64-mingw32-g++ -shared \
    xlsxexport.cpp \
    -o xlsxexport.dll \
    -std=c11 -Wall -Wextra -Wpedantic \
    -static-libstdc++ -static-libgcc \
    -lz -lOpenXLSX
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <OpenXLSX.hpp>
#include <string>
#include <vector>
#include <set>
#include <cstring>
#include <cctype>

using namespace OpenXLSX;

#define EXCEL_MAX_CELL_BYTES 32767
#define EXCEL_MAX_SHEETNAME 31

static std::string sanitize_sheet_name(const char* name) {
    std::string out;
    const char* forbidden = "[]:*?/\\";
    for (const char* p = name; *p; ++p) {
        if (std::strchr(forbidden, *p)) continue;
        if (std::iscntrl((unsigned char)*p)) continue;
        out.push_back(*p);
    }
    if (out.empty()) out = "Sheet";
    if (out.size() > EXCEL_MAX_SHEETNAME) out.resize(EXCEL_MAX_SHEETNAME);
    return out;
}

static void xlsx_export_func(sqlite3_context* ctx, int argc, sqlite3_value** argv) {
    if (argc < 1) {
        sqlite3_result_error(ctx, "xlsx_export(filename, [tables...])", -1);
        return;
    }

    sqlite3* db = sqlite3_context_db_handle(ctx);
    const char* filename = (const char*)sqlite3_value_text(argv[0]);
    if (!filename) {
        sqlite3_result_error(ctx, "invalid filename", -1);
        return;
    }

    std::vector<std::string> tables;
    if (argc == 1) {
        sqlite3_stmt* st = nullptr;
        sqlite3_prepare_v2(db,
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'",
            -1, &st, nullptr);
        while (sqlite3_step(st) == SQLITE_ROW)
            tables.emplace_back((const char*)sqlite3_column_text(st, 0));
        sqlite3_finalize(st);
    } else {
        for (int i = 1; i < argc; ++i) {
            const char* t = (const char*)sqlite3_value_text(argv[i]);
            if (t && *t) tables.emplace_back(t);
        }
    }

    OpenXLSX::XLDocument doc{std::string(filename)};
    auto wb = doc.workbook();

    // Access styles collection
    auto styles = doc.styles();

    /* 1. Create font (returns index) */
    auto fontId = styles.fonts().create();

    /* 2. Access font by index */
    styles.fonts()[fontId].setBold(true);

    /* 3. Create cell format (returns index) */
    auto formatId = styles.cellFormats().create();

    /* 4. Access cell format by index and assign font */
    styles.cellFormats()[formatId].setFontIndex(fontId);

    std::set<std::string> used;
    int exported = 0;

    for (const auto& table : tables) {
        std::string sheet = sanitize_sheet_name(table.c_str());
        while (used.count(sheet)) sheet += "_";
        used.insert(sheet);

        wb.addWorksheet(sheet);
        auto ws = wb.worksheet(sheet);

        std::string sql = "SELECT * FROM \"" + table + "\"";
        sqlite3_stmt* st = nullptr;
        if (sqlite3_prepare_v2(db, sql.c_str(), -1, &st, nullptr) != SQLITE_OK)
            continue;

        // Write header row
        int cols = sqlite3_column_count(st);
        for (int c = 0; c < cols; ++c) {
            auto cell = ws.cell(1, c + 1);
            cell.value() = sqlite3_column_name(st, c);
            cell.setCellFormat(formatId); 
        }
        // OpenXLSX does not support auto filters
        //ws.setAutoFilter(OpenXLSX::XLCellRange(1, 1, 1, cols));

        // Write data rows
        int row = 2;
        while (sqlite3_step(st) == SQLITE_ROW) {
            for (int c = 0; c < cols; ++c) {
                const char* v = (const char*)sqlite3_column_text(st, c);
                if (v && strlen(v) > EXCEL_MAX_CELL_BYTES)
                    sqlite3_log(SQLITE_WARNING, "Cell exceeds Excel limit");
                ws.cell(row, c + 1).value() = v ? v : "";
            }
            ++row;
        }
        sqlite3_finalize(st);
        ++exported;
    }

    doc.save();
    doc.close();
    sqlite3_result_int(ctx, exported);
}

static void xlsx_export_version(sqlite3_context* ctx, int, sqlite3_value**) {
    sqlite3_result_text(ctx, "2026-01-07 ChatGPT-5.2", -1, SQLITE_STATIC);
}

extern "C" int sqlite3_xlsxexport_init(
    sqlite3* db, char**, const sqlite3_api_routines* api) {
    SQLITE_EXTENSION_INIT2(api);
    sqlite3_create_function(db, "xlsx_export", -1, SQLITE_UTF8, nullptr,
                            xlsx_export_func, nullptr, nullptr);
    sqlite3_create_function(db, "xlsx_export_version", 0, SQLITE_UTF8,
                            nullptr, xlsx_export_version, nullptr, nullptr);
    return SQLITE_OK;
}
