cmake_minimum_required(VERSION 3.16)

# =============================================================================
# Project: xlsximpex_ai
#
# One-line build commands:
#   Linux native:
#     cmake -S . -B build && cmake --build build
#
#   Windows cross-compile (MinGW):
#     cmake -S . -B build-win -DCMAKE_SYSTEM_NAME=Windows -DCMAKE_C_COMPILER=x86_64-w64-mingw32-gcc -DCMAKE_CXX_COMPILER=x86_64-w64-mingw32-g++ && cmake --build build-win
# =============================================================================

project(xlsximpex_ai LANGUAGES C CXX)

include(FetchContent)

# -----------------------------------------------------------------------------
# Global warnings
# -----------------------------------------------------------------------------

if (MSVC)
    add_compile_options(/W4)
else()
    add_compile_options(-Wall -Wextra -Wpedantic)
endif()

# -----------------------------------------------------------------------------
# zlib (static, required by OpenXLSX)
# -----------------------------------------------------------------------------

FetchContent_Declare(
    zlib
    URL https://zlib.net/zlib-1.3.1.tar.gz
)

set(ZLIB_BUILD_SHARED OFF CACHE BOOL "" FORCE)
set(ZLIB_BUILD_STATIC ON  CACHE BOOL "" FORCE)

FetchContent_MakeAvailable(zlib)

# Make sure find_package(ZLIB) works
set(ZLIB_ROOT ${zlib_SOURCE_DIR} CACHE PATH "" FORCE)
set(ZLIB_USE_STATIC_LIBS ON CACHE BOOL "" FORCE)

# -----------------------------------------------------------------------------
# OpenXLSX (static)
# -----------------------------------------------------------------------------

FetchContent_Declare(
    openxlsx
    GIT_REPOSITORY https://github.com/troldal/OpenXLSX.git
    GIT_TAG master
)

set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)

FetchContent_MakeAvailable(openxlsx)

# -----------------------------------------------------------------------------
# SQLite extension
# -----------------------------------------------------------------------------

add_subdirectory(chatgpt_libopenxlsx)
