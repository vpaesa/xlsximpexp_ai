cmake_minimum_required(VERSION 3.16)
# Invocation:
# cmake -S . -B build && cmake --build build

project(xlsxexport LANGUAGES C CXX)

# -----------------------------------------------------------------------------
# Language standards
# -----------------------------------------------------------------------------

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_C_EXTENSIONS OFF)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# -----------------------------------------------------------------------------
# Warnings
# -----------------------------------------------------------------------------

if (MSVC)
    add_compile_options(/W4)
else()
    add_compile_options(-Wall -Wextra -Wpedantic)
endif()

# -----------------------------------------------------------------------------
# Define the target FIRST
# -----------------------------------------------------------------------------

add_library(xlsxexport MODULE
    xlsxexport.cpp
)

# SQLite extensions must not have 'lib' prefix
set_target_properties(xlsxexport PROPERTIES
    PREFIX ""
)

# Platform-specific suffix
if (WIN32)
    set_target_properties(xlsxexport PROPERTIES SUFFIX ".dll")
elseif (APPLE)
    set_target_properties(xlsxexport PROPERTIES SUFFIX ".dylib")
else()
    set_target_properties(xlsxexport PROPERTIES SUFFIX ".so")
endif()

# -----------------------------------------------------------------------------
# SQLite amalgamation headers (NO linking)
# -----------------------------------------------------------------------------

target_include_directories(xlsxexport PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../../sqlite-amalgamation-3510200
)

# -----------------------------------------------------------------------------
# Link OpenXLSX (AFTER target exists)
# -----------------------------------------------------------------------------

target_link_libraries(xlsxexport PRIVATE
    OpenXLSX::OpenXLSX
)

# -----------------------------------------------------------------------------
# Static runtime for MinGW
# -----------------------------------------------------------------------------

if (WIN32 AND NOT MSVC)
    target_link_options(xlsxexport PRIVATE
        -static
        -static-libstdc++
        -static-libgcc
    )
endif()
