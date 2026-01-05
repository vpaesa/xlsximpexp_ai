# Makefile for copilot_libxlsxwriter (xlsxexport with libxlsxwriter)

# Compiler settings
CC ?= gcc
CFLAGS ?= -O2 -fPIC -Wall -Wextra -I../../sqlite-amalgamation-3510100
# Assuming libxlsxwriter headers are in system path or standard include path.
# If not, add -I/path/to/libxlsxwriter/include

# Cross-compiler for Windows 64-bit (MinGW-w64)
CC_WIN64 ?= x86_64-w64-mingw32-gcc
# Adjust include path for sqlite3 if necessary, matching copilot/Makefile logic
CFLAGS_WIN64 ?= -O2 -Wall -Wextra -I../../sqlite-amalgamation-3510100

ifeq ($(OS),Windows_NT)
    SHARED_EXT = .dll
    LDFLAGS = -shared
else
    UNAME := $(shell uname)
    ifeq ($(UNAME),Darwin)
        SHARED_EXT = .dylib
        LDFLAGS = -dynamiclib
    else
        SHARED_EXT = .so
        LDFLAGS = -shared
    endif
endif

# Targets
TARGET_EXPORT = xlsxexport$(SHARED_EXT)
TARGET_EXPORT_WIN64 = xlsxexport.dll

all: $(TARGET_EXPORT)

# Export (requires libxlsxwriter)
$(TARGET_EXPORT): xlsxexport.c
	$(CC) $(CFLAGS) $(LDFLAGS) -o $@ $< -lxlsxwriter
	strip $@

# Cross-compile target for Windows
win64: $(TARGET_EXPORT_WIN64)

$(TARGET_EXPORT_WIN64): xlsxexport.c
	$(CC_WIN64) $(CFLAGS_WIN64) -shared -o $@ $< -lxlsxwriter -lz -lmsvcrt
	x86_64-w64-mingw32-strip $@

clean:
	rm -f $(TARGET_EXPORT) $(TARGET_EXPORT_WIN64)

.PHONY: all win64 clean
