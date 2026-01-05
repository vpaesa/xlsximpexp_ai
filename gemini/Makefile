# Makefile for gemini folder extensions

# Compiler settings
CC ?= gcc
CFLAGS ?= -O2 -fPIC -Wall -Wextra -I../../sqlite-amalgamation-3510100

# Cross-compiler for Windows 64-bit (MinGW-w64)
CC_WIN64 ?= x86_64-w64-mingw32-gcc
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
TARGET_IMPORT = xlsximport$(SHARED_EXT)
TARGET_EXPORT = xlsxexport$(SHARED_EXT)

TARGET_IMPORT_WIN64 = xlsximport.dll
TARGET_EXPORT_WIN64 = xlsxexport.dll

all: $(TARGET_IMPORT) $(TARGET_EXPORT)

$(TARGET_IMPORT): xlsximport.c
	$(CC) $(CFLAGS) $(LDFLAGS) -o $@ $< -lexpat
	strip $@

$(TARGET_EXPORT): xlsxexport.c
	$(CC) $(CFLAGS) $(LDFLAGS) -o $@ $<
	strip $@

# Cross-compile
win64: $(TARGET_IMPORT_WIN64) $(TARGET_EXPORT_WIN64)

$(TARGET_IMPORT_WIN64): xlsximport.c
	$(CC_WIN64) $(CFLAGS_WIN64) -shared -o $@ $< -lexpat
	x86_64-w64-mingw32-strip $@

$(TARGET_EXPORT_WIN64): xlsxexport.c
	$(CC_WIN64) $(CFLAGS_WIN64) -shared -o $@ $<
	x86_64-w64-mingw32-strip $@


clean:
	rm -f $(TARGET_IMPORT) $(TARGET_EXPORT) $(TARGET_IMPORT_WIN64) $(TARGET_EXPORT_WIN64)

.PHONY: all win64 clean
