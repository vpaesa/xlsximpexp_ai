# Makefile for copilot folder extensions (xlsximport)

# Compiler settings
CC ?= gcc
CFLAGS ?= -O2 -fPIC -Wall -Wextra

# Cross-compiler for Windows 64-bit (MinGW-w64)
CC_WIN64 ?= x86_64-w64-mingw32-gcc
CFLAGS_WIN64 ?= -O2 -Wall -Wextra -I../../sqlite-amalgamation-3510100

ifeq ($(OS),Windows_NT)
    SHARED_EXT = .dll
    LDFLAGS = -shared
    LIBS = -lexpat
else
    UNAME := $(shell uname)
    ifeq ($(UNAME),Darwin)
        SHARED_EXT = .dylib
        LDFLAGS = -dynamiclib
    else
        SHARED_EXT = .so
        LDFLAGS = -shared
    endif
    LIBS = -lexpat
endif

# Targets
TARGET_IMPORT = xlsximport$(SHARED_EXT)
TARGET_IMPORT_WIN64 = xlsximport.dll
TARGET_EXPORT = xlsxexport$(SHARED_EXT)
TARGET_EXPORT_WIN64 = xlsxexport.dll

all: $(TARGET_IMPORT) $(TARGET_EXPORT)

# Import
$(TARGET_IMPORT): xlsximport.c
	$(CC) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIBS)
	strip $@

# Export 
# Note: For Windows, -lz -lmsvcrt
$(TARGET_EXPORT): xlsxexport.c
	$(CC) $(CFLAGS) $(LDFLAGS) -o $@ $< $(LIBS) -lz -lmsvcrt
	strip $@

# Cross-compile
win64: $(TARGET_IMPORT_WIN64) $(TARGET_EXPORT_WIN64)

$(TARGET_IMPORT_WIN64): xlsximport.c
	$(CC_WIN64) $(CFLAGS_WIN64) -shared -o $@ $< -lexpat
	x86_64-w64-mingw32-strip $@

$(TARGET_EXPORT_WIN64): xlsxexport.c
	$(CC_WIN64) $(CFLAGS_WIN64) -shared -o $@ $< -lz -lmsvcrt
	x86_64-w64-mingw32-strip $@

clean:
	rm -f $(TARGET_IMPORT) $(TARGET_IMPORT_WIN64) $(TARGET_EXPORT) $(TARGET_EXPORT_WIN64)

.PHONY: all win64 clean
