/*
xlsximport.c - SQLite extension to import XLSX files

Uses the SQLite zipfile extension to read XLSX archives and a small built-in
XML parser (no external libraries) for XML parsing. This is the "noexpat"
variant of the opus xlsximport extension: it is functionally identical to the
opus variant but carries its own minimal XML parser instead of linking expat.
Three SQL functions defined
xlsx_import() creates one table for each sheet in the XLSX file, with table name
equal to sheet name, and column names equal to the values in the first row of
the sheet. The first parameter is the XLSX filename. Subsequent optional parameters
are sheet names or sheet numbers (1-based) to import. The return value is the number of sheets imported.
xlsx_import_sheetnames() is a table-valued function that returns the names of the sheets in the file.
xlsx_import_version() returns the version string.

Usage:
.load xlsximport.so
SELECT xlsx_import('filename.xlsx');  -- Import all sheets
SELECT xlsx_import('filename.xlsx', 'Sheet1', 'Sheet2');  -- Import specific sheets by name
SELECT xlsx_import('filename.xlsx', 1, 3);  -- Import sheets by number (1-based)
SELECT xlsx_import('filename.xlsx', 'Sheet1', 2);  -- Mix of names and numbers
SELECT sheet_num, sheet_name FROM xlsx_import_sheetnames('filename.xlsx');
SELECT xlsx_import_version();
**
** ============================================================================
** REQUIREMENTS / DESIGN NOTES
** ============================================================================
**
** This extension uses the SQLite zipfile extension to open XLSX files and
** gather the following content:
**   - xl/sharedStrings.xml
**   - xl/worksheets/sheet1.xml to xl/worksheets/sheetN.xml
**   - xl/workbook.xml
**
** The name of each sheet is stored in xl/workbook.xml.
** The individual sheets are kept in xl/worksheets/sheet1.xml to sheetN.xml.
**
** To save on space, Microsoft stores all character literal values in one
** common xl/sharedStrings.xml dictionary file. The individual cell value
** found for a string in the actual sheet.xml file is just an index into
** this dictionary.
**
** Microsoft does not store empty cells or rows in xl/worksheets/sheet*.xml,
** so any gaps between values must be handled by this code.
**
** COLUMN NAME CONVERSION (Base-26):
** To figure out the number of skipped columns, we need to calculate the
** distance between cells like "AB67" and "C67". The way columns are named
** (A through Z, then AA through AZ, then BA through BZ, etc.) suggests a
** base-26 system. We use a simple conversion method from base-26 to decimal
** and then use subtraction to find empty cells between columns.
**
** XML STRUCTURE NOTES:
**   - xl/sharedStrings.xml has "sst:uniqueCount" with count of unique strings
**   - xl/worksheets/sheet*.xml has "dimension:ref" with enclosing cell range
**   - Inline strings use <c t="inlineStr"><is><t>value</t></is></c>
**   - Shared strings use <c t="s"><v>index</v></c>
**   - Numeric values use <c><v>number</v></c> (no type attribute)
**
** XML parsing is done using the small built-in parser in the "Minimal XML
** Parser" section below, so this variant has no external library dependency.
**
*/

#include <sqlite3ext.h>
SQLITE_EXTENSION_INIT1

#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

/*
** ============================================================================
** Minimal XML Parser (no external dependencies)
** ============================================================================
**
** This replaces the expat dependency used by the sibling "opus" variant.  It
** is a small, non-validating, single-pass parser tailored to the well-formed
** XML that Excel writes inside an XLSX archive.  It understands:
**   - start tags with attributes:        <c r="A1" t="s">
**   - empty-element (self-closing) tags:  <dimension ref="A1:B2"/>
**   - end tags:                           </c>
**   - character data between tags
**   - the five predefined entities and numeric character references
**     (&amp; &lt; &gt; &quot; &apos; &#NN; &#xNN;)
**   - and it skips the XML declaration, comments, CDATA markers and DOCTYPE.
**
** Element and attribute names are reported verbatim (including any namespace
** prefix, e.g. "r:id"), exactly as expat reports them when namespace
** processing is disabled.  The callback signatures deliberately mirror the
** expat handlers, so the parsing logic in the sections below is identical to
** the expat-based variant and only the parser engine differs.
*/

typedef char XML_Char;
#define XMLCALL

typedef void (*xml_start_handler)(void *user, const XML_Char *name,
                                  const XML_Char **atts);
typedef void (*xml_end_handler)(void *user, const XML_Char *name);
typedef void (*xml_char_handler)(void *user, const XML_Char *s, int len);

static int mxml_is_space(char c) {
  return c == ' ' || c == '\t' || c == '\n' || c == '\r';
}

static char *mxml_ndup(const char *s, int n) {
  char *r = (char *)malloc(n + 1);
  if (!r)
    return NULL;
  memcpy(r, s, n);
  r[n] = '\0';
  return r;
}

/* Append the UTF-8 encoding of code point cp to dst; return number of bytes. */
static int mxml_utf8_encode(unsigned cp, char *dst) {
  if (cp < 0x80) {
    dst[0] = (char)cp;
    return 1;
  } else if (cp < 0x800) {
    dst[0] = (char)(0xC0 | (cp >> 6));
    dst[1] = (char)(0x80 | (cp & 0x3F));
    return 2;
  } else if (cp < 0x10000) {
    dst[0] = (char)(0xE0 | (cp >> 12));
    dst[1] = (char)(0x80 | ((cp >> 6) & 0x3F));
    dst[2] = (char)(0x80 | (cp & 0x3F));
    return 3;
  } else {
    dst[0] = (char)(0xF0 | (cp >> 18));
    dst[1] = (char)(0x80 | ((cp >> 12) & 0x3F));
    dst[2] = (char)(0x80 | ((cp >> 6) & 0x3F));
    dst[3] = (char)(0x80 | (cp & 0x3F));
    return 4;
  }
}

/*
** Decode XML entities in src[0..len) into dst (which must hold at least
** len+1 bytes; decoded output is always <= the source length).  Returns the
** decoded length.  Unrecognized entities are copied through verbatim.
*/
static int mxml_decode(const char *src, int len, char *dst) {
  int di = 0;
  int si = 0;
  while (si < len) {
    char c = src[si];
    if (c != '&') {
      dst[di++] = c;
      si++;
      continue;
    }

    /* Locate the terminating ';' within a small window. */
    int semi = -1;
    for (int k = si + 1; k < len && k <= si + 11; k++) {
      if (src[k] == ';') {
        semi = k;
        break;
      }
      if (src[k] == '&' || src[k] == '<')
        break;
    }
    if (semi < 0) {
      dst[di++] = c; /* lone '&' */
      si++;
      continue;
    }

    const char *e = src + si + 1;
    int elen = semi - (si + 1);
    if (elen == 3 && memcmp(e, "amp", 3) == 0) {
      dst[di++] = '&';
    } else if (elen == 2 && memcmp(e, "lt", 2) == 0) {
      dst[di++] = '<';
    } else if (elen == 2 && memcmp(e, "gt", 2) == 0) {
      dst[di++] = '>';
    } else if (elen == 4 && memcmp(e, "quot", 4) == 0) {
      dst[di++] = '"';
    } else if (elen == 4 && memcmp(e, "apos", 4) == 0) {
      dst[di++] = '\'';
    } else if (elen >= 2 && e[0] == '#') {
      unsigned cp = 0;
      int ok = 1;
      if (e[1] == 'x' || e[1] == 'X') {
        if (elen < 3)
          ok = 0;
        for (int k = 2; k < elen && ok; k++) {
          char h = e[k];
          cp <<= 4;
          if (h >= '0' && h <= '9')
            cp += (unsigned)(h - '0');
          else if (h >= 'a' && h <= 'f')
            cp += (unsigned)(h - 'a' + 10);
          else if (h >= 'A' && h <= 'F')
            cp += (unsigned)(h - 'A' + 10);
          else
            ok = 0;
        }
      } else {
        for (int k = 1; k < elen && ok; k++) {
          char d = e[k];
          if (d >= '0' && d <= '9')
            cp = cp * 10 + (unsigned)(d - '0');
          else
            ok = 0;
        }
      }
      if (ok && cp > 0 && cp <= 0x10FFFF) {
        di += mxml_utf8_encode(cp, dst + di);
      } else {
        memcpy(dst + di, src + si, (size_t)(semi - si + 1));
        di += semi - si + 1;
      }
    } else {
      /* Unknown entity: copy verbatim including '&' and ';'. */
      memcpy(dst + di, src + si, (size_t)(semi - si + 1));
      di += semi - si + 1;
    }
    si = semi + 1;
  }
  dst[di] = '\0';
  return di;
}

/*
** Emit a run of character data.  Two transformations mirror what expat does
** so this variant produces byte-identical output:
**   1. XML end-of-line normalization (spec section 2.11): a literal CRLF or a
**      lone CR in the source is folded to a single LF.  Character references
**      such as &#13; are NOT affected, because normalization applies to the
**      literal input before reference expansion.
**   2. Entity decoding, performed only when an '&' is present.
*/
static int mxml_emit_text(const char *t, int runlen, void *user,
                          xml_char_handler on_char) {
  if (!on_char || runlen <= 0)
    return 0;
  int has_amp = memchr(t, '&', (size_t)runlen) != NULL;
  int has_cr = memchr(t, '\r', (size_t)runlen) != NULL;
  if (!has_amp && !has_cr) {
    on_char(user, t, runlen);
    return 0;
  }

  /* Fold line endings: "\r\n" and a lone "\r" both become "\n". */
  char *norm = (char *)malloc((size_t)runlen + 1);
  if (!norm)
    return -1;
  int n = 0;
  for (int i = 0; i < runlen; i++) {
    if (t[i] == '\r') {
      norm[n++] = '\n';
      if (i + 1 < runlen && t[i + 1] == '\n')
        i++;
    } else {
      norm[n++] = t[i];
    }
  }
  norm[n] = '\0';

  if (has_amp) {
    /* Decode into a separate buffer (decoded length is always <= n). */
    char *out = (char *)malloc((size_t)n + 1);
    if (!out) {
      free(norm);
      return -1;
    }
    int dl = mxml_decode(norm, n, out);
    on_char(user, out, dl);
    free(out);
  } else {
    on_char(user, norm, n);
  }
  free(norm);
  return 0;
}

/*
** Parse the XML document in xml[0..len) driving the supplied handlers.
** Returns 0 on success, -1 on a memory failure or truncated markup.
*/
static int mini_xml_parse(const char *xml, int len, void *user,
                          xml_start_handler on_start, xml_end_handler on_end,
                          xml_char_handler on_char) {
  const char *p = xml;
  const char *end = xml + len;

  while (p < end) {
    if (*p != '<') {
      const char *t = p;
      while (p < end && *p != '<')
        p++;
      if (mxml_emit_text(t, (int)(p - t), user, on_char) != 0)
        return -1;
      continue;
    }

    /* p is at '<'. */
    if (end - p >= 4 && memcmp(p, "<!--", 4) == 0) {
      const char *q = p + 4;
      while (q + 3 <= end && memcmp(q, "-->", 3) != 0)
        q++;
      if (q + 3 > end)
        return -1;
      p = q + 3;
      continue;
    }
    if (end - p >= 9 && memcmp(p, "<![CDATA[", 9) == 0) {
      const char *s = p + 9;
      const char *q = s;
      while (q + 3 <= end && memcmp(q, "]]>", 3) != 0)
        q++;
      if (q + 3 > end)
        return -1;
      if (on_char && q > s)
        on_char(user, s, (int)(q - s));
      p = q + 3;
      continue;
    }
    if (end - p >= 2 && p[1] == '?') {
      const char *q = p + 2;
      while (q + 2 <= end && memcmp(q, "?>", 2) != 0)
        q++;
      if (q + 2 > end)
        return -1;
      p = q + 2;
      continue;
    }
    if (end - p >= 2 && p[1] == '!') {
      /* DOCTYPE or similar declaration: skip to the next '>'. */
      const char *q = p + 2;
      while (q < end && *q != '>')
        q++;
      if (q >= end)
        return -1;
      p = q + 1;
      continue;
    }
    if (end - p >= 2 && p[1] == '/') {
      /* End tag. */
      const char *q = p + 2;
      const char *nstart = q;
      while (q < end && *q != '>' && !mxml_is_space(*q))
        q++;
      int nlen = (int)(q - nstart);
      while (q < end && *q != '>')
        q++;
      if (q >= end)
        return -1;
      if (on_end && nlen > 0) {
        char *nm = mxml_ndup(nstart, nlen);
        if (!nm)
          return -1;
        on_end(user, nm);
        free(nm);
      }
      p = q + 1;
      continue;
    }

    /* Start tag, possibly self-closing. */
    {
      const char *q = p + 1;
      const char *nstart = q;
      while (q < end && *q != '>' && *q != '/' && !mxml_is_space(*q))
        q++;
      int nlen = (int)(q - nstart);
      if (nlen == 0)
        return -1;

      char **atts = NULL;
      int natt = 0;
      int attcap = 0;
      int self_closing = 0;
      int fail = 0;

      while (q < end) {
        while (q < end && mxml_is_space(*q))
          q++;
        if (q >= end) {
          fail = 1;
          break;
        }
        if (*q == '>') {
          q++;
          break;
        }
        if (*q == '/') {
          self_closing = 1;
          q++;
          while (q < end && mxml_is_space(*q))
            q++;
          if (q < end && *q == '>')
            q++;
          break;
        }

        /* Attribute name. */
        const char *astart = q;
        while (q < end && *q != '=' && *q != '>' && *q != '/' &&
               !mxml_is_space(*q))
          q++;
        int anlen = (int)(q - astart);

        while (q < end && mxml_is_space(*q))
          q++;

        const char *aval = NULL;
        int avlen = 0;
        if (q < end && *q == '=') {
          q++;
          while (q < end && mxml_is_space(*q))
            q++;
          if (q < end && (*q == '"' || *q == '\'')) {
            char quote = *q;
            q++;
            const char *vstart = q;
            while (q < end && *q != quote)
              q++;
            aval = vstart;
            avlen = (int)(q - vstart);
            if (q < end)
              q++; /* closing quote */
          }
        }

        if (anlen > 0 && aval != NULL) {
          if (natt + 3 > attcap) {
            int ncap = attcap ? attcap * 2 : 8;
            char **na = (char **)realloc(atts, (size_t)ncap * sizeof(char *));
            if (!na) {
              fail = 1;
              break;
            }
            atts = na;
            attcap = ncap;
          }
          char *anm = mxml_ndup(astart, anlen);
          char *avl;
          if (memchr(aval, '&', (size_t)avlen) == NULL) {
            avl = mxml_ndup(aval, avlen);
          } else {
            avl = (char *)malloc((size_t)avlen + 1);
            if (avl)
              mxml_decode(aval, avlen, avl);
          }
          if (!anm || !avl) {
            free(anm);
            free(avl);
            fail = 1;
            break;
          }
          atts[natt++] = anm;
          atts[natt++] = avl;
        }
      }

      if (!fail && natt + 1 > attcap) {
        int ncap = attcap ? attcap + 1 : 1;
        char **na = (char **)realloc(atts, (size_t)ncap * sizeof(char *));
        if (!na)
          fail = 1;
        else
          atts = na;
      }

      if (fail) {
        for (int i = 0; i < natt; i++)
          free(atts[i]);
        free(atts);
        return -1;
      }
      atts[natt] = NULL;

      char *nm = mxml_ndup(nstart, nlen);
      if (!nm) {
        for (int i = 0; i < natt; i++)
          free(atts[i]);
        free(atts);
        return -1;
      }

      if (on_start)
        on_start(user, nm, (const XML_Char **)atts);
      if (self_closing && on_end)
        on_end(user, nm);

      free(nm);
      for (int i = 0; i < natt; i++)
        free(atts[i]);
      free(atts);

      p = q;
    }
  }

  return 0;
}

/*
** ============================================================================
** Utility Functions
** ============================================================================
*/

/*
** Parse a cell reference like "AB67" into column number and row number.
** Returns the column number (1-based) and sets *row to the row number
* (1-based).
*/
static int parse_cell_ref(const char *ref, int *row) {
  int col = 0;
  const char *p = ref;

  /* Parse column letters */
  while (*p && isalpha((unsigned char)*p)) {
    col = col * 26 + (toupper((unsigned char)*p) - 'A' + 1);
    p++;
  }

  /* Parse row number */
  if (row) {
    *row = atoi(p);
  }

  return col;
}

/*
** ============================================================================
** Shared Strings Parser
** ============================================================================
*/

typedef struct {
  char **strings;  /* Array of string values */
  int count;       /* Number of strings */
  int capacity;    /* Allocated capacity */
  int in_t;        /* Currently inside <t> element */
  char *current;   /* Current string being built */
  int current_len; /* Length of current string */
  int current_cap; /* Capacity of current string buffer */
} SharedStrings;

static void ss_init(SharedStrings *ss) { memset(ss, 0, sizeof(*ss)); }

static void ss_free(SharedStrings *ss) {
  for (int i = 0; i < ss->count; i++) {
    free(ss->strings[i]);
  }
  free(ss->strings);
  free(ss->current);
  memset(ss, 0, sizeof(*ss));
}

static void ss_add_string(SharedStrings *ss, const char *str) {
  if (ss->count >= ss->capacity) {
    int new_cap = ss->capacity ? ss->capacity * 2 : 64;
    char **new_strings = realloc(ss->strings, new_cap * sizeof(char *));
    if (!new_strings)
      return;
    ss->strings = new_strings;
    ss->capacity = new_cap;
  }
  ss->strings[ss->count++] = strdup(str ? str : "");
}

static void ss_append_text(SharedStrings *ss, const char *text, int len) {
  if (ss->current_len + len >= ss->current_cap) {
    int new_cap = ss->current_cap ? ss->current_cap * 2 : 256;
    while (new_cap <= ss->current_len + len)
      new_cap *= 2;
    char *new_current = realloc(ss->current, new_cap);
    if (!new_current)
      return;
    ss->current = new_current;
    ss->current_cap = new_cap;
  }
  memcpy(ss->current + ss->current_len, text, len);
  ss->current_len += len;
  ss->current[ss->current_len] = '\0';
}

static void XMLCALL ss_start_element(void *userData, const XML_Char *name,
                                     const XML_Char **atts) {
  SharedStrings *ss = (SharedStrings *)userData;
  (void)atts;

  if (strcmp(name, "si") == 0) {
    /* Start of a new string item - reset the accumulator. We do NOT reset on
    ** <t> so that rich-text runs (<si><r><t>..</t></r><r><t>..</t></r></si>)
    ** concatenate into a single value instead of keeping only the last run. */
    ss->current_len = 0;
    if (ss->current)
      ss->current[0] = '\0';
  } else if (strcmp(name, "t") == 0) {
    ss->in_t = 1;
  }
}

static void XMLCALL ss_end_element(void *userData, const XML_Char *name) {
  SharedStrings *ss = (SharedStrings *)userData;

  if (strcmp(name, "si") == 0) {
    /* End of string item - add accumulated text */
    ss_add_string(ss, ss->current ? ss->current : "");
    ss->current_len = 0;
    if (ss->current)
      ss->current[0] = '\0';
  } else if (strcmp(name, "t") == 0) {
    ss->in_t = 0;
  }
}

static void XMLCALL ss_char_data(void *userData, const XML_Char *s, int len) {
  SharedStrings *ss = (SharedStrings *)userData;

  if (ss->in_t) {
    ss_append_text(ss, s, len);
  }
}

static int parse_shared_strings(const char *xml, int xml_len,
                                SharedStrings *ss) {
  ss_init(ss);

  if (mini_xml_parse(xml, xml_len, ss, ss_start_element, ss_end_element,
                     ss_char_data) != 0) {
    ss_free(ss);
    return -1;
  }

  return 0;
}

/*
** ============================================================================
** Workbook Parser (Sheet Names)
** ============================================================================
*/

typedef struct {
  char *name;  /* Sheet name */
  int sheetId; /* Sheet ID */
  char *rid;   /* Relationship id (r:id), used to find the worksheet file */
} SheetInfo;

typedef struct {
  SheetInfo *sheets; /* Array of sheet info */
  int count;         /* Number of sheets */
  int capacity;      /* Allocated capacity */
} Workbook;

static void wb_init(Workbook *wb) { memset(wb, 0, sizeof(*wb)); }

static void wb_free(Workbook *wb) {
  for (int i = 0; i < wb->count; i++) {
    free(wb->sheets[i].name);
    free(wb->sheets[i].rid);
  }
  free(wb->sheets);
  memset(wb, 0, sizeof(*wb));
}

static void wb_add_sheet(Workbook *wb, const char *name, int sheetId,
                         const char *rid) {
  if (wb->count >= wb->capacity) {
    int new_cap = wb->capacity ? wb->capacity * 2 : 8;
    SheetInfo *new_sheets = realloc(wb->sheets, new_cap * sizeof(SheetInfo));
    if (!new_sheets)
      return;
    wb->sheets = new_sheets;
    wb->capacity = new_cap;
  }
  wb->sheets[wb->count].name = strdup(name ? name : "");
  wb->sheets[wb->count].sheetId = sheetId;
  wb->sheets[wb->count].rid = rid ? strdup(rid) : NULL;
  wb->count++;
}

static void XMLCALL wb_start_element(void *userData, const XML_Char *name,
                                     const XML_Char **atts) {
  Workbook *wb = (Workbook *)userData;

  if (strcmp(name, "sheet") == 0) {
    const char *sheet_name = NULL;
    const char *rid = NULL;
    int sheetId = 0;

    for (int i = 0; atts[i]; i += 2) {
      if (strcmp(atts[i], "name") == 0) {
        sheet_name = atts[i + 1];
      } else if (strcmp(atts[i], "sheetId") == 0) {
        sheetId = atoi(atts[i + 1]);
      } else if (strcmp(atts[i], "r:id") == 0) {
        rid = atts[i + 1];
      }
    }

    if (sheet_name) {
      wb_add_sheet(wb, sheet_name, sheetId, rid);
    }
  }
}

static void XMLCALL wb_end_element(void *userData, const XML_Char *name) {
  (void)userData;
  (void)name;
}

static int parse_workbook(const char *xml, int xml_len, Workbook *wb) {
  wb_init(wb);

  if (mini_xml_parse(xml, xml_len, wb, wb_start_element, wb_end_element,
                     NULL) != 0) {
    wb_free(wb);
    return -1;
  }

  return 0;
}

/*
** ============================================================================
** Workbook Relationships Parser (r:id -> worksheet file)
** ============================================================================
**
** The mapping from a sheet's r:id to its worksheet file lives in
** xl/_rels/workbook.xml.rels. The positional "sheet1.xml, sheet2.xml, ..."
** naming is a convention, not a guarantee, so we resolve the relationship
** target instead of assuming it.
*/

typedef struct {
  char *id;     /* Relationship id, e.g. "rId1" */
  char *target; /* Target path relative to xl/, e.g. "worksheets/sheet1.xml" */
} Relationship;

typedef struct {
  Relationship *items;
  int count;
  int capacity;
} Relationships;

static void rel_init(Relationships *r) { memset(r, 0, sizeof(*r)); }

static void rel_free(Relationships *r) {
  for (int i = 0; i < r->count; i++) {
    free(r->items[i].id);
    free(r->items[i].target);
  }
  free(r->items);
  memset(r, 0, sizeof(*r));
}

static void rel_add(Relationships *r, const char *id, const char *target) {
  if (r->count >= r->capacity) {
    int new_cap = r->capacity ? r->capacity * 2 : 8;
    Relationship *new_items = realloc(r->items, new_cap * sizeof(Relationship));
    if (!new_items)
      return;
    r->items = new_items;
    r->capacity = new_cap;
  }
  r->items[r->count].id = strdup(id ? id : "");
  r->items[r->count].target = strdup(target ? target : "");
  r->count++;
}

static void XMLCALL rel_start_element(void *userData, const XML_Char *name,
                                      const XML_Char **atts) {
  Relationships *r = (Relationships *)userData;

  if (strcmp(name, "Relationship") == 0) {
    const char *id = NULL;
    const char *target = NULL;
    for (int i = 0; atts[i]; i += 2) {
      if (strcmp(atts[i], "Id") == 0) {
        id = atts[i + 1];
      } else if (strcmp(atts[i], "Target") == 0) {
        target = atts[i + 1];
      }
    }
    if (id && target) {
      rel_add(r, id, target);
    }
  }
}

static int parse_relationships(const char *xml, int xml_len, Relationships *r) {
  rel_init(r);

  if (mini_xml_parse(xml, xml_len, r, rel_start_element, NULL, NULL) != 0) {
    rel_free(r);
    return -1;
  }

  return 0;
}

static const char *rel_find_target(Relationships *r, const char *id) {
  if (!id)
    return NULL;
  for (int i = 0; i < r->count; i++) {
    if (strcmp(r->items[i].id, id) == 0) {
      return r->items[i].target;
    }
  }
  return NULL;
}

/*
** Resolve a relationship Target into a full path within the zip archive.
** Targets in xl/_rels/workbook.xml.rels are relative to xl/ (e.g.
** "worksheets/sheet1.xml"); a target beginning with '/' is taken from the
** archive root. Returns a string allocated with sqlite3_mprintf(), or NULL.
*/
static char *resolve_worksheet_path(const char *target) {
  if (!target || !*target)
    return NULL;
  if (target[0] == '/') {
    return sqlite3_mprintf("%s", target + 1);
  }
  return sqlite3_mprintf("xl/%s", target);
}

/*
** ============================================================================
** Worksheet Parser
** ============================================================================
*/

typedef struct {
  char *value; /* Cell value (string or numeric as string) */
  int is_null; /* Whether this cell is empty/null */
} CellValue;

typedef struct {
  CellValue *cells; /* Row of cells */
  int count;        /* Number of cells in row */
  int capacity;     /* Allocated capacity */
} Row;

typedef struct {
  Row *rows;    /* Array of rows */
  int count;    /* Number of rows */
  int capacity; /* Allocated capacity */
  int max_col;  /* Maximum column number seen */
} Worksheet;

typedef struct {
  Worksheet *ws;     /* Worksheet being built */
  SharedStrings *ss; /* Shared strings reference */

  /* Current cell state */
  int cur_row;   /* Current row number (1-based) */
  int cur_col;   /* Current column number (1-based) */
  char cur_type; /* Cell type: 's'=shared string, 'n'=number, 'i'=inline,
                    'b'=boolean */
  int in_v;      /* Inside <v> element */
  int in_t;      /* Inside <t> element (for inline strings) */
  int in_is;     /* Inside <is> element (inline string container) */
  char *text;    /* Accumulated text */
  int text_len;  /* Length of accumulated text */
  int text_cap;  /* Capacity of text buffer */
} WorksheetParser;

static void ws_init(Worksheet *ws) { memset(ws, 0, sizeof(*ws)); }

static void ws_free(Worksheet *ws) {
  for (int i = 0; i < ws->count; i++) {
    for (int j = 0; j < ws->rows[i].count; j++) {
      free(ws->rows[i].cells[j].value);
    }
    free(ws->rows[i].cells);
  }
  free(ws->rows);
  memset(ws, 0, sizeof(*ws));
}

static Row *ws_get_row(Worksheet *ws, int row_num) {
  /* Ensure we have enough rows (row_num is 1-based) */
  while (ws->count < row_num) {
    if (ws->count >= ws->capacity) {
      int new_cap = ws->capacity ? ws->capacity * 2 : 64;
      Row *new_rows = realloc(ws->rows, new_cap * sizeof(Row));
      if (!new_rows)
        return NULL;
      ws->rows = new_rows;
      ws->capacity = new_cap;
    }
    memset(&ws->rows[ws->count], 0, sizeof(Row));
    ws->count++;
  }
  return &ws->rows[row_num - 1];
}

static void ws_set_cell(Worksheet *ws, int row_num, int col_num,
                        const char *value) {
  Row *row = ws_get_row(ws, row_num);
  if (!row)
    return;

  /* Ensure we have enough columns (col_num is 1-based) */
  while (row->count < col_num) {
    if (row->count >= row->capacity) {
      int new_cap = row->capacity ? row->capacity * 2 : 16;
      CellValue *new_cells = realloc(row->cells, new_cap * sizeof(CellValue));
      if (!new_cells)
        return;
      row->cells = new_cells;
      row->capacity = new_cap;
    }
    row->cells[row->count].value = NULL;
    row->cells[row->count].is_null = 1;
    row->count++;
  }

  /* Set the cell value */
  free(row->cells[col_num - 1].value);
  row->cells[col_num - 1].value = value ? strdup(value) : NULL;
  row->cells[col_num - 1].is_null = (value == NULL);

  /* Track max column */
  if (col_num > ws->max_col) {
    ws->max_col = col_num;
  }
}

static void wsp_init(WorksheetParser *wsp, Worksheet *ws, SharedStrings *ss) {
  memset(wsp, 0, sizeof(*wsp));
  wsp->ws = ws;
  wsp->ss = ss;
  wsp->cur_type = 'n'; /* Default to number */
}

static void wsp_free(WorksheetParser *wsp) { free(wsp->text); }

static void wsp_append_text(WorksheetParser *wsp, const char *s, int len) {
  if (wsp->text_len + len >= wsp->text_cap) {
    int new_cap = wsp->text_cap ? wsp->text_cap * 2 : 256;
    while (new_cap <= wsp->text_len + len)
      new_cap *= 2;
    char *new_text = realloc(wsp->text, new_cap);
    if (!new_text)
      return;
    wsp->text = new_text;
    wsp->text_cap = new_cap;
  }
  memcpy(wsp->text + wsp->text_len, s, len);
  wsp->text_len += len;
  wsp->text[wsp->text_len] = '\0';
}

static void XMLCALL ws_start_element(void *userData, const XML_Char *name,
                                     const XML_Char **atts) {
  WorksheetParser *wsp = (WorksheetParser *)userData;

  if (strcmp(name, "c") == 0) {
    /* Cell element */
    wsp->cur_type = 'n'; /* Default to number */
    wsp->cur_row = 0;
    wsp->cur_col = 0;

    for (int i = 0; atts[i]; i += 2) {
      if (strcmp(atts[i], "r") == 0) {
        /* Parse cell reference */
        wsp->cur_col = parse_cell_ref(atts[i + 1], &wsp->cur_row);
      } else if (strcmp(atts[i], "t") == 0) {
        /* Cell type */
        if (strcmp(atts[i + 1], "s") == 0) {
          wsp->cur_type = 's'; /* Shared string */
        } else if (strcmp(atts[i + 1], "inlineStr") == 0) {
          wsp->cur_type = 'i'; /* Inline string */
        } else if (strcmp(atts[i + 1], "b") == 0) {
          wsp->cur_type = 'b'; /* Boolean */
        } else if (strcmp(atts[i + 1], "str") == 0) {
          wsp->cur_type = 'f'; /* Formula string result */
        }
      }
    }

    /* Reset text buffer */
    wsp->text_len = 0;
    if (wsp->text)
      wsp->text[0] = '\0';
  } else if (strcmp(name, "v") == 0) {
    wsp->in_v = 1;
    wsp->text_len = 0;
    if (wsp->text)
      wsp->text[0] = '\0';
  } else if (strcmp(name, "is") == 0) {
    wsp->in_is = 1;
  } else if (strcmp(name, "t") == 0 && wsp->in_is) {
    wsp->in_t = 1;
    wsp->text_len = 0;
    if (wsp->text)
      wsp->text[0] = '\0';
  }
}

static void XMLCALL ws_end_element(void *userData, const XML_Char *name) {
  WorksheetParser *wsp = (WorksheetParser *)userData;

  if (strcmp(name, "c") == 0) {
    /* End of cell - store the value */
    if (wsp->cur_row > 0 && wsp->cur_col > 0) {
      const char *value = NULL;

      if (wsp->cur_type == 's' && wsp->text && wsp->ss) {
        /* Shared string - look up by index */
        int idx = atoi(wsp->text);
        if (idx >= 0 && idx < wsp->ss->count) {
          value = wsp->ss->strings[idx];
        }
      } else if (wsp->cur_type == 'i') {
        /* Inline string - use accumulated text */
        value = wsp->text;
      } else if (wsp->text && wsp->text_len > 0) {
        /* Number or other - use as-is */
        value = wsp->text;
      }

      ws_set_cell(wsp->ws, wsp->cur_row, wsp->cur_col, value);
    }
  } else if (strcmp(name, "v") == 0) {
    wsp->in_v = 0;
  } else if (strcmp(name, "is") == 0) {
    wsp->in_is = 0;
  } else if (strcmp(name, "t") == 0 && wsp->in_is) {
    wsp->in_t = 0;
  }
}

static void XMLCALL ws_char_data(void *userData, const XML_Char *s, int len) {
  WorksheetParser *wsp = (WorksheetParser *)userData;

  if (wsp->in_v || wsp->in_t) {
    wsp_append_text(wsp, s, len);
  }
}

static int parse_worksheet(const char *xml, int xml_len, SharedStrings *ss,
                           Worksheet *ws) {
  ws_init(ws);

  WorksheetParser wsp;
  wsp_init(&wsp, ws, ss);

  int result = mini_xml_parse(xml, xml_len, &wsp, ws_start_element,
                              ws_end_element, ws_char_data);

  wsp_free(&wsp);

  if (result != 0) {
    ws_free(ws);
  }

  return result;
}

/*
** ============================================================================
** Table Name Escaping
** ============================================================================
*/

/*
** Escape a sheet name to make it a valid SQLite identifier.
** Replaces problematic characters and wraps in quotes if necessary.
*/
static char *escape_identifier(const char *name) {
  if (!name || !*name) {
    return strdup("\"unnamed\"");
  }

  /* Calculate needed size (worst case: every char needs escaping) */
  int len = (int)strlen(name);
  char *escaped = malloc(len * 2 + 3); /* Extra for quotes and null */
  if (!escaped)
    return NULL;

  char *p = escaped;
  *p++ = '"';

  for (const char *s = name; *s; s++) {
    if (*s == '"') {
      *p++ = '"'; /* Double the quote */
    }
    *p++ = *s;
  }

  *p++ = '"';
  *p = '\0';

  return escaped;
}

/*
** ============================================================================
** Main Import Function
** ============================================================================
*/

/*
** Read a file from the XLSX archive using the zipfile extension.
*/
static int read_zip_entry(sqlite3 *db, const char *xlsx_path,
                          const char *entry_name, char **data, int *data_len) {
  char *sql = sqlite3_mprintf("SELECT data FROM zipfile(%Q) WHERE name = %Q",
                              xlsx_path, entry_name);
  if (!sql)
    return SQLITE_NOMEM;

  sqlite3_stmt *stmt = NULL;
  int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
  sqlite3_free(sql);

  if (rc != SQLITE_OK) {
    return rc;
  }

  rc = sqlite3_step(stmt);
  if (rc == SQLITE_ROW) {
    const void *blob = sqlite3_column_blob(stmt, 0);
    int blob_len = sqlite3_column_bytes(stmt, 0);

    *data = malloc(blob_len + 1);
    if (*data) {
      memcpy(*data, blob, blob_len);
      (*data)[blob_len] = '\0';
      *data_len = blob_len;
      rc = SQLITE_OK;
    } else {
      rc = SQLITE_NOMEM;
    }
  } else if (rc == SQLITE_DONE) {
    /* Entry not found - not an error, just return empty */
    *data = NULL;
    *data_len = 0;
    rc = SQLITE_OK;
  }

  sqlite3_finalize(stmt);
  return rc;
}

/*
** Compare two strings case-insensitively (ASCII). SQLite identifiers compare
** case-insensitively, so column names must be made unique under this rule.
*/
static int ci_equal(const char *a, const char *b) {
  while (*a && *b) {
    if (tolower((unsigned char)*a) != tolower((unsigned char)*b))
      return 0;
    a++;
    b++;
  }
  return *a == *b;
}

/*
** Build the column names for the table from the worksheet's first row. Blank
** header cells become "colN". A name that collides with an earlier column
** (case-insensitively) gets a numeric suffix so that CREATE TABLE does not
** fail on duplicate headers. Returns an array of ws->max_col strings (each
** allocated with sqlite3_malloc), or NULL on out-of-memory.
*/
static char **build_column_names(Worksheet *ws) {
  Row *first_row = &ws->rows[0];
  char **names = sqlite3_malloc(ws->max_col * sizeof(char *));
  if (!names)
    return NULL;
  memset(names, 0, ws->max_col * sizeof(char *));

  for (int col = 0; col < ws->max_col; col++) {
    const char *raw = NULL;
    if (col < first_row->count && first_row->cells[col].value)
      raw = first_row->cells[col].value;

    char *candidate = (raw && *raw) ? sqlite3_mprintf("%s", raw)
                                    : sqlite3_mprintf("col%d", col + 1);
    if (!candidate)
      goto oom;

    /* Resolve collisions with already-chosen column names. */
    int suffix = 2;
    int collides = 1;
    while (collides) {
      collides = 0;
      for (int k = 0; k < col; k++) {
        if (ci_equal(names[k], candidate)) {
          collides = 1;
          break;
        }
      }
      if (collides) {
        /* "%z" frees the previous candidate before formatting. */
        char *next = sqlite3_mprintf("%z_%d", candidate, suffix++);
        if (!next)
          goto oom;
        candidate = next;
      }
    }
    names[col] = candidate;
  }
  return names;

oom:
  for (int k = 0; k < ws->max_col; k++)
    sqlite3_free(names[k]);
  sqlite3_free(names);
  return NULL;
}

static void free_column_names(char **names, int count) {
  if (!names)
    return;
  for (int i = 0; i < count; i++)
    sqlite3_free(names[i]);
  sqlite3_free(names);
}

/*
** Create a table from a worksheet.
*/
static int create_table_from_worksheet(sqlite3 *db, const char *table_name,
                                       Worksheet *ws, char **pzErrMsg) {
  if (ws->count == 0 || ws->max_col == 0) {
    /* Empty worksheet */
    return SQLITE_OK;
  }

  /* Compute unique column names from the first row. */
  char **col_names = build_column_names(ws);
  if (!col_names) {
    return SQLITE_NOMEM;
  }

  char *escaped_table = escape_identifier(table_name);
  if (!escaped_table) {
    free_column_names(col_names, ws->max_col);
    return SQLITE_NOMEM;
  }

  /* Drop any existing table of the same name first, so that re-importing a
  ** file replaces its data instead of silently appending duplicate rows. */
  char *drop_sql = sqlite3_mprintf("DROP TABLE IF EXISTS %s", escaped_table);
  if (!drop_sql) {
    free(escaped_table);
    free_column_names(col_names, ws->max_col);
    return SQLITE_NOMEM;
  }
  int rc = sqlite3_exec(db, drop_sql, NULL, NULL, pzErrMsg);
  sqlite3_free(drop_sql);
  if (rc != SQLITE_OK) {
    free(escaped_table);
    free_column_names(col_names, ws->max_col);
    return rc;
  }

  /* Build CREATE TABLE statement. */
  sqlite3_str *sql = sqlite3_str_new(db);
  sqlite3_str_appendf(sql, "CREATE TABLE %s (", escaped_table);
  free(escaped_table);

  for (int col = 0; col < ws->max_col; col++) {
    if (col > 0)
      sqlite3_str_appendall(sql, ", ");

    char *escaped_col = escape_identifier(col_names[col]);
    if (escaped_col) {
      sqlite3_str_appendall(sql, escaped_col);
      free(escaped_col);
    } else {
      sqlite3_str_appendf(sql, "\"col%d\"", col + 1);
    }
  }

  sqlite3_str_appendall(sql, ")");
  free_column_names(col_names, ws->max_col);

  char *create_sql = sqlite3_str_finish(sql);
  if (!create_sql) {
    return SQLITE_NOMEM;
  }

  rc = sqlite3_exec(db, create_sql, NULL, NULL, pzErrMsg);
  sqlite3_free(create_sql);

  if (rc != SQLITE_OK) {
    return rc;
  }

  /* Build INSERT statement */
  sql = sqlite3_str_new(db);
  escaped_table = escape_identifier(table_name);
  sqlite3_str_appendf(sql, "INSERT INTO %s VALUES (", escaped_table);
  free(escaped_table);

  for (int col = 0; col < ws->max_col; col++) {
    if (col > 0)
      sqlite3_str_appendall(sql, ", ");
    sqlite3_str_appendall(sql, "?");
  }
  sqlite3_str_appendall(sql, ")");

  char *insert_sql = sqlite3_str_finish(sql);
  if (!insert_sql) {
    return SQLITE_NOMEM;
  }

  sqlite3_stmt *stmt = NULL;
  rc = sqlite3_prepare_v2(db, insert_sql, -1, &stmt, NULL);
  sqlite3_free(insert_sql);

  if (rc != SQLITE_OK) {
    return rc;
  }

  /* Insert data rows (skip first row which is headers) */
  for (int row_idx = 1; row_idx < ws->count; row_idx++) {
    Row *row = &ws->rows[row_idx];

    /* Check if row has any data */
    int has_data = 0;
    for (int col = 0; col < row->count; col++) {
      if (!row->cells[col].is_null) {
        has_data = 1;
        break;
      }
    }
    if (!has_data && row->count == 0)
      continue;

    sqlite3_reset(stmt);

    for (int col = 0; col < ws->max_col; col++) {
      if (col < row->count && !row->cells[col].is_null &&
          row->cells[col].value) {
        sqlite3_bind_text(stmt, col + 1, row->cells[col].value, -1,
                          SQLITE_TRANSIENT);
      } else {
        sqlite3_bind_null(stmt, col + 1);
      }
    }

    rc = sqlite3_step(stmt);
    if (rc != SQLITE_DONE) {
      sqlite3_finalize(stmt);
      return rc;
    }
  }

  sqlite3_finalize(stmt);
  return SQLITE_OK;
}

/*
** Helper function to check if a sheet should be imported based on the
** optional sheetname1..sheetnameN parameters.
** If no sheet parameters are provided (argc == 1), all sheets are imported.
** If sheet parameters are provided:
**   - Integer parameter: import sheet with that 1-based index
**   - String parameter: import sheet with that name
** Returns 1 if the sheet should be imported, 0 otherwise.
*/
static int should_import_sheet(int argc, sqlite3_value **argv,
                               int sheet_index,   /* 0-based index */
                               const char *sheet_name) {
  /* If no sheet parameters provided, import all sheets */
  if (argc <= 1) {
    return 1;
  }

  /* Check each sheet parameter (argv[1] through argv[argc-1]) */
  for (int i = 1; i < argc; i++) {
    int value_type = sqlite3_value_type(argv[i]);

    if (value_type == SQLITE_INTEGER) {
      /* Integer parameter: compare with 1-based sheet number */
      int sheet_num = sqlite3_value_int(argv[i]);
      if (sheet_num == sheet_index + 1) {
        return 1;
      }
    } else if (value_type == SQLITE_TEXT) {
      /* String parameter: compare with sheet name */
      const char *param_name = (const char *)sqlite3_value_text(argv[i]);
      if (param_name && sheet_name && strcmp(param_name, sheet_name) == 0) {
        return 1;
      }
    }
  }

  return 0;
}

/*
** xlsx_import(filename, [sheetname1, sheetname2, ...]) - Import sheets from
** an XLSX file as tables.
**
** Parameters:
**   filename    - Path to the XLSX file to import
**   sheetname1..sheetnameN - Optional sheet selectors (string name or integer
**                            number). If none provided, all sheets are
**                            imported. Integer parameters specify 1-based
**                            sheet numbers.
*/
static void xlsx_import_func(sqlite3_context *ctx, int argc,
                             sqlite3_value **argv) {
  if (argc < 1) {
    sqlite3_result_error(ctx, "xlsx_import requires a filename argument", -1);
    return;
  }

  const char *filename = (const char *)sqlite3_value_text(argv[0]);
  if (!filename) {
    sqlite3_result_error(ctx, "Invalid filename", -1);
    return;
  }

  sqlite3 *db = sqlite3_context_db_handle(ctx);
  int rc;
  char *errmsg = NULL;

  /* Read shared strings */
  SharedStrings ss;
  ss_init(&ss);

  char *ss_data = NULL;
  int ss_len = 0;
  rc = read_zip_entry(db, filename, "xl/sharedStrings.xml", &ss_data, &ss_len);
  if (rc != SQLITE_OK) {
    sqlite3_result_error(
        ctx, "Failed to read XLSX file (is zipfile extension loaded?)", -1);
    return;
  }

  if (ss_data && ss_len > 0) {
    if (parse_shared_strings(ss_data, ss_len, &ss) != 0) {
      free(ss_data);
      sqlite3_result_error(ctx, "Failed to parse shared strings", -1);
      return;
    }
  }
  free(ss_data);

  /* Read workbook to get sheet names */
  Workbook wb;
  wb_init(&wb);

  char *wb_data = NULL;
  int wb_len = 0;
  rc = read_zip_entry(db, filename, "xl/workbook.xml", &wb_data, &wb_len);
  if (rc != SQLITE_OK || !wb_data) {
    ss_free(&ss);
    sqlite3_result_error(ctx, "Failed to read workbook", -1);
    return;
  }

  if (parse_workbook(wb_data, wb_len, &wb) != 0) {
    free(wb_data);
    ss_free(&ss);
    sqlite3_result_error(ctx, "Failed to parse workbook", -1);
    return;
  }
  free(wb_data);

  /* Read workbook relationships so we can map each sheet's r:id to its actual
  ** worksheet file. Best-effort: if the rels are missing or a target cannot be
  ** resolved, we fall back to the positional sheetN.xml name below. */
  Relationships rels;
  rel_init(&rels);
  {
    char *rels_data = NULL;
    int rels_len = 0;
    int rels_rc = read_zip_entry(db, filename, "xl/_rels/workbook.xml.rels",
                                 &rels_data, &rels_len);
    if (rels_rc == SQLITE_OK && rels_data && rels_len > 0) {
      parse_relationships(rels_data, rels_len, &rels);
    }
    free(rels_data);
  }

  /* Process each sheet */
  int tables_created = 0;
  for (int i = 0; i < wb.count; i++) {
    /* Check if this sheet should be imported based on optional parameters */
    if (!should_import_sheet(argc, argv, i, wb.sheets[i].name)) {
      continue;
    }

    /* Resolve the worksheet file via the relationship target, falling back to
    ** the positional sheetN.xml name if the rels are missing or incomplete. */
    char *sheet_path = NULL;
    const char *target = rel_find_target(&rels, wb.sheets[i].rid);
    if (target) {
      sheet_path = resolve_worksheet_path(target);
    }
    if (!sheet_path) {
      sheet_path = sqlite3_mprintf("xl/worksheets/sheet%d.xml", i + 1);
    }
    if (!sheet_path) {
      continue; /* out of memory; skip this sheet */
    }

    char *sheet_data = NULL;
    int sheet_len = 0;
    rc = read_zip_entry(db, filename, sheet_path, &sheet_data, &sheet_len);
    sqlite3_free(sheet_path);

    if (rc != SQLITE_OK || !sheet_data || sheet_len == 0) {
      free(sheet_data);
      continue;
    }

    Worksheet ws;
    if (parse_worksheet(sheet_data, sheet_len, &ss, &ws) != 0) {
      free(sheet_data);
      continue;
    }
    free(sheet_data);

    rc = create_table_from_worksheet(db, wb.sheets[i].name, &ws, &errmsg);
    ws_free(&ws);

    if (rc != SQLITE_OK) {
      wb_free(&wb);
      ss_free(&ss);
      rel_free(&rels);
      if (errmsg) {
        sqlite3_result_error(ctx, errmsg, -1);
        sqlite3_free(errmsg);
      } else {
        sqlite3_result_error(ctx, "Failed to create table", -1);
      }
      return;
    }

    tables_created++;
  }

  wb_free(&wb);
  ss_free(&ss);
  rel_free(&rels);

  sqlite3_result_int(ctx, tables_created);
}

/*
** xlsx_import_version() - Return the version string.
*/
static void xlsx_import_version_func(sqlite3_context *ctx, int argc,
                                     sqlite3_value **argv) {
  (void)argc;
  (void)argv;
  sqlite3_result_text(ctx, "2026-01-07 Claude Opus 4.5 (Thinking)", -1,
                      SQLITE_STATIC);
}

/*
** ============================================================================
** Table-Valued Function: xlsx_import_sheetnames
** ============================================================================
**
** Returns the sheet names from an XLSX file as a table with columns:
**   sheet_num  - 1-based sheet number
**   sheet_name - Name of the sheet
**
** Usage:
**   SELECT * FROM xlsx_import_sheetnames('filename.xlsx');
*/

/* Virtual table cursor structure */
typedef struct sheetnames_cursor {
  sqlite3_vtab_cursor base; /* Base class - must be first */
  Workbook wb;              /* Parsed workbook with sheet info */
  int current;              /* Current row index (0-based) */
  int eof;                  /* True if past last row */
} sheetnames_cursor;

/* Virtual table structure */
typedef struct sheetnames_vtab {
  sqlite3_vtab base; /* Base class - must be first */
  sqlite3 *db;       /* Database connection */
} sheetnames_vtab;

/* xConnect/xCreate - Create a new virtual table instance */
static int sheetnamesConnect(sqlite3 *db, void *pAux, int argc,
                             const char *const *argv, sqlite3_vtab **ppVtab,
                             char **pzErr) {
  (void)pAux;
  (void)argc;
  (void)argv;
  (void)pzErr;

  int rc = sqlite3_declare_vtab(
      db, "CREATE TABLE x(sheet_num INTEGER, sheet_name TEXT, "
          "filename HIDDEN)");
  if (rc != SQLITE_OK) {
    return rc;
  }

  sheetnames_vtab *pNew = sqlite3_malloc(sizeof(*pNew));
  if (!pNew) {
    return SQLITE_NOMEM;
  }
  memset(pNew, 0, sizeof(*pNew));
  pNew->db = db;

  *ppVtab = &pNew->base;
  return SQLITE_OK;
}

/* xDisconnect/xDestroy - Destroy virtual table instance */
static int sheetnamesDisconnect(sqlite3_vtab *pVtab) {
  sqlite3_free(pVtab);
  return SQLITE_OK;
}

/* xOpen - Create a new cursor */
static int sheetnamesOpen(sqlite3_vtab *pVtab, sqlite3_vtab_cursor **ppCursor) {
  (void)pVtab;

  sheetnames_cursor *pCur = sqlite3_malloc(sizeof(*pCur));
  if (!pCur) {
    return SQLITE_NOMEM;
  }
  memset(pCur, 0, sizeof(*pCur));

  *ppCursor = &pCur->base;
  return SQLITE_OK;
}

/* xClose - Close and free a cursor */
static int sheetnamesClose(sqlite3_vtab_cursor *cur) {
  sheetnames_cursor *pCur = (sheetnames_cursor *)cur;
  wb_free(&pCur->wb);
  sqlite3_free(pCur);
  return SQLITE_OK;
}

/* xNext - Advance cursor to next row */
static int sheetnamesNext(sqlite3_vtab_cursor *cur) {
  sheetnames_cursor *pCur = (sheetnames_cursor *)cur;
  pCur->current++;
  if (pCur->current >= pCur->wb.count) {
    pCur->eof = 1;
  }
  return SQLITE_OK;
}

/* xColumn - Return value for column iCol of current row */
static int sheetnamesColumn(sqlite3_vtab_cursor *cur, sqlite3_context *ctx,
                            int iCol) {
  sheetnames_cursor *pCur = (sheetnames_cursor *)cur;

  if (pCur->current >= pCur->wb.count) {
    sqlite3_result_null(ctx);
    return SQLITE_OK;
  }

  switch (iCol) {
  case 0: /* sheet_num (1-based) */
    sqlite3_result_int(ctx, pCur->current + 1);
    break;
  case 1: /* sheet_name */
    sqlite3_result_text(ctx, pCur->wb.sheets[pCur->current].name, -1,
                        SQLITE_TRANSIENT);
    break;
  default:
    sqlite3_result_null(ctx);
    break;
  }
  return SQLITE_OK;
}

/* xRowid - Return rowid for current row */
static int sheetnamesRowid(sqlite3_vtab_cursor *cur, sqlite3_int64 *pRowid) {
  sheetnames_cursor *pCur = (sheetnames_cursor *)cur;
  *pRowid = pCur->current + 1;
  return SQLITE_OK;
}

/* xEof - Return true if cursor is past last row */
static int sheetnamesEof(sqlite3_vtab_cursor *cur) {
  sheetnames_cursor *pCur = (sheetnames_cursor *)cur;
  return pCur->eof;
}

/* xFilter - Begin a search; parse the XLSX file */
static int sheetnamesFilter(sqlite3_vtab_cursor *cur, int idxNum,
                            const char *idxStr, int argc,
                            sqlite3_value **argv) {
  (void)idxNum;
  (void)idxStr;

  sheetnames_cursor *pCur = (sheetnames_cursor *)cur;
  sheetnames_vtab *pVtab = (sheetnames_vtab *)cur->pVtab;

  /* Free any previous workbook data */
  wb_free(&pCur->wb);
  pCur->current = 0;
  pCur->eof = 0;

  if (argc < 1) {
    pVtab->base.zErrMsg =
        sqlite3_mprintf("xlsx_import_sheetnames requires a filename argument");
    return SQLITE_ERROR;
  }

  const char *filename = (const char *)sqlite3_value_text(argv[0]);
  if (!filename) {
    pVtab->base.zErrMsg = sqlite3_mprintf("Invalid filename");
    return SQLITE_ERROR;
  }

  /* Read workbook.xml to get sheet names */
  char *wb_data = NULL;
  int wb_len = 0;
  int rc =
      read_zip_entry(pVtab->db, filename, "xl/workbook.xml", &wb_data, &wb_len);
  if (rc != SQLITE_OK || !wb_data) {
    pVtab->base.zErrMsg = sqlite3_mprintf("Failed to read workbook from %s", filename);
    return SQLITE_ERROR;
  }

  if (parse_workbook(wb_data, wb_len, &pCur->wb) != 0) {
    free(wb_data);
    pVtab->base.zErrMsg = sqlite3_mprintf("Failed to parse workbook");
    return SQLITE_ERROR;
  }
  free(wb_data);

  /* Check if we have any sheets */
  if (pCur->wb.count == 0) {
    pCur->eof = 1;
  }

  return SQLITE_OK;
}

/* xBestIndex - Determine query plan; require filename parameter */
static int sheetnames_BestIndex(sqlite3_vtab *pVtab,
                                sqlite3_index_info *pIdxInfo) {
  (void)pVtab;

  /* Look for equality constraint on filename (column 2, the hidden column) */
  int filenameIdx = -1;
  for (int i = 0; i < pIdxInfo->nConstraint; i++) {
    if (pIdxInfo->aConstraint[i].usable &&
        pIdxInfo->aConstraint[i].iColumn == 2 &&
        pIdxInfo->aConstraint[i].op == SQLITE_INDEX_CONSTRAINT_EQ) {
      filenameIdx = i;
      break;
    }
  }

  if (filenameIdx < 0) {
    /* Filename is required */
    return SQLITE_CONSTRAINT;
  }

  pIdxInfo->aConstraintUsage[filenameIdx].argvIndex = 1;
  pIdxInfo->aConstraintUsage[filenameIdx].omit = 1;
  pIdxInfo->estimatedCost = 1000.0;
  pIdxInfo->estimatedRows = 10;

  return SQLITE_OK;
}

/* Virtual table module definition */
static sqlite3_module sheetnamesModule = {
    0,                    /* iVersion */
    sheetnamesConnect,    /* xCreate */
    sheetnamesConnect,    /* xConnect */
    sheetnames_BestIndex, /* xBestIndex */
    sheetnamesDisconnect, /* xDisconnect */
    sheetnamesDisconnect, /* xDestroy */
    sheetnamesOpen,       /* xOpen */
    sheetnamesClose,      /* xClose */
    sheetnamesFilter,     /* xFilter */
    sheetnamesNext,       /* xNext */
    sheetnamesEof,        /* xEof */
    sheetnamesColumn,     /* xColumn */
    sheetnamesRowid,      /* xRowid */
    0,                    /* xUpdate */
    0,                    /* xBegin */
    0,                    /* xSync */
    0,                    /* xCommit */
    0,                    /* xRollback */
    0,                    /* xFindFunction */
    0,                    /* xRename */
    0,                    /* xSavepoint */
    0,                    /* xRelease */
    0,                    /* xRollbackTo */
    0,                    /* xShadowName */
    0                     /* xIntegrity */
};

/*
** ============================================================================
** Extension Entry Point
** ============================================================================
*/

#ifdef _WIN32
__declspec(dllexport)
#endif
int sqlite3_xlsximport_init(sqlite3 *db, char **pzErrMsg,
                            const sqlite3_api_routines *pApi) {
  SQLITE_EXTENSION_INIT2(pApi);
  (void)pzErrMsg;

  /* Register xlsx_import with -1 for nArg to accept variable number of
   * arguments (filename plus optional sheet selectors) */
  int rc = sqlite3_create_function(db, "xlsx_import", -1,
                                   SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL,
                                   xlsx_import_func, NULL, NULL);
  if (rc != SQLITE_OK)
    return rc;

  rc = sqlite3_create_function(db, "xlsx_import_version", 0,
                               SQLITE_UTF8 | SQLITE_DETERMINISTIC, NULL,
                               xlsx_import_version_func, NULL, NULL);
  if (rc != SQLITE_OK)
    return rc;

  /* Register xlsx_import_sheetnames as an eponymous table-valued function */
  rc = sqlite3_create_module(db, "xlsx_import_sheetnames", &sheetnamesModule,
                             NULL);

  return rc;
}


