# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**zexmlss** (Z Excel XML SpreadSheet) is a Delphi component library for reading and writing spreadsheet files without requiring Excel or LibreOffice. It supports Excel XML (.xml), Office Open XML (.xlsx/.xlsm), and OpenDocument (.ods/.fods) formats.

- **License**: zlib
- **Author**: Ruslan V. Neborak
- **Supported compilers**: Delphi 5 through 12+, C++Builder 6, Free Pascal/Lazarus

## Build

The library is packaged as Delphi design-time packages. The active Delphi 12 packages are in `zexmlss/packages/delphi12/`:

- `zexmlsslib.dpk` / `.dproj` — Core library (without ZColorStringGrid)
- `zexmlsslibe.dpk` / `.dproj` — Extended library (with ZColorStringGrid VCL component)

Use the delphi-build MCP server to compile:
```
compile_delphi_project("zexmlss/packages/delphi12/zexmlsslib.dproj")
```

There are no automated tests. The `zexmlss/examples/` directory contains integration examples (createexml, openexml, frozentst, odsconsole).

## Architecture

### Core Object Model (`zexmlss/src/zexmlss.pas`)

```
TZEXMLSS                          — Root workbook object
├── Sheets: TZSheets              — Collection of worksheets
│   └── TZSheet
│       ├── Cell[col,row]: TZCell — Cell data, formula, type, style ref
│       ├── ColOptions[]: TZColOptions  — Column width, hidden, style
│       ├── RowOptions[]: TZRowOptions  — Row height, hidden, style
│       ├── MergeCells: TZMergeCells
│       └── SheetOptions: TZSheetOptions (margins, freeze/split, print settings)
└── Styles: TZStyles              — Collection of cell styles
    └── TZStyle (font, borders, alignment, number format, fill pattern)
```

Key types: `TZCellType` (ZENumber, ZEDateTime, ZEBoolean, ZEString, ZEError, ZEGeneral), `TZSplitMode` (ZSplitNone, ZSplitFrozen, ZSplitSplit).

### Format Modules

| Unit | Format | Notes |
|------|--------|-------|
| `zexmlssutils.pas` | Excel XML (.xml) | `SaveXmlssToEXML()`, `ReadEXMLSS()`. Also has Grid↔XMLSS converters and HTML export |
| `zexlsx.pas` | XLSX (.xlsx/.xlsm) | `SaveXmlssToXLSX()`, `ReadXLSX()`. Requires zip library |
| `zeodfs.pas` | ODS (.ods/.fods) | `SaveXmlssToODFS()`, `ReadODFS()`. Requires zip library |

### Zip Layer

XLSX and ODS both use the same zip pattern:
- **Delphi**: `zeZipper.pas` — wraps `System.Zip` (TZipper/TUnZipper mimicking FPC's API). Auto-detected for Delphi XE2+ via `{$define XE2ZIP}` in `zexml.inc`.
- **FPC/Lazarus**: RTL `zipper` unit directly.

### Supporting Units

- `zsspxml.pas` — Custom XML reader/writer (SAX-like, no DOM dependency)
- `zeformula.pas` — Formula notation conversion (A1 ↔ R1C1)
- `zenumberformats.pas` — Number format string parsing and cross-format conversion
- `zearchhelper.pas` — Archive manipulation helpers
- `compver.inc` — Compiler version detection macros

## Conditional Compilation (`zexml.inc`)

| Define | Effect |
|--------|--------|
| `NOZCOLORSTRINGGRID` | Exclude ZColorStringGrid dependency (enabled by default) |
| `ZUSE_CONDITIONAL_FORMATTING` | Enable conditional formatting support |
| `ZUSE_CHARTS` | Enable chart support |
| `ZUSE_DRAWINGS` | Enable drawing/image support |
| `NOXMLINDENTS` | Compact XML output without pretty-printing |
| `XE2ZIP` | Use System.Zip (auto-detected for Delphi XE2+) |

## ZColorStringGrid (`zcolorstringgrid/`)

Separate VCL component extending TStringGrid with per-cell styling, merged cells, text rotation, and auto-fit. Source in `zcolorstringgrid/src/`. Packages in `zcolorstringgrid/package/delphi11/`.

## Code Conventions

- Comments are mixed Russian/English; prefer English for new code
- Formulas are stored internally in R1C1 notation; conversion is in `zeformula.pas`
- The three large format units (`zexlsx.pas` ~247KB, `zeodfs.pas` ~248KB, `zexmlss.pas` ~178KB) contain complete read/write logic for their respective formats
- Known format quirks are documented in `zexmlss/Format Quirks.txt`
