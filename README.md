# zexmlss

Delphi/Free Pascal component library for reading and writing spreadsheet files without requiring Excel or LibreOffice.

**Supported formats:** Excel XML (.xml), Office Open XML (.xlsx/.xlsm), OpenDocument (.ods/.fods)

**Original author:** Ruslan V. Neborak ([avemey.com](http://avemey.com))

**License:** zlib

**Fork lineage:** [Avemey/zexmlss](https://github.com/Avemey/zexmlss) -> [serbod/zexmlss](https://github.com/serbod/zexmlss) -> this fork

## Compiler Support

- Delphi XE2 through 12.3 Athens (tested and compiled on Delphi 12.3)
- Free Pascal / Lazarus

## Changes in This Fork

### Zip Backend Cleanup

The original library shipped with 6 alternative zip backends (KaZip, Abbrevia, JCL 7-zip, SciZip, Synzip, ZipMaster) plus a pluggable abstraction layer (`zeZippy.pas`). This fork removes all of them and unifies on a single zip strategy per platform:

- **Delphi (XE2+):** `zeZipper.pas` wrapping `System.Zip` (auto-detected via `{$define XE2ZIP}`)
- **Free Pascal:** RTL `zipper` unit directly

Both XLSX and ODS modules now use the same zip pattern.

**Removed files:** `zeZippy.pas`, `zeZippyXE2.pas`, `zeZippyAB.pas`, `zeZippyJCL7z.pas`, `zeZippyZipMaster.pas`, `zeZippyLazip.pp`, all zip include stubs (`odszipuses.inc`, `xlsxzipuses.inc`, etc.), and the entire `delphizip/` directory.

### Fluent Save API Removed

The `IZXMLSSave` fluent interface (`zeSave.pas`, `zeSaveEXML.pas`, `zeSaveODS.pas`, `zeSaveXLSX.pas`) has been removed. Use the format-specific export functions directly:

```pascal
SaveXmlssToODFS(XMLSS, 'output.ods');
SaveXmlssToXLSX(XMLSS, 'output.xlsx');
SaveXmlssToEXML(XMLSS, 'output.xml', [0], [], @TextConverter, 'utf-8');
```

### Package Files Updated

- **Delphi 12:** `zexmlss/packages/delphi12/zexmlsslib.dpk` and `.dproj` recreated for Delphi 12.3 Athens (Win32 + Win64)
- **Lazarus:** `zexmlss/packages/lazarus/zexmlsslib.lpk` updated to remove deleted units

### Linux64 Cross-Platform Support

The original library assumes Delphi = Windows and uses `{$IFNDEF FPC}` guards for Windows-specific code. This fork adds Delphi Linux64 support:

**Compatibility shims** (`zexmlss/compat/`):

- `Graphics.pas` — Provides `TColor`, `TFont`, `TFontStyles`, `ColorToRGB`, and VCL color constants using `System.UITypes` and `System.Types`. Replaces `Vcl.Graphics` on non-Windows targets.
- `windows.pas` — Provides `HWND`, `TRect`, `TPoint`, `TSize` type aliases and a `GetDeviceCaps` stub. Replaces `Winapi.Windows` on non-Windows targets.

These units must only be on the search path for non-Windows targets. On Windows, the unqualified `Graphics` and `windows` resolve to their VCL/WinAPI counterparts via namespace search.

**Source fixes:**

- `zeZipper.pas` — Changed `FOnPercent` field type from `LongInt` to `Integer` (incompatible types on 64-bit Linux where `LongInt` is 64-bit)
- `zearchhelper.pas` — Changed `{$ifndef FPC}` guard around `Windows.CopyFile` to `{$ifdef MSWINDOWS}` so the stream-based fallback is used on Delphi Linux64

**Project integration example** (`.dproj`):

```xml
<!-- Win32: add Vcl to namespace search so 'Graphics' resolves to Vcl.Graphics -->
<PropertyGroup Condition="'$(Platform)'=='Win32'">
  <DCC_Namespace>Winapi;System.Win;Data.Win;...;Vcl;$(DCC_Namespace)</DCC_Namespace>
</PropertyGroup>

<!-- Linux64: add compat directory to unit search path -->
<PropertyGroup Condition="'$(Platform)'=='Linux64'">
  <DCC_UnitSearchPath>path\to\zexmlss\compat;$(DCC_UnitSearchPath)</DCC_UnitSearchPath>
</PropertyGroup>
```

### Bug Fix

- Fixed `ForceDirectories('')` crash in `zeZipper.pas` when saving to a filename without a directory path (e.g. `SaveXmlssToODFS(XMLSS, 'output.ods')`)

## Console Example

`zexmlss/examples/odsconsole/` contains a console application that reads participant data from a JSON file and generates an ODS spreadsheet with styles, formulas, borders, landscape orientation, and 75% print scaling.

```
odsconsole <input.json> [output.ods]
```

### External Dependency

The console example requires [JsonDataObjects](https://github.com/AHausladen/JsonDataObjects) for JSON parsing. The unit is referenced via relative path from `dmvcframework/sources/`:

```
..\..\..\..\..\dmvcframework\sources\JsonDataObjects.pas
```

Adjust this path in `odsconsole.dpr` if your copy of JsonDataObjects is located elsewhere. The library itself (zexmlsslib package) has no external dependencies.

## Building

### Package (library only)

Open `zexmlss/packages/delphi12/zexmlsslib.dproj` in Delphi 12 and compile. No external dependencies required.

### Console Example

Open `zexmlss/examples/odsconsole/odsconsole.dproj`. Ensure `JsonDataObjects.pas` is reachable via the unit search path, then compile and run:

```
odsconsole response.json output.ods
```

## Project Structure

```
zexmlss/
  src/              Source units (zexmlss, zeodfs, zexlsx, zeZipper, ...)
  compat/           Cross-platform shims for non-Windows targets (Graphics, windows)
  packages/
    delphi12/       Delphi 12 package (.dpk/.dproj)
    lazarus/        Lazarus package (.lpk)
  examples/
    odsconsole/     Console ODS generator (JSON -> ODS)
    createexml/     GUI example for Excel XML export
zcolorstringgrid/   VCL TStringGrid extension with per-cell styling
```
