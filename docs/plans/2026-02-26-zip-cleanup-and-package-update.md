# Zip Cleanup, Package Update, and Console Example Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Remove all alternative zip backends, unify ODS to use the same zip pattern as XLSX, update Delphi 12 package files, and create a console example for ODS generation.

**Architecture:** ODS adopts the same `zeZipper.pas` (Delphi) / `zipper` (FPC) pattern that XLSX already uses. The `{$IFDEF FPC}` blocks around `TODFZipHelper`, `ReadODFS`, and `SaveXmlssToODFS` are removed so these functions become cross-platform. The pluggable `zeZippy` abstraction and fluent `zeSave` API are deleted entirely.

**Tech Stack:** Delphi 12.2 Athens, Free Pascal / Lazarus, System.Zip (Delphi XE2+), RTL zipper (FPC)

---

### Task 1: Clean zexml.inc — remove alternative zip defines

**Files:**
- Modify: `zexmlss/src/zexml.inc`

**Step 1: Edit zexml.inc**

Remove the `{$define KAZIP}` line (line 31) and all commented-out zip defines (lines 33-43). Keep only the XE2ZIP auto-detection block (lines 23-28).

The archivers section should become:
```pascal
// Archivers for Delphi, not used for Free Pascal / Lazarus
{ use Delphi XE2 and above - System.Zip }
{$ifndef FPC}
{$if CompilerVersion >= 23.0}         // bds xe2 (2012)
  {$define XE2ZIP}
{$ifend}
{$endif}
```

**Step 2: Verify no other files define KAZIP**

Run: `grep -r "define KAZIP\|define JCL7Z\|define ABZIP\|define SCIZIP\|define SYNZIP" zexmlss/src/`
Expected: No matches (only zexml.inc had them).

**Step 3: Commit**

```bash
git add zexmlss/src/zexml.inc
git commit -m "Remove alternative zip backend defines from zexml.inc"
```

---

### Task 2: Clean zeZipper.pas — remove non-XE2ZIP backends

**Files:**
- Modify: `zexmlss/src/zeZipper.pas`

**Step 1: Remove alternative backend code**

In the uses clause (lines 14-21), remove lines 17-20:
```pascal
  {$ifdef KAZIP}, KAZip{$endif}
  {$ifdef JCL7Z}, JclCompression{$endif}
  {$ifdef ABZIP}, AbZipper, AbUnzper, AbArcTyp, AbUtils{$endif}
  {$ifdef SCIZIP}, SciZipFile{$endif}
```

In TZipper class (around line 130), remove:
```pascal
    {$ifdef KAZIP}procedure BuildZipDirectoryKaZip();{$endif}
    {$ifdef JCL7Z}procedure BuildZipDirectoryJCL7Z();{$endif}
    {$ifdef ABZIP}procedure BuildZipDirectoryAbZip();{$endif}
    {$ifdef SCIZIP}procedure BuildZipDirectorySciZip();{$endif}
```

In TUnZipper class (around line 191), remove:
```pascal
    {$ifdef KAZIP}procedure ReadKaZip(AExtract: Boolean);{$endif}
    {$ifdef JCL7Z}procedure ReadJCL7Z(AExtract: Boolean);{$endif}
```

In the implementation section, remove all `{$ifdef KAZIP}...{$endif}`, `{$ifdef JCL7Z}...{$endif}`, `{$ifdef ABZIP}...{$endif}`, `{$ifdef SCIZIP}...{$endif}` blocks — these are entire procedure implementations (BuildZipDirectoryKaZip, BuildZipDirectoryJCL7Z, etc.).

In the `{$elseif}` chains (around lines 432 and 859), remove the alternative branches, keeping only XE2ZIP as the primary and raising an error if undefined:
```pascal
{$ifdef XE2ZIP}
  // XE2ZIP implementation
{$else}
  {$error No zip backend defined. Requires Delphi XE2 or later.}
{$endif}
```

**Step 2: Verify it compiles**

The file should now only reference `System.Zip` (via `{$ifdef XE2ZIP}`).

**Step 3: Commit**

```bash
git add zexmlss/src/zeZipper.pas
git commit -m "Remove KAZIP, JCL7Z, ABZIP, SCIZIP backends from zeZipper"
```

---

### Task 3: Unify zeodfs.pas — replace zeZippy with zeZipper

This is the core change. The ODS module currently has three code paths: folder output, FPC-only zip, and Delphi zeZippy zip. We unify to one cross-platform path using zeZipper/zipper.

**Files:**
- Modify: `zexmlss/src/zeodfs.pas`

**Step 1: Fix uses clause (line 55)**

Replace:
```pascal
  zeZippy
  {$IFDEF FPC},zipper {$ELSE}{$I odszipuses.inc}{$ENDIF};
```

With:
```pascal
  {$IFDEF FPC}zipper{$ELSE}zeZipper{$ENDIF};
```

Note: `zeZippy` is completely removed from uses.

**Step 2: Remove `{$IFDEF FPC}` wrapper around TODFZipHelper (lines 533-608)**

Remove the `{$IFDEF FPC}` at line 533 and the `{$ENDIF}` at line 608. The `TODFZipHelper` class, its methods, and the `ReadODFStyles` forward declaration become available on all platforms.

Note: `TODFZipHelper` uses `TFullZipFileEntry` which is defined in both FPC's `zipper` unit and in `zeZipper.pas` for Delphi. The `@` operator for method references in the FPC ReadODFS function (line 929-930) needs to be adapted:
```pascal
// FPC uses @ operator for method references
u_zip.OnCreateStream := @ZH.DoCreateOutZipStream;  // FPC
u_zip.OnCreateStream := ZH.DoCreateOutZipStream;    // Delphi
```
Use conditional compilation:
```pascal
u_zip.OnCreateStream := {$IFDEF FPC}@{$ENDIF}ZH.DoCreateOutZipStream;
u_zip.OnDoneStream := {$IFDEF FPC}@{$ENDIF}ZH.DoDoneOutZipStream;
```

**Step 3: Remove `{$IFDEF FPC}` wrapper around SaveXmlssToODFS functions**

Remove the `{$IFDEF FPC}` at the line before SaveXmlssToODFS declarations in the interface (line 302) and the `{$ENDIF}` (line 308). These function declarations become cross-platform.

Remove the `{$IFDEF FPC}` before the SaveXmlssToODFS implementations (before line 5235) and the `{$ENDIF}` at line 5343. These implementations become cross-platform.

**Step 4: Add mimetype to SaveXmlssToODFS**

The FPC SaveXmlssToODFS (line 5235) is missing the `mimetype` file. Add it before the other entries:
```pascal
    StreamMT := TMemoryStream.Create();
    mime := AnsiString('application/vnd.oasis.opendocument.spreadsheet');
    StreamMT.WriteBuffer(mime[1], Length(mime));

    // ... existing stream creation code ...

    StreamMT.Position := 0;
    zip.Entries.AddFileEntry(StreamMT, 'mimetype');
    // ... rest of entries ...
```

Add `StreamMT: TStream;` and `mime: AnsiString;` to the var block.

**Step 5: Remove `{$IFDEF FPC}` wrapper around ReadODFS (lines 6896-6963)**

Remove the `{$IFDEF FPC}` at line 6896 and the `{$ENDIF}` at line 6963.

Also remove from the interface section: the `{$IFDEF FPC}` at line 318 and `{$ENDIF}` at line 320.

Adapt the `@` operator for Delphi compatibility (same as Step 2):
```pascal
u_zip.OnCreateStream := {$IFDEF FPC}@{$ENDIF}ZH.DoCreateOutZipStream;
u_zip.OnDoneStream := {$IFDEF FPC}@{$ENDIF}ZH.DoDoneOutZipStream;
```

**Step 6: Replace ExportXmlssToODFS with the unified version**

The old `ExportXmlssToODFS` (lines 5345-5418) uses `zeZippy` types (`TZxZipGen`, `CZxZipGens`, `EZxZipGen`). Replace it entirely.

New signature (remove ZipGenerator parameter):
```pascal
function ExportXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string;
  const SheetsNumbers: array of integer;
  const SheetsNames: array of string;
  TextConverter: TAnsiToCPConverter; CodePageName: String;
  BOM: ansistring = '';
  AllowUnzippedFolder: boolean = false): integer; overload;
```

New implementation — delegates to SaveXmlssToODFSPath when AllowUnzippedFolder is true, otherwise delegates to SaveXmlssToODFS:
```pascal
function ExportXmlssToODFS(var XMLSS: TZEXMLSS; FileName: string;
  const SheetsNumbers: array of integer;
  const SheetsNames: array of string;
  TextConverter: TAnsiToCPConverter; CodePageName: String;
  BOM: ansistring = '';
  AllowUnzippedFolder: boolean = false): integer; overload;
begin
  if AllowUnzippedFolder then
    Result := SaveXmlssToODFSPath(XMLSS, FileName, SheetsNumbers,
      SheetsNames, TextConverter, CodePageName, BOM)
  else
    Result := SaveXmlssToODFS(XMLSS, FileName, SheetsNumbers,
      SheetsNames, TextConverter, CodePageName, BOM);
end;
```

Update the interface declaration to match (remove ZipGenerator parameter).

**Step 7: Remove include file references**

Remove line 322-324:
```pascal
{$IFNDEF FPC}
{$I odszipfunc.inc}
{$ENDIF}
```

Remove lines 7003-7005:
```pascal
{$IFNDEF FPC}
{$I odszipfuncimpl.inc}
{$ENDIF}
```

**Step 8: Commit**

```bash
git add zexmlss/src/zeodfs.pas
git commit -m "Unify ODS zip handling to use zeZipper/zipper like XLSX"
```

---

### Task 4: Delete removed files

**Files to delete:**

From `zexmlss/src/`:
- `zeZippy.pas`
- `zeZippyXE2.pas`
- `zeZippyAB.pas`
- `zeZippyJCL7z.pas`
- `zeZippyZipMaster.pas`
- `zeZippyLazip.pp`
- `zeSave.pas`
- `zeSaveEXML.pas`
- `zeSaveODS.pas`
- `zeSaveXLSX.pas`
- `odszipuses.inc`
- `odszipfunc.inc`
- `odszipfuncimpl.inc`
- `xlsxzipuses.inc`
- `xlsxzipfunc.inc`
- `xlsxzipfuncimpl.inc`
- `zezippyfpc.inc`

From `zexmlss/`:
- Entire `delphizip/` directory (contains kazip/, abbreviazip/, jcl7z/, SciZip/, synzip/, dummy/ subdirectories)

**Step 1: Delete files**

```bash
cd zexmlss/src
rm -f zeZippy.pas zeZippyXE2.pas zeZippyAB.pas zeZippyJCL7z.pas zeZippyZipMaster.pas zeZippyLazip.pp
rm -f zeSave.pas zeSaveEXML.pas zeSaveODS.pas zeSaveXLSX.pas
rm -f odszipuses.inc odszipfunc.inc odszipfuncimpl.inc
rm -f xlsxzipuses.inc xlsxzipfunc.inc xlsxzipfuncimpl.inc
rm -f zezippyfpc.inc
cd ..
rm -rf delphizip/
```

**Step 2: Verify no remaining references**

Search for any remaining references to deleted units:
```bash
grep -rn "zeZippy\|zeSave\|zeSaveEXML\|zeSaveODS\|zeSaveXLSX\|odszipuses\|xlsxzipuses\|odszipfunc\|xlsxzipfunc\|zezippyfpc\|zipusekazip" zexmlss/src/
```

Expected: no matches (zeodfs.pas was already cleaned in Task 3).

Check example code too:
```bash
grep -rn "zeZippy\|zeSave\b" zexmlss/examples/
```

The `createexml` example (`unit_main.pas` line 42) references `zeSave, zeSaveODS, zeSaveXLSX, zeSaveEXML, zeZippyXE2`. This will need updating later when we address examples, but for now note it as a known breakage of the old VCL example.

**Step 3: Commit**

```bash
git add -A
git commit -m "Delete zeZippy, zeSave*, alternative zip backends and delphizip/"
```

---

### Task 5: Update existing example references

**Files:**
- Modify: `zexmlss/examples/createexml/unit_main.pas`

**Step 1: Update uses clause (lines 40-42)**

Replace:
```pascal
 {$IFNDEF FPC}
 {$IF CompilerVersion > 22}zeZippyXE2, {$ELSE} zeZippyAB,{. . .} {$IFEND}
 {$Else} zeZippyLaz,{$EndIf}
  zeSave, zeSaveODS, zeSaveXLSX, zeSaveEXML, zexlsx, zeodfs;
```

With:
```pascal
  zexlsx, zeodfs;
```

The example code uses `ExportXmlssToODFS` and `ExportXmlssToXLSX` directly — check the body for any `IZXMLSSave` usage and replace with direct function calls if needed.

**Step 2: Check for zeSave/IZXMLSSave usage in the example body**

Read the full `unit_main.pas` to find any fluent API usage (`.As_()`, `.To_()`, `.Save()` etc.). Replace with direct `ExportXmlssToODFS()` / `SaveXmlssToEXML()` / `ExportXmlssToXLSX()` calls.

**Step 3: Commit**

```bash
git add zexmlss/examples/createexml/unit_main.pas
git commit -m "Update createexml example to use direct export functions"
```

---

### Task 6: Update Delphi 12 package (.dpk)

**Files:**
- Modify: `zexmlss/packages/delphi12/zexmlsslib.dpk`

**Step 1: Update contains clause**

Replace the current contains (lines 32-41) with:
```pascal
contains
  zexmlss in '..\..\src\zexmlss.pas',
  zsspxml in '..\..\src\zsspxml.pas',
  zeodfs in '..\..\src\zeodfs.pas',
  zeformula in '..\..\src\zeformula.pas',
  zesavecommon in '..\..\src\zesavecommon.pas',
  zexmlssutils in '..\..\src\zexmlssutils.pas',
  zearchhelper in '..\..\src\zearchhelper.pas',
  zexlsx in '..\..\src\zexlsx.pas',
  zenumberformats in '..\..\src\zenumberformats.pas',
  zeZipper in '..\..\src\zeZipper.pas';
```

Note: `zeZipper` is added (was missing). No `zeZippy` or `zeSave*` units.

**Step 2: Commit**

```bash
git add zexmlss/packages/delphi12/zexmlsslib.dpk
git commit -m "Update Delphi 12 package to include zeZipper, remove deleted units"
```

---

### Task 7: Create Delphi 12 package (.dproj)

**Files:**
- Create: `zexmlss/packages/delphi12/zexmlsslib.dproj` (overwrite the old one)

**Step 1: Write the .dproj**

Create a modern Delphi 12.2 Athens MSBuild project file. Key properties:
- ProjectVersion: 20.2
- Platforms: Win32, Win64
- FrameworkType: VCL
- Configurations: Debug, Release
- DCC_UnitSearchPath: `..\..\src`
- All DCCReference entries matching the .dpk contains clause

The file should follow the standard Delphi 12 package .dproj format with proper Embarcadero targets imports.

**Step 2: Commit**

```bash
git add zexmlss/packages/delphi12/zexmlsslib.dproj
git commit -m "Recreate Delphi 12 package project file for Athens"
```

---

### Task 8: Update Lazarus package (.lpk)

**Files:**
- Modify: `zexmlss/packages/lazarus/zexmlsslib.lpk`

**Step 1: Remove deleted units from Files section**

Remove these items from the `<Files>` section:
- `zeSave.pas` (Item7)
- `zeSaveEXML.pas` (Item8)
- `zeSaveODS.pas` (Item9)
- `zeSaveXLSX.pas` (Item10)
- `zeZippy.pas` (Item11)
- `zeZippyLazip.pp` (Item12)

Renumber remaining items. Update `Count` attribute to reflect the new count (9 items).

**Step 2: Commit**

```bash
git add zexmlss/packages/lazarus/zexmlsslib.lpk
git commit -m "Update Lazarus package, remove deleted units"
```

---

### Task 9: Create console example project

**Files:**
- Create: `zexmlss/examples/odsconsole/odsconsole.dpr`
- Create: `zexmlss/examples/odsconsole/odsconsole.dproj`
- Create: `zexmlss/examples/odsconsole/u_odsexport.pas`

**Step 1: Create the console program entry point**

`odsconsole.dpr`:
```pascal
program odsconsole;

{$APPTYPE CONSOLE}

{$I ..\..\src\zexml.inc}

uses
  SysUtils,
  u_odsexport in 'u_odsexport.pas';

begin
  try
    WriteLn('zexmlss ODS export example');
    WriteLn('Generating response.ods...');
    GenerateODS('response.ods');
    WriteLn('Done. File saved to: ', ExpandFileName('response.ods'));
  except
    on E: Exception do
    begin
      WriteLn('Error: ', E.Message);
      ExitCode := 1;
    end;
  end;
end.
```

**Step 2: Create the export unit**

`u_odsexport.pas` — demonstrates creating a workbook matching the Python script's output structure (header row with gray background, data rows, summary with yellow background, formulas).

The unit should:
1. Create `TZEXMLSS` instance
2. Add one sheet
3. Define styles: gray header, white normal, bold, yellow summary, green income, blue expense
4. Set column widths for A-O
5. Write header row (A1..O1)
6. Write ~5 hardcoded data rows with cell values and formulas
7. Write summary row with SUM formulas
8. Call `SaveXmlssToODFS(XMLSS, FileName)`

Key formula example (OpenFormula notation for ODS):
```pascal
// Formula: =SUM(I2:I6)  in OpenFormula notation
Cell[8, SummaryRow].Formula := 'of:=SUM([.I2:.I6])';
Cell[8, SummaryRow].CellType := ZENumber;
```

**Step 3: Create the .dproj**

Standard Delphi 12 console application .dproj. DCC_UnitSearchPath must include `..\..\src`.

**Step 4: Commit**

```bash
git add zexmlss/examples/odsconsole/
git commit -m "Add console example for ODS generation"
```

---

### Task 10: Update CLAUDE.md

**Files:**
- Modify: `CLAUDE.md`

**Step 1: Update architecture section**

Remove references to:
- zeZippy abstraction layer
- zeSave fluent interface (IZXMLSSave)
- Multiple zip backends

Update the zip section to reflect the simplified architecture:
- Delphi: `zeZipper.pas` wraps `System.Zip`
- FPC: uses RTL `zipper` unit directly

Add the console example to the examples section.

**Step 2: Commit**

```bash
git add CLAUDE.md
git commit -m "Update CLAUDE.md to reflect simplified zip architecture"
```

---

### Task 11: Compile and verify

**Step 1: Compile the Delphi 12 package**

Use the delphi-build MCP server:
```
compile_delphi_project("zexmlss/packages/delphi12/zexmlsslib.dproj")
```

Fix any compilation errors.

**Step 2: Compile the console example**

```
compile_delphi_project("zexmlss/examples/odsconsole/odsconsole.dproj")
```

Fix any compilation errors.

**Step 3: Commit any fixes**

```bash
git add -A
git commit -m "Fix compilation issues found during verification"
```
