# xls-common-devtools

`xls-common-devtools` is a development repository for `CommonModules`, which are shared by multiple Excel VBA tools, and for supporting reflection, extraction, distribution, formatting, and documentation generation.

The source of truth for common modules is `CommonModules/modules`. Common modules under an individual tool's `common_modules_repo` or `modules` are treated as distributed or reflected copies.

## Features

This repository lets Excel VBA macros be treated as development assets that can be source-controlled, reused, and tested, rather than as logic confined inside a workbook.

One way to share common processing is to place it in an `.xlsb` or add-in. However, when deliverables depend on external workbooks or add-in references, they are no longer self-contained single files and become less portable. Another approach is to manually copy common processing into each `.xlsm`; that keeps each `.xlsm` easy to distribute, but makes updates across multiple workbooks easy to miss.

`xls-common-devtools` centrally manages common modules as VBA source in `CommonModules`, then reflects the necessary common modules into each `.xlsm` with scripts. Distribution keeps the portability of a single `.xlsm`, while development can keep common modules current across multiple individual tools.

Excel macros have very few widely established unit-test frameworks. `xls-common-devtools` provides a lightweight test runner, assertions, and a test double foundation that run inside Excel workbooks, allowing regression checks with `UnitTestMain` after VBA source has been reflected into the actual `.xlsm`.

- Externalizes VBA source as `.bas` / `.cls` files so Git can manage diffs.
- Centrally manages common modules and distributes the same implementation to multiple individual tools.
- Allows each individual tool to be distributed as a single `.xlsm` with the necessary VBA source reflected into it.
- Wraps Excel workbook, worksheet, file, and text I/O operations as services so individual modules are easier to test.
- Provides typed sets, keyed references, and virtual tables over input ranges through `ObjectList`, `ObjectSet`, `ObjectDictionary`, and `WorksheetVirtualTable`.
- Provides regression testing for VBA through `Lib_UnitTest`, `UnitTestAssert`, and each `*TestDouble`.
- Generates API reference documentation for VBA source through Doxygen and `DoxyVB6`.

## Core Capabilities

`CommonModules` centralizes the following kinds of processing.

| Category | Main VBA sources |
| --- | --- |
| Excel workbook/worksheet operations | `WorkbookService`, `WorksheetService`, `WorksheetRangeBounds`, `WorksheetVirtualTable`, `Lib_Common`, `Lib_CommonConstructor` |
| Collections/enumeration | `ObjectList`, `ObjectSet`, `ObjectDictionary`, `Counter`, `CounterSet`, `Enumerator`, `ArrayObject`, `IElementTypeProvider` |
| Input sheets | `Lib_InputSheet`, `IUserInputSheet`, `UserInputSheet`, `UserInputSheetTestDouble` |
| IPv4 | `Lib_IPv4` |
| Files/text | `FileSystemService`, `TextFileService`, `TextFileEntity` |
| Test support | `Lib_UnitTest`, `UnitTestAssert`, `TestDoubleBehaviorStore`, each `*TestDouble` |
| Run state/diagnostics | `ApplicationScreenUpdateManager`, `CommonRunStateManager`, `DebugInformation`, `ProgressStatus` |

Normal GUI entry points call `InitializeCommonService(Force:=True)` and use `WbSrv`, `WsSrv`, `New_RangeBounds`, and related APIs. Public UDF entry points used as Excel worksheet functions use `InitializeUdfCommonService` to avoid side effects during cell recalculation.

For range processing, `WorksheetRangeBounds.GetRows()` / `GetColumns()` and `WorksheetVirtualTable` allow worksheet rectangles to be handled as rows, columns, or headed records. `ObjectDictionary` is a typed dictionary that supports key references, and its standard `For Each` enumerates keys. `FileSystemService.GetAbsolutePath` expands OS environment variables such as `%LOCALAPPDATA%` in non-URL local paths before converting them to absolute paths.

```vb
Call InitializeCommonService(Force:=True)

Dim target_cell As WorksheetRangeBounds
Set target_cell = New_RangeBounds(Row:=1, Column:=1, Sheet:="INPUT")

Call WsSrv.WriteCell(target_cell, "value", TypeConvert:=False)
```

## Requirements

- Windows and Microsoft Excel are required.
- Reflecting or extracting Excel VBA projects and running `UnitTestMain` require Excel's "Trust access to the VBA project object model" setting.
- Do not leave the target `.xlsm` open manually while reflecting or extracting.
- API reference generation requires Doxygen. `tools/DoxyVB6/DoxyVB6.exe` is used as the VBA input filter.

## Development Workflow

Individual tool developers edit the individual tool's `modules`, reflect them into the target `.xlsm`, and test the workbook. This section uses `xls-web-tools/SampleWebTool` as an example.

1. Edit the individual modules under `xls-web-tools/SampleWebTool/modules`.
2. If you need to change a copy of a common module, first decide whether the work should be handled as a `CommonModules` change instead of as an individual module change.
3. Update `xls-web-tools/SampleWebTool/modules/Test_*.bas` or test `.cls` files as needed.
4. Run the formatting check for `xls-web-tools/SampleWebTool/modules` with `tools/format_vba_source_main.ps1`.
5. Reflect VBA source into `SampleWebTool.xlsm` with `xls-web-tools/SampleWebTool/IMP_MODS.lnk` or `tools/IMP_MODS.BAT`.
6. Run `UnitTestMain` in `SampleWebTool.xlsm` and verify that every result on `UNIT_TEST_SHEET` is `OK`.
7. If worksheet buttons or GUI entry points changed, also verify the real operation path.
8. If you had to edit in VBE, extract VBA source from `.xlsm` with `xls-web-tools/SampleWebTool/EXP_MODS.lnk` or `tools/EXP_MODS.BAT`.
9. When taking in common-module distribution updates, reflect them into `.xlsm` files under `xls-web-tools` with `xls-web-tools/IMP_COMMON_MODS.lnk` or `tools/IMP_COMMON_MODS.BAT`, then rerun tests for the target individual tools.

## Unit Testing

Tests are written as `Test_...(ByVal Assert As UnitTestAssert)` subprocedures in `Test_*.bas`, then executed in the target workbook with `UnitTestMain`. Results are output to `UNIT_TEST_SHEET`.

For example, `xls-web-tools/SampleWebTool/modules` contains test modules such as:

- `Test_WebDriverSessionClient.bas`
- `Test_WebDriverClient.bas`
- `Test_ToolSettings.bas`
- `Test_OutputSheetWriter.bas`

The basic form is Arrange to prepare the subject and test doubles, Act to call the target processing, and Assert to verify return values, state, and call history.

```vb
Public Sub Test_Sample(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim actual_value As String

    ' --- Act ---
    actual_value = "expected"

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "expected", actual_value
End Sub
```

Actual tests substitute dependencies such as `WorksheetServiceTestDouble`, `WorkbookServiceTestDouble`, `FileSystemServiceTestDouble`, and `UserInputSheetTestDouble` for `WbSrv` / `WsSrv` / `FsSrv` or input sheet dependencies. This allows reads, writes, and issued calls to be verified without mutating real Excel sheets or files.

## Tooling

Use the `.bat` files or the `.lnk` files placed in individual tool folders when those entry points exist. The usual operation is to drag and drop the target folder or `.xlsm` onto the `.bat` / `.lnk`. When running from the command line, call the `.bat` files.

| Purpose | Recommended entry point | Target example |
| --- | --- | --- |
| Reflect individual modules into `SampleWebTool.xlsm` | `xls-web-tools/SampleWebTool/IMP_MODS.lnk` or `tools/IMP_MODS.BAT` | `xls-web-tools/SampleWebTool/SampleWebTool.xlsm` |
| Extract VBA source from `SampleWebTool.xlsm` | `xls-web-tools/SampleWebTool/EXP_MODS.lnk` or `tools/EXP_MODS.BAT` | `xls-web-tools/SampleWebTool/SampleWebTool.xlsm` |
| Generate the API reference for `SampleWebTool` | `xls-web-tools/SampleWebTool/GEN_DOC.lnk` or `tools/GEN_DOC.BAT` | `xls-web-tools/SampleWebTool/modules` |
| Reflect common modules into `.xlsm` files under `xls-web-tools` | `xls-web-tools/IMP_COMMON_MODS.lnk` or `tools/IMP_COMMON_MODS.BAT` | `xls-web-tools` |
| Collect common modules into `common_modules_repo` | `tools/COLLECT_COMMON_MODS.BAT` | `xls-common-devtools` |
| Distribute to each individual tool repository's `common_modules_repo` | `tools/DIST_COMMON_MODS_REPO.BAT` | `xls-common-devtools/common_modules_repo` |

### Command-Line Examples

The following examples are run from the `xls-common-devtools` root.

```powershell
# Reflect individual modules into SampleWebTool.xlsm
.\tools\IMP_MODS.BAT ..\xls-web-tools\SampleWebTool\SampleWebTool.xlsm

# Extract VBA source from SampleWebTool.xlsm
.\tools\EXP_MODS.BAT ..\xls-web-tools\SampleWebTool\SampleWebTool.xlsm

# Generate the API reference for SampleWebTool
.\tools\GEN_DOC.BAT ..\xls-web-tools\SampleWebTool\modules

# Reflect common modules into .xlsm files under xls-web-tools
.\tools\IMP_COMMON_MODS.BAT ..\xls-web-tools

# Check VBA source formatting
powershell -ExecutionPolicy bypass -NoLogo -NonInteractive -File .\tools\format_vba_source_main.ps1 ..\xls-web-tools\SampleWebTool\modules -Recurse -Check

# Apply VBA source formatting
powershell -ExecutionPolicy bypass -NoLogo -NonInteractive -File .\tools\format_vba_source_main.ps1 ..\xls-web-tools\SampleWebTool\modules -Recurse
```

## Documentation

- CommonModules specification: `docs/product-spec.md`
- API reference: `docs/api-reference.zip`
- Glossary: `CONTEXT.md`
