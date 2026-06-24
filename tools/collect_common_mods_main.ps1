$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$current_dir = (Get-Location).ProviderPath
Set-Location $PSScriptRoot

$bat_dir = $args[0]
$bat_file = $args[1]
$bat_path = $args[2]
$arg_dir = $args[3]
$arg_file = $args[4]
$arg_path = $args[5]

Set-Variable -Name IN_DIR_NAME -Value "modules" -Option Constant
Set-Variable -Name OUT_DIR_NAME -Value "common_modules_repo" -Option Constant
$mod_names = @(
    "ArrayObject.cls",
    "Fx_Common.bas",
    "Lib_Common.bas",
    "Lib_CommonConstructor.bas",
    "Lib_FileSystem.bas",
    "Lib_InputSheet.bas",
    "Lib_IPv4.bas",
    "Lib_TextFile.bas",
    "Lib_UnitTest.bas",
    "ApplicationScreenUpdateManager.cls",
    "CommonRunStateManager.cls",
    "Counter.cls",
    "CounterSet.cls",
    "DebugInformation.cls",
    "Enumerator.cls",
    "FileSystemService.cls",
    "FileSystemServiceTestDouble.cls",
    "IComparable.cls",
    "IDuplicateCheckable.cls",
    "IElementTypeProvider.cls",
    "IEnumerator.cls",
    "IEquatable.cls",
    "IFileSystemService.cls",
    "IStringable.cls",
    "ITextFileEntity.cls",
    "ITextFileService.cls",
    "IUserInputSheet.cls",
    "IWorkbookService.cls",
    "IWorksheetService.cls",
    "ObjectList.cls",
    "ObjectDictionary.cls",
    "ObjectSet.cls",
    "ProgressStatus.cls",
    "TestDoubleBehaviorStore.cls",
    "TestDoubleCallRecord.cls",
    "TestDoubleVariantKeyBuilder.cls",
    "TextFileEntity.cls",
    "TextFileEntityTestDouble.cls",
    "TextFileService.cls",
    "TextFileServiceTestDouble.cls",
    "UnitTestAssert.cls",
    "UserInputSheet.cls",
    "UserInputSheetTestDouble.cls",
    "WorkbookService.cls",
    "WorkbookServiceTestDouble.cls",
    "WorksheetRangeBounds.cls",
    "WorksheetRangeBoundsEnumerator.cls",
    "WorksheetVirtualTable.cls",
    "WorksheetVirtualTableEnumerator.cls",
    "WorksheetService.cls",
    "WorksheetServiceTestDouble.cls")

Import-Module ./ExcelModuleManager

if ([System.String]::IsNullOrEmpty($arg_path)) {
    $target_root = $current_dir
}
else {
    $target_root = $arg_path
}

foreach ($mod_name in $mod_names) {
    $dirs = (Get-ChildItem $current_dir -Recurse -Directory -Filter $IN_DIR_NAME)
    $files = @()
    foreach ($dir_item in $dirs) {
        $files += (Get-ChildItem $dir_item.FullName -File -Filter $mod_name | Where-Object { $_.DirectoryName -ne $OUT_DIR_NAME })
    }
    [System.IO.FileInfo]$latest_file = $null
    foreach ($file in $files) {
        if ($null -eq $latest_file -or $latest_file.LastWriteTime -lt $file.LastWriteTime) {
            $latest_file = $file
        }
    }
    
    if ($null -eq $latest_file) {
        Write-Warning "Module $mod_name was not found."
        continue
    } else {
        $output_path = (Join-Path (Join-Path $target_root $OUT_DIR_NAME) $latest_file.Name)

        Write-Host $latest_file.FullName
        Copy-Item $latest_file.FullName $output_path -Force
    }
}
