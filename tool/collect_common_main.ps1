$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Set-Variable -Name IN_DIR_NAME -Value "modules" -Option Constant
Set-Variable -Name OUT_DIR_NAME -Value "common_modules_repo" -Option Constant
$mod_names = @(
    "ArrayObject.cls",
    "Lib_Common.bas",
    "Lib_IPv4.bas",
    "Lib_UnitTest.bas",
    "ApplicationScreenUpdateManager.cls",
    "Counter.cls",
    "CounterSet.cls",
    "DebugInformation.cls",
    "Enumerator.cls",
    "FileSystemService.cls",
    "FileSystemServiceTestDouble.cls",
    "IComparable.cls",
    "IDuplicateCheckable.cls",
    "IEnumerator.cls",
    "IEquatable.cls",
    "IFileSystemService.cls",
    "IStringable.cls",
    "ITextFileEntity.cls",
    "ITextFileService.cls",
    "IWorkbookService.cls",
    "IWorksheetService.cls",
    "ObjectList.cls",
    "ObjectSet.cls",
    "ProgressStatus.cls",
    "TextFileEntity.cls",
    "TextFileEntityTestDouble.cls",
    "TextFileService.cls",
    "TextFileServiceTestDouble.cls",
    "UnitTestAssert.cls",
    "UnitTestUtils.cls",
    "UserInputSheet.cls",
    "WorkbookService.cls",
    "WorkbookServiceTestDouble.cls",
    "WorksheetRangeBounds.cls",
    "WorksheetRangeBoundsEnumerator.cls",
    "WorksheetService.cls",
    "WorksheetServiceTestDouble.cls")

Set-Location $PSScriptRoot
Import-Module ./ExcelModuleManager

foreach ($mod_name in $mod_names) {
    $dirs = (Get-ChildItem $args[0]-Recurse -Directory -Filter $IN_DIR_NAME)
    $files = @()
    foreach ($dir_item in $dirs) {
        $files += (Get-ChildItem $dir_item.FullName -File -Filter $mod_name | Where-Object { $_.DirectoryName -ne "common_modules_repo" })
    }
    [System.IO.FileInfo]$latest_file = $null
    foreach ($file in $files) {
        if ($null -eq $latest_file -or $latest_file.LastWriteTime -lt $file.LastWriteTime) {
            $latest_file = $file
        }
    }
    
    $output_path = (Join-Path (Join-Path $args[0] $OUT_DIR_NAME) $latest_file.Name)

    Write-Host $latest_file.FullName
    Copy-Item $latest_file.FullName $output_path -Force
}