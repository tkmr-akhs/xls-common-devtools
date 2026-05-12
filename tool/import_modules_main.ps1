$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Set-Variable -Name IN_DIR_NAME -Value "modules" -Option Constant
$xlsm_file = Get-Item $args[0]
$mod_dir = (Join-Path (Split-Path $args[0]) $IN_DIR_NAME)

Set-Location $PSScriptRoot
Import-Module ./ExcelModuleManager

$mod_files = (Get-ChildItem $mod_dir -File | ForEach-Object { $_.FullName })

Import-ModuleToExcelFile $xlsm_file.FullName $mod_files -Flush > $null