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

$xlsm_file = Get-Item $arg_path
$mod_dir = (Join-Path (Split-Path $arg_path) $IN_DIR_NAME)

Import-Module ./ExcelModuleManager

$mod_files = (Get-ChildItem $mod_dir -File | ForEach-Object { $_.FullName })

Import-ModuleToExcelFile $xlsm_file.FullName $mod_files -Flush > $null
