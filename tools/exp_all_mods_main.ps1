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

Set-Variable -Name OUT_DIR_NAME -Value "modules" -Option Constant

Import-Module ./ExcelModuleManager

if ([System.String]::IsNullOrEmpty($arg_path)) {
    $target_root = $current_dir
}
else {
    $target_root = $arg_path
}

$xlsm_files = Get-ChildItem $target_root -Recurse -File -Filter "*.xlsm"

foreach ($xlsm_file in $xlsm_files) {
    $input_file_path = $xlsm_file.FullName
    $xlsm_dir_path = Split-Path $input_file_path
    $output_dir_path = (Join-Path $xlsm_dir_path $OUT_DIR_NAME)

    Export-AllModuleFromExcelFile $input_file_path $output_dir_path -Force
}
