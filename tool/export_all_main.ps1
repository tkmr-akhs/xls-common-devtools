$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Set-Variable -Name OUT_DIR_NAME -Value "modules" -Option Constant

Set-Location $PSScriptRoot
Import-Module ./ExcelModuleManager

$xlsm_files = Get-ChildItem $args[0] -Recurse -File -Filter "*.xlsm"

foreach ($xlsm_file in $xlsm_files) {
    $input_file_path = $xlsm_file.FullName
    $xlsm_dir_path = Split-Path $input_file_path
    $output_dir_path = (Join-Path $xlsm_dir_path $OUT_DIR_NAME)

    Export-AllModuleFromExcelFile $input_file_path $output_dir_path -Force
}