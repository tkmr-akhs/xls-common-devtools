$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
#$DebugPreference = "Continue"

$current_dir = (Get-Location).ProviderPath
Set-Location $PSScriptRoot

$bat_dir = $args[0]
$bat_file = $args[1]
$bat_path = $args[2]
$arg_dir = $args[3]
$arg_file = $args[4]
$arg_path = $args[5]

Set-Variable -Name OUT_DIR_NAME -Value "modules" -Option Constant

if ([System.String]::IsNullOrEmpty($arg_path)) {
    $target_root = $current_dir
}
else {
    $target_root = $arg_path
}

$mod_dir = (Join-Path $target_root "common_modules_repo")

Import-Module ./ExcelModuleManager

$xlsm_files = Get-ChildItem $target_root -Recurse -File -Filter "*.xlsm"
$mod_files = (Get-ChildItem $mod_dir -File | ForEach-Object { $_.FullName })

foreach ($xlsm_file in $xlsm_files) {
    $retry_count = 1
    $import_result = $null
    while ($retry_count -le 3) {
        try {
            $import_result = Import-ModuleToExcelFile $xlsm_file.FullName $mod_files -Update
            #$import_result = Import-ModuleToExcelFile $xlsm_file.FullName $mod_files
            break
        }
        catch {
            Write-Warning "Failed on attempt #($retry_count) for '$($xlsm_file.FullName)'."
            Write-Warning $_
            $retry_count ++
        }
    }
    if ($import_result) {
        # Import succeeded
        $output_dir = (Join-Path (Split-Path $xlsm_file.FullName) $OUT_DIR_NAME)
        if (Test-Path $output_dir) {
            foreach ($mod_file in $mod_files) {
                if ($import_result[$mod_file]) {
                    Copy-Item $mod_file $output_dir -Force
                    Write-Information "Copied file '$mod_file'."
                }
                else {
                    Write-Information "Skipped copying file '$mod_file'."
                }
            }
        }
    }
    else {
        # Import failed
        Write-Information "Skipped Excel file '$mod_file'."
    }
}
