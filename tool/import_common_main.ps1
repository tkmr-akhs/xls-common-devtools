$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
#$DebugPreference = "Continue"

Set-Variable -Name OUT_DIR_NAME -Value "modules" -Option Constant
$mod_dir = (Join-Path $args[0] "common_modules_repo")

Set-Location $PSScriptRoot
Import-Module ./ExcelModuleManager

$xlsm_files = Get-ChildItem $args[0] -Recurse -File -Filter "*.xlsm"
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