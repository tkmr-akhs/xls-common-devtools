$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Set-Variable -Name OUT_DIR_NAME -Value "modules" -Option Constant

$root_path = (Split-Path -Path $args[1])
$script_config = Get-Content (Join-Path (Join-Path $root_path $args[0]) "sync.json") | ConvertFrom-Json

$projects = $script_config.projects
$modules = $script_config.modules

$proj_dirs = @()
$mod_dirs = @()
foreach ($project in $projects) {
    $proj_dirs += (Join-Path $root_path $project)
    $mod_dirs += (Join-Path (Join-Path $root_path $project) "modules")
}

$newest_modules = @{}
foreach ($mod_dir in $mod_dirs) {
    foreach ($module in $modules) {
        $check_module = (Get-ChildItem $mod_dir "$module")
        if ($null -ne $check_module) {
            if ($null -eq $newest_modules[$module] -or $newest_modules[$module].LastWriteTime -lt $check_module.LastWriteTime ) {
                $newest_modules[$module] = $check_module
            }
        }
    }
}

Set-Location $PSScriptRoot
Import-Module ./ExcelModuleManager

$xlsm_files = @()
foreach ($proj_dir in $proj_dirs) {
    $xlsm_files += Get-ChildItem $proj_dir -File -Filter "*.xlsm"
}

$import_dict = @{}
foreach ($xlsm_file in $xlsm_files) {
    if ($null -eq $import_dict[$xlsm_file.FullName]) { $import_dict[$xlsm_file.FullName] = @() }

    foreach ($mod_filename in $newest_modules.Keys) {
        $src_path = $newest_modules[$mod_filename].FullName
        $dst_path = (Join-Path (Split-Path -Path $xlsm_file.FullName) "modules")
        if ($src_path -eq (Join-Path $dst_path $mod_filename)) {
            Write-Debug "Skip $($xlsm_file.Name) due to $(Split-Path $src_path -Leaf) is newest."
        }
        else {
            Write-Debug "Import to $($xlsm_file.Name) from $src_path"
            $import_dict[$xlsm_file.FullName] += $src_path
        }
    }
}

foreach ($xlsm_file in $import_dict.Keys) {
    if (0 -lt $import_dict[$xlsm_file].Count) {
        $dst_path = (Join-Path (Split-Path -Path $xlsm_file) "modules")
        Import-ModuleToExcelFile $xlsm_file $import_dict[$xlsm_file] | Out-Null
        foreach ($src_path in $import_dict[$xlsm_file]) {
            Copy-Item $src_path $dst_path -Force
        }
    }
    else {
        Write-Information  "Skip $xlsm_file"
    }
}
