param(
    [string]$JsonPath = "$PSScriptRoot\..\..\config\Launcher\Locales.json",
    [string]$OutputDir = "$PSScriptRoot\..\..\src\OfficeToolPlus\Dictionaries\Languages"
)

if (-not (Test-Path $JsonPath)) {
    Write-Error "Locales.json not found: $JsonPath"
    exit 1
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$Data = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

# Collect all unique locale codes
$localeSet = @{}
foreach ($product in $Data) {
    foreach ($locale in $product.Locales.PSObject.Properties.Name) {
        $localeSet[$locale] = $true
    }
}
$locales = $localeSet.Keys

foreach ($locale in $locales) {
    $lines = @()
    foreach ($product in $Data) {
        if ($product.Locales.$locale) {
            $desc = $product.Locales.$locale.Description
            $lines += "$($product.Name): $desc"
        }
    }
    $filePath = Join-Path $OutputDir "$locale.txt"
    $lines | Set-Content -Path $filePath -Encoding UTF8
    Write-Host "Generated: $filePath"
}

Write-Host "Language files generated in $OutputDir" -ForegroundColor Green