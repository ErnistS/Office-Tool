param(
    [string]$JsonPath = "$PSScriptRoot\..\..\config\Launcher\Locales.json"
)

if (-not (Test-Path $JsonPath)) {
    Write-Error "Locales.json not found at: $JsonPath"
    exit 1
}

try {
    $Data = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json
    if ($null -eq $Data) {
        Write-Host "Locales.json is empty." -ForegroundColor Yellow
        $Data = @()
    }
} catch {
    Write-Error "Failed to parse JSON: $($_.Exception.Message)"
    exit 1
}

Write-Host "Locales loaded: $($Data.Count) item(s)" -ForegroundColor Green
Write-Host

foreach ($product in $Data) {
    Write-Host "Product: $($product.Name)" -ForegroundColor Cyan
    Write-Host "  Description: $($product.Description)"
    Write-Host "  URL: $($product.URL)"
    Write-Host "  Locales: $($product.Locales.PSObject.Properties.Name -join ', ')"
    Write-Host
}

Read-Host "Press Enter to exit"
