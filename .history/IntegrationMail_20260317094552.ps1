Import-Module ImportExcel

$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$assetsRoot = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMail\yuccan_assets"
$outputJson = Join-Path $assetsRoot "yuccan_assets.json"

# --- FONCTION : IMAGE -> BASE64 ---
function Get-ImageDataUri($filePath) {
    if (Test-Path $filePath) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $base64 = [Convert]::ToBase64String($bytes)
        $ext = [System.IO.Path]::GetExtension($filePath).Replace(".", "")
        return "data:image/$ext;base64,$base64"
    }
    return $null
}

# --- FONCTION : GÉNÉRER QR CODE DEPUIS URL ---
function Get-QrCodeBase64Remote($url) {
    # Utilisation d'une API gratuite pour générer le QR Code à la volée
    $apiUrl = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=$([Uri]::EscapeDataString($url))"
    try {
        $bytes = Invoke-WebRequest -Uri $apiUrl -UseBasicParsing -TimeoutSec 5 | Select-Object -ExpandProperty Content
        $base64 = [Convert]::ToBase64String($bytes)
        return "data:image/png;base64,$base64"
    }
    catch {
        return $null
    }
}

Write-Host "--- GÉNÉRATION FLUX AVIVA ---" -ForegroundColor Cyan

if (Test-Path $excelPath) {
    $magasins = Import-Excel -Path $excelPath -WorksheetName "Magasins AvivA"
    
    # On essaie de charger la bannière locale
    $bannerPath = Join-Path $assetsRoot "banner.png"
    $bannerBase64 = Get-ImageDataUri $bannerPath

    $result = foreach ($m in $magasins) {
        $id = [int]$m.ID
        $urlPlaquette = $m."Plaquette Digitale"
        
        Write-Host "Traitement : $($m."Nom du Magasin")" -NoNewline

        # 1. Tentative QR Code Local
        $qrFolder = Join-Path $assetsRoot "qr_codes"
        $qrFile = Get-ChildItem -Path $qrFolder -Filter "$id.*" | Select-Object -First 1
        
        if ($qrFile) {
            $qrBase64 = Get-ImageDataUri $qrFile.FullName
            Write-Host " [OK Local]" -ForegroundColor Green
        }
        else {
            # 2. Si pas de fichier, on le génère via l'URL de la plaquette !
            $qrBase64 = Get-QrCodeBase64Remote $urlPlaquette
            Write-Host " [GÉNÉRÉ VIA API]" -ForegroundColor Yellow
        }

        [PSCustomObject]@{
            id                 = $id
            magasin            = $m."Nom du Magasin"
            plaquette_digitale = $urlPlaquette
            assets             = @{
                banner_html = $bannerBase64
                qrcode_html = $qrBase64
            }
        }
    }

    $result | ConvertTo-Json -Depth 5 | Out-File $outputJson -Encoding UTF8
    Write-Host "`nSuccès ! Le JSON est complet." -ForegroundColor Green
}