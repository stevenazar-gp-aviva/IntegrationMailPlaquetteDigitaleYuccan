Import-Module ImportExcel

# --- CONFIGURATION ---
$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$assetsRoot = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMail\yuccan_assets"

# On définit des chemins par défaut au cas où
$bannerPath = Join-Path $assetsRoot "banner.png"
$qrFolder = Join-Path $assetsRoot "qr_codes"
$outputJson = Join-Path $assetsRoot "yuccan_assets.json"

# --- FONCTION DE CONVERSION HAUTE DISPONIBILITÉ ---
function Get-ImageDataUri($filePath) {
    if (Test-Path $filePath) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $base64 = [Convert]::ToBase64String($bytes)
        $ext = [System.IO.Path]::GetExtension($filePath).Replace(".", "")
        return "data:image/$ext;base64,$base64"
    }
    return $null # Le dév PHP devra gérer le cas si c'est vide
}

# --- LOGIQUE DE TRAITEMENT ---
Write-Host "--- DÉBUT DU TRAITEMENT AVIVA ---" -ForegroundColor Cyan

if (Test-Path $excelPath) {
    $magasins = Import-Excel -Path $excelPath -WorksheetName "Magasins AvivA"
    $bannerBase64 = Get-ImageDataUri $bannerPath
    
    $result = foreach ($m in $magasins) {
        $id = [int]$m.ID
        
        # Tentative de trouver le QR code avec plusieurs extensions possibles
        $qrFile = Get-ChildItem -Path $qrFolder -Filter "$id.*" | Select-Object -First 1
        $qrBase64 = if ($qrFile) { Get-ImageDataUri $qrFile.FullName } else { $null }

        if (-not $qrFile) {
            Write-Host " [!] QR Code manquant pour : $($m."Nom du Magasin") (ID: $id)" -ForegroundColor Yellow
        }

        [PSCustomObject]@{
            id                 = $id
            magasin            = $m."Nom du Magasin"
            plaquette_digitale = $m."Plaquette Digitale"
            # On passe les données prêtes à l'emploi pour le HTML
            assets             = @{
                banner_html = $bannerBase64
                qrcode_html = $qrBase64
            }
        }
    }

    # Export avec formatage pour que le dev PHP puisse le lire à l'oeil nu
    $result | ConvertTo-Json -Depth 5 | Out-File $outputJson -Encoding UTF8
    Write-Host "`n[OK] JSON généré : $outputJson" -ForegroundColor Green
}
else {
    Write-Error "Fichier Excel introuvable à l'adresse : $excelPath"
}