Import-Module ImportExcel

# --- CONFIGURATION DES CHEMINS ---
$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$sheetName = "Magasins AvivA"
$assetsRoot = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMail\yuccan_assets"

$bannerPath = Join-Path $assetsRoot "banner.png"
$qrFolder = Join-Path $assetsRoot "qr_codes"
$outputJson = Join-Path $assetsRoot "yuccan_assets.json"

# --- VERIFICATION DES DOSSIERS ---
if (!(Test-Path $qrFolder)) { 
    Write-Error "Le dossier des QR Codes est introuvable : $qrFolder"
    # On s'arrête si le dossier source n'existe pas pour éviter les 'null' en série
    return 
}

# --- FONCTION BASE64 AMÉLIORÉE ---
# Ajoute le préfixe 'data:image/png;base64,' indispensable pour l'intégration HTML/PHP
function ConvertTo-Base64-DataUri($filePath) {
    if (Test-Path $filePath) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $base64 = [Convert]::ToBase64String($bytes)
            $extension = [System.IO.Path]::GetExtension($filePath).Replace(".", "")
            return "data:image/$extension;base64,$base64"
        }
        catch {
            Write-Warning "Erreur lors de la lecture de : $filePath"
            return $null
        }
    }
    else {
        Write-Host "Fichier introuvable : $filePath" -ForegroundColor Yellow
        return $null
    }
}

# --- ENCODAGE DE LA BANNIÈRE (Une seule fois) ---
$bannerDataUri = ConvertTo-Base64-DataUri $bannerPath

# --- LECTURE EXCEL ET TRAITEMENT ---
$magasins = Import-Excel -Path $excelPath -WorksheetName $sheetName
$result = @()

foreach ($m in $magasins) {
    # Nettoyage de l'ID
    $id = [int]$m.ID
    $nom = $m."Nom du Magasin"
    $plaquette = $m."Plaquette Digitale"

    # Construction du chemin du QR Code (Vérifie bien que c'est du .png !)
    $qrPath = Join-Path $qrFolder "$id.png"
    
    Write-Host "Traitement Magasin : $nom (ID: $id)"
    
    $qrDataUri = ConvertTo-Base64-DataUri $qrPath

    $obj = [PSCustomObject]@{
        id                 = $id
        magasin            = $nom
        plaquette_digitale = $plaquette
        banner_base64      = $bannerDataUri
        qrcode_base64      = $qrDataUri
    }

    $result += $obj
}

# --- EXPORT JSON ---
# Utilisation de -Compress si le fichier est trop lourd, sinon garde tel quel pour la lisibilité
$result | ConvertTo-Json -Depth 5 | Out-File $outputJson -Encoding UTF8

Write-Host "`nTerminé ! JSON généré : $outputJson" -ForegroundColor Green