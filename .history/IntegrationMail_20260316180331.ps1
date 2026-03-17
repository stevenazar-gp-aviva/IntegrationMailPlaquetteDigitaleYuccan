Import-Module ImportExcel

# ================================
# CHEMINS
# ================================

$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$sheetName = "Magasins AvivA"

$assetsRoot = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMail\yuccan_assets"
$bannerPath = Join-Path $assetsRoot "banner.png"
$qrFolder = Join-Path $assetsRoot "qr_codes"

$outputJson = Join-Path $assetsRoot "yuccan_assets.json"

# ================================
# VERIFICATION DOSSIERS
# ================================

if (!(Test-Path $assetsRoot)) {
    Write-Host "Création dossier assets"
    New-Item -ItemType Directory -Path $assetsRoot | Out-Null
}

if (!(Test-Path $qrFolder)) {
    Write-Host "Création dossier qr_codes"
    New-Item -ItemType Directory -Path $qrFolder | Out-Null
}

# FONCTION BASE64

function ConvertTo-Base64($filePath) {

    if (Test-Path $filePath) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        return [Convert]::ToBase64String($bytes)
    }
    else {
        Write-Host "Fichier introuvable :" $filePath
        return $null
    }
}

# BANNIERE

$bannerBase64 = ConvertTo-Base64 $bannerPath

# LECTURE EXCEL

$magasins = Import-Excel -Path $excelPath -WorksheetName $sheetName

$result = @()

foreach ($m in $magasins) {

    # Nettoyage ID Excel (supprime .0)
    $id = [int]$m.ID

    $nom = $m."Nom du Magasin"
    $plaquette = $m."Plaquette Digitale"

    $qrPath = Join-Path $qrFolder "$id.png"

    Write-Host "Recherche QR :" $qrPath

    $qrBase64 = ConvertTo-Base64 $qrPath

    $obj = [PSCustomObject]@{

        id                 = $id
        magasin            = $nom
        plaquette_digitale = $plaquette
        banner_base64      = $bannerBase64
        qrcode_base64      = $qrBase64
    }

    $result += $obj
}

# EXPORT JSON

$result | ConvertTo-Json -Depth 5 | Out-File $outputJson -Encoding UTF8

Write-Host ""
Write-Host "JSON généré ici :"
Write-Host $outputJson