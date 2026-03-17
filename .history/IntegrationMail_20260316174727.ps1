Import-Module ImportExcel

# Chemins
$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$sheetName = "Magasins AvivA"

$baseFolder = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMail\yuccan_assets"
$bannerPath = Join-Path $baseFolder "banner.png"
$qrFolder = Join-Path $baseFolder "qr_codes"

$outputJson = Join-Path $baseFolder "yuccan_assets.json"

# Création dossiers si besoin
if (!(Test-Path $baseFolder)) {
    New-Item -ItemType Directory -Path $baseFolder | Out-Null
}

if (!(Test-Path $qrFolder)) {
    New-Item -ItemType Directory -Path $qrFolder | Out-Null
}

# Fonction conversion Base64
function ConvertTo-Base64($filePath) {

    if (!(Test-Path $filePath)) {
        Write-Host "Fichier introuvable :" $filePath -ForegroundColor Red
        return $null
    }

    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    return [Convert]::ToBase64String($bytes)
}

Write-Host "Lecture bannière :" $bannerPath

$bannerBase64 = ConvertTo-Base64 $bannerPath

# Lecture Excel
$magasins = Import-Excel -Path $excelPath -WorksheetName $sheetName

$result = @()

foreach ($m in $magasins) {

    # Nettoyage ID (Excel met souvent .0)
    $id = [int]$m.ID

    $nom = $m."Nom du Magasin"
    $plaquette = $m."Plaquette Digitale"

    $qrPath = Join-Path $qrFolder "$id.png"

    Write-Host "Recherche QR code :" $qrPath

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

$result | ConvertTo-Json -Depth 5 | Set-Content $outputJson -Encoding UTF8

Write-Host ""
Write-Host "JSON généré :" $outputJson -ForegroundColor Green