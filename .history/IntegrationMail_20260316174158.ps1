# Installation du module si nécessaire
# Install-Module ImportExcel -Scope CurrentUser

Import-Module ImportExcel

# Chemins
$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$sheetName = "Magasins AvivA"

$baseFolder = "C:\Users\steven.azar\Documents\yuccan_assets"
$bannerPath = "$baseFolder\banner.png"
$qrFolder = "$baseFolder\qr_codes"

$outputJson = "C:\Users\steven.azar\Documents\yuccan_assets\yuccan_assets.json"

# Fonction conversion Base64
function ConvertTo-Base64 {
    param($filePath)

    if (Test-Path $filePath) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        return [Convert]::ToBase64String($bytes)
    }
    else {
        return $null
    }
}

# Encodage bannière
$bannerBase64 = ConvertTo-Base64 $bannerPath

# Lecture du fichier Excel
$magasins = Import-Excel -Path $excelPath -WorksheetName $sheetName

$result = @()

foreach ($m in $magasins) {

    $id = $m.ID
    $nom = $m."Nom du Magasin"
    $plaquette = $m."Plaquette Digitale"

    # QR code basé sur l'ID
    $qrPath = Join-Path $qrFolder "$id.png"

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

# Export JSON
$result | ConvertTo-Json -Depth 5 | Out-File $outputJson -Encoding UTF8

Write-Host "JSON généré ici :" $outputJson