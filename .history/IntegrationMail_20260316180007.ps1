# =====================================
# MODULE EXCEL
# =====================================

Import-Module ImportExcel

# =====================================
# CHEMINS
# =====================================

$excelPath = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMailAvivA.xlsx"
$sheetName = "Magasins AvivA"

$assetsRoot = "C:\Users\steven.azar\Documents\PlaquetteDigitalIntegrationMail\yuccan_assets"
$bannerPath = Join-Path $assetsRoot "banner.png"
$qrFolder = Join-Path $assetsRoot "qr_codes"

$outputJson = Join-Path $assetsRoot "yuccan_assets.json"

# =====================================
# VERIFICATION DOSSIERS
# =====================================

if (!(Test-Path $assetsRoot)) {
    New-Item -ItemType Directory -Path $assetsRoot | Out-Null
}

if (!(Test-Path $qrFolder)) {
    New-Item -ItemType Directory -Path $qrFolder | Out-Null
}

# =====================================
# FONCTION BASE64
# =====================================

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

# =====================================
# ENCODAGE BANNIERE
# =====================================

Write-Host "Encodage bannière..."

$bannerBase64 = ConvertTo-Base64 $bannerPath

# =====================================
# LECTURE EXCEL
# =====================================

Write-Host "Lecture Excel..."

$magasinsExcel = Import-Excel -Path $excelPath -WorksheetName $sheetName

$magasins = @()

foreach ($m in $magasinsExcel) {

    # nettoyage ID Excel
    $id = [int]$m.ID

    $nom = $m."Nom du Magasin"
    $plaquette = $m."Plaquette Digitale"

    $qrPath = Join-Path $qrFolder "$id.png"

    Write-Host "Recherche QR :" $qrPath

    $qrBase64 = ConvertTo-Base64 $qrPath

    $magasin = [PSCustomObject]@{

        id                 = $id
        nom                = $nom
        plaquette_digitale = $plaquette
        qrcode_base64      = $qrBase64
    }

    $magasins += $magasin
}

# =====================================
# STRUCTURE JSON OPTIMISEE
# =====================================

$jsonStructure = [PSCustomObject]@{

    banner_base64 = $bannerBase64
    magasins      = $magasins
}

# =====================================
# EXPORT JSON
# =====================================

$jsonStructure | ConvertTo-Json -Depth 6 | Out-File $outputJson -Encoding UTF8

Write-Host ""
Write-Host "JSON généré avec succès :"
Write-Host $outputJson