param(
    [string] $SnipeUrl = $env:SNIPE_URL,      # e.g. https://assets.example.com
    [string] $ApiToken = $env:SNIPE_TOKEN,    # Snipe-IT API token
    [string] $AssetTag,                       # e.g. LAP-001234 (if omitted, you’ll be prompted)
    [string] $PrinterName,                    # Optional: installed Zebra/Generic-Text printer name
    [switch] $PreviewOnly                     # If set, only saves label_zpl file, no print
)

$ErrorActionPreference = 'Stop'

function Get-InputIfMissing {
    param([string]$Value, [string]$Prompt)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return Read-Host $Prompt
    }
    return $Value
}

# --- Collect inputs interactively if not provided ---
$SnipeUrl  = Get-InputIfMissing $SnipeUrl  "Enter Snipe-IT base URL (e.g. https://assets.example.com)"
$ApiToken  = Get-InputIfMissing $ApiToken  "Enter Snipe-IT API token"
$AssetTag  = Get-InputIfMissing $AssetTag  "Enter Asset Tag to print"

# --- Build and call /hardware/bytag/{asset_tag} ---
# Docs: https://snipe-it.readme.io/reference/hardware-by-asset-tag
$headers = @{
  "Authorization" = "Bearer $ApiToken"       # Snipe-IT API auth: Bearer token + Accept header
  "Accept"        = "application/json"
  "Content-Type"  = "application/json"
}

$encodedTag = [uri]::EscapeDataString($AssetTag)
$uri = "$($SnipeUrl.TrimEnd('/'))/api/v1/hardware/bytag/$encodedTag?deleted=false"

try {
    $resp = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
} catch {
    Write-Error "API call failed. $_"
    exit 1
}

if (-not $resp -or $resp.status -eq 'error') {
    $msg = if ($resp) { $resp.messages } else { "No response from Snipe-IT." }
    Write-Error "Snipe-IT returned an error: $msg"
    exit 1
}

# --- Extract fields (purchase_date sometimes string or object) ---
function Get-PurchaseText($pd) {
    if ($pd -is [string]) { return $pd }
    if ($pd -is [hashtable] -or $pd -is [pscustomobject]) {
        if ($pd.formatted) { return $pd.formatted }
        if ($pd.date)      { return $pd.date }
    }
    return ""
}

$name        = $resp.name
$serial      = $resp.serial
$tag         = $resp.asset_tag
$purchase    = Get-PurchaseText $resp.purchase_date

# Safety: Remove ZPL control chars (rare, but avoid ^ or ~ in text)
function Sanitize-Zpl([string]$s) { return $s -replace '[\^~]','-' }

$name     = Sanitize-Zpl $name
$serial   = Sanitize-Zpl $serial
$tagText  = Sanitize-Zpl $tag
$purchase = Sanitize-Zpl $purchase

# --- Create ZPL (2.25" x 1.25" @ 203 dpi => ~457 x 254 dots) ---
$zpl = @"
^XA
^PW457
^LL254
^LH0,0
^CI28

^CF0,28
^FO20,15^FDName:^FS
^CF0,32
^FO140,12^FD$($name)^FS

^CF0,28
^FO20,55^FDPurchase:^FS
^CF0,30
^FO140,53^FD$($purchase)^FS

^CF0,28
^FO20,95^FDSerial:^FS
^CF0,30
^FO140,93^FD$($serial)^FS

^CF0,28
^FO20,135^FDAsset Tag:^FS
^CF0,34
^FO140,133^FD$($tagText)^FS

^BY2,2,60
^FO20,170^BCN,60,Y,N,N
^FD$($tag)^FS

^XZ
"@

# --- Save and (optionally) print ---
$out = Join-Path (Get-Location) ("label_{0}.zpl" -f ($tagText -replace '\s','_'))
$zpl | Set-Content -LiteralPath $out -Encoding ascii
Write-Host "Saved ZPL label to: $out"

if ($PreviewOnly -or -not $PrinterName) {
    Write-Host "PreviewOnly or no PrinterName provided. To print later, send the ZPL to a Zebra/Generic-Text printer."
    Write-Host "Tip: lpr -S <printerIP> -P raw $([System.IO.Path]::GetFileName($out))"
    exit 0
}

try {
    # Many Zebra drivers accept raw ZPL via Out-Printer; if not, use lpr raw on the printer's TCP queue.
    Get-Content -Raw $out | Out-Printer -Name $PrinterName
    Write-Host "Sent label to printer '$PrinterName'."
} catch {
    Write-Warning "Out-Printer failed. Consider using 'lpr -S <printerIP> -P raw ""$out""' or installing a Generic/Text Only driver."
}