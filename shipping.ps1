###############################################################################
# CONFIGURATION
###############################################################################

# ---- PRINTER ----
$PrinterIP = "192.168.1.22"
$PrinterPort = 9100

# ---- FILE PATHS ----
$CsvFile        = "\\bwerp01.domain.local\Shipping-Schnittstelle\sendungen.csv"
$OutputFolder   = "\\bwerp01.domain.local\Shipping-Schnittstelle\Labels"
$FinishedFolder = "\\bwerp01.domain.local\Shipping-Schnittstelle\finished"

# Ensure folders exist
foreach ($folder in @($OutputFolder, $FinishedFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
    }
}

# ---- SENDER INFORMATION (shared UPS + DPD) ----
$SenderName       = "Firmenname Gmbh"
$SenderStreet     = "Firmenstrasse 51"
$SenderPostalCode = "83123"
$SenderCity       = "Stadt Name"
$SenderCountry    = "DE"
$SenderPhone      = "+4980123456"

# ---- UPS AUTH CONFIG ----
$UpsClientID     = "xxxxxxxxXXXXXsadfadasdasd3342"
$UpsClientSecret = "xxxxxXXXxxxXXXxXXxxxxXXX213"
$UpsShipperNumber = "123465"

$UpsTokenUrl = "https://onlinetools.ups.com/security/v1/oauth/token"
$UpsShipUrl  = "https://onlinetools.ups.com/api/shipments/v1/ship"

# ---- DPD AUTH CONFIG ----
$DpdDelisId   = "sandboxdpd"
$DpdPassword  = "-putsanboxpassword-"
$DpdDepot     = "0184"

$DpdLoginWsdl  = "https://public-ws-stage.dpd.com/services/LoginService/V2_0/?wsdl"
$DpdShipmentUrl = "https://public-ws-stage.dpd.com/services/ShipmentService/V4_4"
$DpdSoapAction  = "http://dpd.com/common/service/ShipmentService/4.4/storeOrders"

###############################################################################
# HELPER FUNCTIONS
###############################################################################

function Normalize-Umlauts {
    param($text)
    if (-not $text) { return "" }
    return $text.Normalize([Text.NormalizationForm]::FormC)
}

function Show-ServerResponse {
    param(
        $response,
        [string]$label
    )

    Write-Host "---- $label ----" -ForegroundColor Cyan

    # On Invoke-RestMethod error
    if ($response -is [System.Management.Automation.ErrorRecord]) {

        if ($response.ErrorDetails.Message) {
            Write-Host $response.ErrorDetails.Message
            Write-Host "-----------------------------"
            return
        }

        if ($response.Exception.Response -and $response.Exception.Response.GetResponseStream) {
            $reader = New-Object System.IO.StreamReader($response.Exception.Response.GetResponseStream())
            $raw = $reader.ReadToEnd()
            Write-Host $raw
            Write-Host "-----------------------------"
            return
        }

        Write-Host "(no response body available)"
        Write-Host "-----------------------------"
        return
    }

    try {
        ($response | ConvertTo-Json -Depth 50) | Write-Host
    }
    catch {
        Write-Host $response
    }

    Write-Host "-----------------------------"
}

function Send-ZPL {
    param(
        [string]$printerIP,
        [int]$port,
        [string]$zpl
    )

    Write-Host "Sending ZPL to printer $printerIP ..."
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $client.Connect($printerIP, $port)
        $stream = $client.GetStream()
        $writer = New-Object System.IO.StreamWriter($stream)
        $writer.Write($zpl)
        $writer.Flush()
        $writer.Close()
        $client.Close()
        Write-Host "✔ ZPL sent to printer" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to send ZPL to printer $printerIP" -ForegroundColor Red
        Write-Host $_
    }
}

function Save-And-Print-Label {
    param(
        $row,
        [string]$carrier,
        [string]$tracking,
        [string]$zpl
    )

    $fileName = "$($row.belegnummer)-$tracking-$carrier.zpl"
    $filePath = Join-Path $OutputFolder $fileName

    Write-Host "Saving ZPL label to: $filePath"
    $zpl | Out-File -FilePath $filePath -Encoding ASCII

    Send-ZPL -printerIP $PrinterIP -port $PrinterPort -zpl $zpl

    return $filePath
}

###############################################################################
# UPS FUNCTIONS
###############################################################################

function Get-UpsToken {
    Write-Host "UPS: Requesting OAuth token..."

    $pair   = "$UpsClientID`:$UpsClientSecret"
    $bytes  = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [Convert]::ToBase64String($bytes)

    $headers = @{
        "Authorization" = "Basic $base64"
        "Content-Type"  = "application/x-www-form-urlencoded"
    }

    $body = "grant_type=client_credentials"

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $UpsTokenUrl -Headers $headers -Body $body -ErrorAction Stop
        Write-Host "UPS: OAuth token OK"
        return $resp.access_token
    }
    catch {
        Show-ServerResponse -response $_ -label "UPS TOKEN ERROR"
        throw "UPS token request failed."
    }
}

function Process-UpsShipment {
    param($row)

    Write-Host "UPS: Starting shipment for order $($row.belegnummer)..."

    $token = Get-UpsToken

    Write-Host "UPS: Building request body..."

    switch ($row.service.ToLower()) {
        "standard"       { $serviceCode = "11" }
        "express saver"  { $serviceCode = "65" }
        "express"        { $serviceCode = "07" }
        default          { $serviceCode = "11" }
    }

    $pkgCount = 1
    if ($row.anzahlpakete -and [int]$row.anzahlpakete -gt 0) {
        $pkgCount = [int]$row.anzahlpakete
    }

    # Normalize text
    $row.name          = Normalize-Umlauts $row.name
    $row.nameaddition  = Normalize-Umlauts $row.nameaddition
    $row.strasse       = Normalize-Umlauts $row.strasse
    $row.Ort           = Normalize-Umlauts $row.Ort

    $senderNameUPS   = Normalize-Umlauts $SenderName
    $senderStreetUPS = Normalize-Umlauts $SenderStreet
    $senderCityUPS   = Normalize-Umlauts $SenderCity

    # Weight (UPS expects dot)
    $weightText = $row.gewicht.ToString().Replace(",", ".")

    # Build packages array
    $packages = @()
    for ($i = 1; $i -le $pkgCount; $i++) {
        $packages += @{
            Packaging = @{ Code = "02" }
            PackageWeight = @{
                UnitOfMeasurement = @{ Code = "KGS" }
                Weight = $weightText
            }
        }
    }

    # ShipTo object
    $shipToObj = @{
        Name          = $row.name
        AttentionName = $row.nameaddition
        Address       = @{
            AddressLine = @($row.strasse)
            City        = $row.Ort
            PostalCode  = $row.plz
            CountryCode = $row.country
        }
    }
    if ($row.telefonnummer -and $row.telefonnummer.Trim() -ne "") {
        $shipToObj.Phone = @{ Number = $row.telefonnummer }
    }

    $bodyObj = @{
        ShipmentRequest = @{
            Request = @{
                RequestOption = "nonvalidate"
                TransactionReference = @{
                    CustomerContext = "Order $($row.belegnummer)"
                }
            }
            Shipment = @{
                Shipper = @{
                    Name          = $senderNameUPS
                    ShipperNumber = $UpsShipperNumber
                    Phone         = @{ Number = $SenderPhone }
                    Address       = @{
                        AddressLine = @($senderStreetUPS)
                        City        = $senderCityUPS
                        PostalCode  = $ShipperPostalCode
                        CountryCode = $ShipperCountry
                    }
                }
                ShipTo = $shipToObj
                PaymentInformation = @{
                    ShipmentCharge = @{
                        Type        = "01"
                        BillShipper = @{ AccountNumber = $UpsShipperNumber }
                    }
                }
                Service = @{
                    Code = $serviceCode
                }
                Package = $packages
            }
            LabelSpecification = @{
                LabelImageFormat = @{
                    Code = "ZPL"
                }
                LabelStockSize = @{
                    Width  = "4"
                    Height = "6"
                }
                HTTPUserAgent = "Mozilla/5.0"
            }
        }
    }

    $jsonBody = $bodyObj | ConvertTo-Json -Depth 20
    Write-Host "UPS: Request JSON:"
    Write-Host $jsonBody

    $headers = @{
        "Authorization" = "Bearer $token"
    }

    Write-Host "UPS: Sending shipment request..."
    try {
        $utf8body = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)

        $resp = Invoke-RestMethod `
            -Method Post `
            -Uri $UpsShipUrl `
            -Headers $headers `
            -Body $utf8body `
            -ContentType "application/json; charset=utf-8" `
            -ErrorAction Stop

        Show-ServerResponse -response $resp -label "UPS SUCCESS RESPONSE"
    }
    catch {
        Show-ServerResponse -response $_ -label "UPS ERROR RESPONSE"
        throw "UPS shipment failed."
    }

    $tracking = $resp.ShipmentResponse.ShipmentResults.ShipmentIdentificationNumber
    Write-Host "UPS: Tracking number: $tracking"

    $labelNode = $resp.ShipmentResponse.ShipmentResults.PackageResults.ShippingLabel
    $zplBytes  = [Convert]::FromBase64String($labelNode.GraphicImage)
    $zpl       = [System.Text.Encoding]::ASCII.GetString($zplBytes)

    return [pscustomobject]@{
        Carrier  = "UPS"
        Tracking = $tracking
        Zpl      = $zpl
    }
}

###############################################################################
# DPD FUNCTIONS
###############################################################################

function Get-DpdAuthToken {
    Write-Host "DPD: Requesting auth token via SOAP LoginService..."

    $loginProxy = New-WebServiceProxy -Uri $DpdLoginWsdl -Namespace "DPDLogin"

    try {
        $loginResult = $loginProxy.getAuth($DpdDelisId, $DpdPassword, "de_DE")
        $authToken   = $loginResult.authToken
        Write-Host "DPD: Auth token OK"
        return $authToken
    }
    catch {
        Write-Host "DPD LOGIN ERROR:"
        Write-Host $_
        throw "DPD login failed."
    }
}

function Process-DpdShipment {
    param($row)

    Write-Host "DPD: Starting shipment for order $($row.belegnummer)..."

    $authToken = Get-DpdAuthToken

    Write-Host "DPD: Building SOAP request body..."

    $identificationNumber = $row.belegnummer
    $sendingDepot = $DpdDepot
    $product = "CL"

    # Use same sender data as UPS
    $senderName   = $SenderName
    $senderStreet = $SenderStreet
    $senderZip    = $SenderPostalCode
    $senderCity   = $SenderCity
    $senderCountry = $SenderCountry

    $recName    = $row.name
    $recStreet  = $row.strasse
    $recZip     = $row.plz
    $recCity    = $row.Ort
    $recCountry = $row.country

    # gewicht in CSV is KG → convert to grams (int)
    $gewichtKg = [double]($row.gewicht.ToString().Replace(",", "."))
    $weight    = [int]([math]::Round($gewichtKg * 1000))

$soapBody = @"
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:ns="http://dpd.com/common/service/types/Authentication/2.0"
    xmlns:ns1="http://dpd.com/common/service/types/ShipmentService/4.4">
  <soapenv:Header>
    <ns:authentication>
      <delisId>$DpdDelisId</delisId>
      <authToken>$authToken</authToken>
      <messageLanguage>de_DE</messageLanguage>
    </ns:authentication>
  </soapenv:Header>
  <soapenv:Body>
    <ns1:storeOrders>
      <printOptions>
        <printOption>
          <outputFormat>ZPL</outputFormat>
          <paperFormat>A6</paperFormat>
          <startPosition>UPPER_LEFT</startPosition>
        </printOption>
      </printOptions>
      <order>
        <generalShipmentData>
          <identificationNumber>$identificationNumber</identificationNumber>
          <sendingDepot>$sendingDepot</sendingDepot>
          <product>$product</product>
            <sender>
                <name1>$senderName</name1>
                <street>$senderStreet</street>
                <country>$senderCountry</country>
                <zipCode>$senderZip</zipCode>
                <city>$senderCity</city>
            </sender>
          <recipient>
            <name1>$recName</name1>
            <street>$recStreet</street>
            <country>$recCountry</country>
            <zipCode>$recZip</zipCode>
            <city>$recCity</city>
          </recipient>
        </generalShipmentData>
        <parcels>
          <weight>$weight</weight>
        </parcels>
        <productAndServiceData>
          <orderType>consignment</orderType>
        </productAndServiceData>
      </order>
    </ns1:storeOrders>
  </soapenv:Body>
</soapenv:Envelope>
"@

    Write-Host "DPD: Sending SOAP request to ShipmentService..."

    $headers = @{
        "Content-Type" = "text/xml;charset=UTF-8"
        "SOAPAction"   = $DpdSoapAction
    }

    try {
        $response = Invoke-WebRequest `
            -Uri $DpdShipmentUrl `
            -Method POST `
            -Headers $headers `
            -Body $soapBody `
            -ContentType "text/xml; charset=utf-8" `
            -ErrorAction Stop
    }
    catch {
        Write-Host "❌ DPD ERROR RESPONSE (HTTP):"
        if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $raw = $reader.ReadToEnd()
            Write-Host $raw
        } else {
            Write-Host $_
        }
        throw "DPD shipment failed."
    }

    Write-Host "DPD: Parsing SOAP response..."

    [xml]$xml = $response.Content

    $contentNode = $xml.GetElementsByTagName("content") | Select-Object -First 1
    if (-not $contentNode) {
        Write-Host "❌ DPD: No <content> element found in response!"
        Write-Host $response.Content
        throw "DPD response has no label content."
    }

    # DPD content is Base64-encoded ZPL
    $zplBytes = [Convert]::FromBase64String($contentNode.InnerText)
    $zpl      = [System.Text.Encoding]::ASCII.GetString($zplBytes)

    $trackingNode = $xml.GetElementsByTagName("parcelLabelNumber") | Select-Object -First 1
    if (-not $trackingNode) {
        Write-Host "⚠ DPD: No parcelLabelNumber found, using 'UNKNOWN' as tracking."
        $tracking = "UNKNOWN"
    } else {
        $tracking = $trackingNode.InnerText
    }

    Write-Host "DPD: Tracking number: $tracking"

    return [pscustomobject]@{
        Carrier  = "DPD"
        Tracking = $tracking
        Zpl      = $zpl
    }
}

###############################################################################
# MAIN SCRIPT FLOW
###############################################################################

Write-Host "================ SHIPPING SCRIPT START ================" -ForegroundColor Cyan
Write-Host "Loading CSV: $CsvFile"

if (-not (Test-Path $CsvFile)) {
    Write-Host "❌ CSV file not found: $CsvFile" -ForegroundColor Red
    exit 1
}

$rows = @(Import-Csv -Path $CsvFile -Delimiter ';')


if ($rows.Count -eq 0) {
    Write-Host "No rows in CSV. Nothing to do."
    exit 0
}

$lastResult = $null

foreach ($row in $rows) {
    Write-Host "------------------------------------------------------"
    Write-Host "Order: $($row.belegnummer)"
    Write-Host "Carrier (versanddienstleister): $($row.versanddienstleister)"
    Write-Host "Service: $($row.service)"
    Write-Host "Weight (kg): $($row.gewicht)"
    Write-Host "Packages: $($row.anzahlpakete)"
    Write-Host "------------------------------------------------------"

    try {
        $carrierValue = $row.versanddienstleister
        $result = $null

        switch ($carrierValue.ToUpper()) {
            "UPS" {
                Write-Host "Selected carrier: UPS" -ForegroundColor Green
                $result = Process-UpsShipment -row $row
            }
            "DPD" {
                Write-Host "Selected carrier: DPD" -ForegroundColor Green
                $result = Process-DpdShipment -row $row
            }
            default {
                Write-Host "❌ Unknown service provider in CSV: '$carrierValue'" -ForegroundColor Red
                continue
            }
        }

        if ($null -ne $result) {
            $labelPath = Save-And-Print-Label -row $row -carrier $result.Carrier -tracking $result.Tracking -zpl $result.Zpl
            $lastResult = @{
                tracking = $result.Tracking
                label    = $labelPath
                carrier  = $result.Carrier
            }

            Write-Host "✔ Label created: $labelPath"
            Write-Host "✔ Tracking number: $($result.Tracking)"
        } else {
            Write-Host "⚠ No result for order $($row.belegnummer)"
        }
    }
    catch {
        Write-Host "❌ Failed order: $($row.belegnummer)" -ForegroundColor Red
        Write-Host $_
        continue
    }
}

###############################################################################
# MOVE CSV TO FINISHED FOLDER
###############################################################################

if ($rows.Count -gt 0 -and $lastResult -ne $null) {
    $first = $rows[0]
    $trackingForFile = $lastResult.tracking
    $carrierForFile  = $lastResult.carrier

    $targetCsvName = "$($first.belegnummer)-$trackingForFile-$carrierForFile.csv"
    $targetCsvPath = Join-Path $FinishedFolder $targetCsvName

    Write-Host "Moving CSV to: $targetCsvPath"
    Move-Item -Force -Path $CsvFile -Destination $targetCsvPath
    Write-Host "✔ CSV moved."
}
else {
    Write-Host "No successful labels – CSV not moved."
}

Write-Host "================ SHIPPING SCRIPT END =================" -ForegroundColor Cyan

