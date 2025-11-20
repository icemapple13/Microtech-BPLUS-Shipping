# Microtech-BPLUS-Shipping
Powershell Script zum erstellen von Shipping Labels ohne "Versand und Logistik" Modul in Microtech ERP 
![Schnellwahl Versandlabel Erstellen](.attachments.53426/image.png)

## Feld-Validierung – Expo

![1024-729.png](.attachments.53426/1024-729.png)

## rt prüfen

Über eine verzweigte Formel definieren  
Diese Formel wird genutzt, um zu prüfen, ob alle notwendigen Felder vorhanden sind:

```
$cond(
  $length(«Vsd.Na2»)=0, "fail-name",
  $cond(
    $length(«Vsd.Str»)=0, "fail-strasse",
    $cond(
      $length(«Vsd.Land»)=0, "fail-land",
      $cond(
        $length(«Vsd.PLZ»)=0, "fail-plz",
        $cond(
          $length(«Vsd.Gew»)=0, "fail-gewicht",
          $cond(
            «Vsd.Gew» > 0,
            $cond(
              $length(«Vsd.BelegNr»)=0, "warn-beleg",
              ""
            ),
            "fail-gewicht"
          )
        )
      )
    )
  )
)
```

### Beispiel-Fehleranzeige

![image (6).png](.attachments.53426/image%20%286%29.png)

Export Trigger

![Export übersicht](.attachments.53426/image%20%284%29.png)

## CSV-Layout

### Vorspann (Mapping der Felder):

![Vorspann](.attachments.53426/561-597-max.png)

```
id;anrede;name;nameaddition;country;plz;Ort;strasse;telefonnummer;emailaddress;upskundennummer;versanddienstleister;service;gewicht;anzahlpakete;belegnummer;isprivateaddresse;
```

### Layout der Exportdaten:

```
«Vsd.ID»;«Vsd.Na1»;«Vsd.Na2»;«Vsd.Na3»;«Vsd.LandISOKennz»;«Vsd.PLZ»;«Vsd.Ort»;«Vsd.Str»;«Vsd.Tel1»;«Vsd.EMail1»;«Vsd.VsA.KdNr»;«Vsd.VsA.Versender»;«Vsd.VsA.VsdArt»;«Vsd.Gew»;«Vsd.AzPakete»;«Vsd.BelegNr»;«Formel\[;;\]»;
```

### Privatadresse bestimmen:

\-> Formel um zu bestimmen ob Priv Addresse wichtig bei UPS 

```
(«Vsd.Anr»='Firma',FALSE,TRUE)
```

# 

## 🗂️ Ordnerstruktur

```
Shipping-Schnittstelle/
│
├── sendungen.csv               ← Eingabedatei
├── shipping.ps1                ← Powershell Versandscript
│
├── Labels/                     ← erzeugte ZPL-Labeldateien
│     └── *.zpl
│
└── finished/                   ← verarbeitete CSV-Dateien
      └── belegnummer-tracking-carrier.csv
```

Der Share wird über UNC bereitgestellt:

```
\\bwerp01.baw.local\Shipping-Schnittstelle
```

Alle berechtigten Benutzer müssen Schreibrechte haben.  
Shipping-schnittstelle Freigabe UNC mit allen Usern berechtigt die das bedienen sollen.

script configuration:  
  
###############################################################################

# CONFIGURATION

###############################################################################

# Drucker IP und Port setzen (nur ZPL fähige drucker)

$PrinterIP = "192.168.1.22"
$PrinterPort = 9100

# ---- FILE PATHS ----

$CsvFile        = "\\\\ServerName.domain.local\\Shipping-Schnittstelle\\sendungen.csv" 

$OutputFolder   = "\\\\ServerName.domain.local\\Shipping-Schnittstelle\\Labels" 

$FinishedFolder = "\\\\ServerName.domain.local\\Shipping-Schnittstelle\\finished"

# Ensure folders exist

foreach ($folder in @($OutputFolder, $FinishedFolder)) {
if (-not (Test-Path $folder)) {
New-Item -ItemType Directory -Force -Path $folder | Out-Null
}
}

# ---- SENDER INFORMATION (shared UPS + DPD) ----

$SenderName       = "DemoFirma"   
$SenderStreet     = "Demosrasse 3"   
$SenderPostalCode = "83000"   
$SenderCity       = "Rosenheim"   
$SenderCountry    = "DE"   
$SenderPhone      = "+4980123456"

# ---- UPS AUTH CONFIG ----

$UpsClientID     = "XXXxxxxXXXXX"   
$UpsClientSecret = "ccccXxxxxxxxXXXX"   
$UpsShipperNumber = "123456"

$UpsTokenUrl = "https://onlinetools.ups.com/security/v1/oauth/token"
$UpsShipUrl  = "https://onlinetools.ups.com/api/shipments/v1/ship"

# ---- DPD AUTH CONFIG ----

$DpdDelisId   = "sandboxdpd"
$DpdPassword  = "xMmshh1"
$DpdDepot     = "0184"

$DpdLoginWsdl  = "https://public-ws-stage.dpd.com/services/LoginService/V2_0/?wsdl"
$DpdShipmentUrl = "https://public-ws-stage.dpd.com/services/ShipmentService/V4_4"
$DpdSoapAction  = "http://dpd.com/common/service/ShipmentService/4.4/storeOrders"
