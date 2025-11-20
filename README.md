# Microtech-BPLUS-Shipping
Powershell Script zum erstellen von Shipping Labels ohne "Versand und Logistik" Modul in Microtech ERP 
<img width="355" height="307" alt="355-307-max" src="https://github.com/user-attachments/assets/4cc1aee1-1403-40f8-9526-6e98d46e1bb2" />
<img width="395" height="1600" alt="labels" src="https://github.com/user-attachments/assets/0e614248-f455-48f3-9daf-1e02f335e754" />

## Feld-Validierung – Expo

<img width="1024" height="729" alt="1024-729-max" src="https://github.com/user-attachments/assets/a5146d82-c2c0-41c7-8fda-2b77c332f0ae" />

## Werte prüfen

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

<img width="845" height="574" alt="845-574-max" src="https://github.com/user-attachments/assets/6437b591-8418-4d3e-9f7e-c59a6118089d" />

Export Trigger

<img width="841" height="559" alt="841-559-max" src="https://github.com/user-attachments/assets/f66791a7-2afb-4266-a1a7-deaadef6896a" />


## CSV-Layout

### Vorspann (Mapping der Felder):

<img width="561" height="597" alt="561-597-max" src="https://github.com/user-attachments/assets/16aa95ee-f0d4-4e40-8d90-6c4ce21da75a" />

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
