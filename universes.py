"""Curated peer universes per sector/region.

All tickers are Yahoo Finance symbols (verified to return EQUITY data —
indices, ETFs, and mutual funds are filtered out by the data source as
a safety net, but the lists below are kept clean by hand.
"""

UNIVERSES = {

    # ---------------- SOFTWARE / SAAS ----------------

    "Software / SaaS — US (broad)": [
        # Large
        "MSFT", "ORCL", "CRM", "NOW", "ADBE", "INTU", "WDAY", "PANW", "FTNT",
        "ADSK", "CDNS", "SNPS",
        # Mid
        "TEAM", "ZS", "OKTA", "DDOG", "MDB", "ZM", "DOCU", "HUBS", "VEEV",
        "TWLO", "NET", "U", "S", "TYL", "ANSS",
        # Small
        "ASAN", "MNDY", "BOX", "DBX", "SMAR", "PCTY", "PAYC", "FRSH", "BILL",
        "ESTC", "GTLB", "BRZE", "PATH", "DOCN", "FROG", "PD", "ZI", "WK",
        "JAMF", "PRGS", "QLYS", "VRNS",
    ],

    "Software / SaaS — Europe": [
        # Large
        "SAP.DE",
        # Mid
        "DSY.PA",      # Dassault Systèmes
        "TEMN.SW",     # Temenos
        "HEXA-B.ST",   # Hexagon
        "NEM.DE",      # Nemetschek
        "SOP.PA",      # Sopra Steria
        # Smaller / niche
        "ATO.DE",      # Atoss Software
        "IXX.DE",      # INIT
        "GFT.DE",      # GFT Technologies
        "1U1.DE",      # 1&1
        "BC8.DE",      # Bechtle
        "CTN.DE",      # Cancom
        "TFG.DE",      # TeamViewer
        "ADN1.DE",     # Adesso
    ],

    # ---------------- AUTO PARTS / SUPPLIERS ----------------

    "Auto Parts / Suppliers — Europe/DACH": [
        # Large
        "CON.DE",      # Continental
        "SHA0.DE",     # Schaeffler
        # Mid
        "VTSC.DE",     # Vitesco Technologies
        "KBX.DE",      # Knorr-Bremse
        "DUE.DE",      # Dürr
        "NOEJ.DE",     # Norma Group
        "BDT.DE",      # Bertrandt
        "ELG.DE",      # Elmos Semiconductor (auto)
        "AUTN.SW",     # Autoneum
        "GMM.DE",      # Grammer
        # Smaller / specialty
        "PWO.DE",      # PWO AG
        "PGN.DE",      # paragon
        "KOMN.SW",     # Komax Holding
        "PYT.VI",      # Polytec Holding
        "SCE.DE",      # Schweizer Electronic
        # French/Other EU
        "FRVIA.PA",    # Forvia (Faurecia + Hella)
        "ML.PA",       # Michelin
        "FR.PA",       # Valeo
        "OPM.PA",      # OPmobility (former Plastic Omnium)
    ],

    # ---------------- INDUSTRIAL AUTOMATION / MASCHINENBAU ----------------

    "Industrial Automation / Machinery — Europe": [
        # Large
        "SIE.DE",      # Siemens
        "ABBN.SW",     # ABB
        "SU.PA",       # Schneider Electric
        "ATCO-A.ST",   # Atlas Copco
        # Mid
        "KGX.DE",      # KION Group
        "DUE.DE",      # Dürr
        "JEN.DE",      # Jenoptik
        "AIXA.DE",     # Aixtron
        "RAA.DE",      # Rational
        "VOS.DE",      # Vossloh
        "GIL.DE",      # DMG Mori
        "NWO.DE",      # Nordex
        "PFV.DE",      # Pfeiffer Vacuum
        # Smaller
        "WCH.DE",      # Wacker Chemie
        "SDF.DE",      # K+S
    ],

    # ---------------- MEDTECH / HEALTHCARE EQUIPMENT ----------------

    "MedTech / Healthcare — Europe": [
        # Large
        "SHL.DE",      # Siemens Healthineers
        "EL.PA",       # EssilorLuxottica
        "PHIA.AS",     # Philips
        # Mid — medical devices / diagnostics
        "STMN.SW",     # Straumann (dental implants)
        "SRT3.DE",     # Sartorius Vz (bioprocess) — use one share class only
        "QIA.DE",      # Qiagen (diagnostics)
        # Denmark cluster — pure medical devices
        "GN.CO",       # GN Store Nord (hearing)
        "DEMANT.CO",   # Demant (hearing)
        "COLO-B.CO",   # Coloplast (ostomy/continence)
        "DRW3.DE",     # Drägerwerk (medical/safety)
    ],

    # ---------------- SEMICONDUCTORS ----------------

    "Semiconductors — Global": [
        "NVDA", "AMD", "INTC", "AVGO", "QCOM", "TXN", "MU", "MRVL", "AMAT",
        "LRCX", "KLAC", "ASML", "TSM", "NXPI", "ON", "MCHP", "ENPH",
        # Europe-listed
        "ASML.AS", "STMPA.PA", "INF.DE", "ELG.DE", "AMS.SW", "AIXA.DE",
        "SOI.PA", "XFAB.PA",
    ],

    # ---------------- CONSUMER BRANDS ----------------

    "Consumer Brands / Luxury — Europe": [
        # Luxury
        "MC.PA",       # LVMH
        "KER.PA",      # Kering
        "CFR.SW",      # Richemont
        "RMS.PA",      # Hermès
        "MONC.MI",     # Moncler
        "BRBY.L",      # Burberry
        # Beverage / FMCG
        "ABI.BR",      # AB InBev
        "HEIA.AS",     # Heineken
        "DGE.L",       # Diageo
        "ULVR.L",      # Unilever
        "NESN.SW",     # Nestlé
        "RKT.L",       # Reckitt
        "BN.PA",       # Danone
        "RI.PA",       # Pernod Ricard
        "BVI.PA",      # Bureau Veritas
    ],

    # ---------------- RENEWABLES / CLEANTECH ----------------

    "Renewables / Cleantech — Europe": [
        "VWS.CO",      # Vestas
        "ORSTED.CO",   # Ørsted
        "NWO.DE",      # Nordex
        "ENR.DE",      # Siemens Energy
        "S92.DE",      # SMA Solar
        "EDPR.LS",     # EDP Renováveis
        "IBE.MC",      # Iberdrola
        "EOAN.DE",     # E.ON
        "RWE.DE",      # RWE
        "ENI.MI",      # Eni
    ],

    # ---------------- PHARMA / BIOTECH ----------------

    "Pharma / Biotech — Global": [
        # Big Pharma
        "LLY", "JNJ", "MRK", "ABBV", "PFE", "BMY", "GILD",
        "NOVN.SW",     # Novartis
        "RO.SW",       # Roche
        "AZN.L",       # AstraZeneca
        "GSK.L",       # GSK
        "SAN.PA",      # Sanofi
        "NOVO-B.CO",   # Novo Nordisk
        "BAYN.DE",     # Bayer
        # Biotech / Specialty
        "GMAB.CO",     # Genmab
        "UCB.BR",      # UCB
        "IPN.PA",      # Ipsen
    ],

    # ---------------- AEROSPACE & DEFENSE ----------------

    "Aerospace & Defense — Global": [
        # US
        "BA", "RTX", "LMT", "GD", "NOC", "HON", "TXT", "HWM", "TDG",
        # Europe
        "SAF.PA",      # Safran
        "AIR.PA",      # Airbus
        "BA.L",        # BAE Systems
        "LDO.MI",      # Leonardo
        "RHM.DE",      # Rheinmetall
        "HO.PA",       # Thales
    ],

    # ---------------- CHEMICALS ----------------

    "Chemicals — Global": [
        # US
        "LIN", "DOW", "DD", "LYB", "PPG", "ECL", "APD", "ALB", "EMN",
        # Europe
        "BAS.DE",      # BASF
        "AKZA.AS",     # Akzo Nobel
        "EVK.DE",      # Evonik
        "LXS.DE",      # Lanxess
        "SIKA.SW",     # Sika
        "GIVN.SW",     # Givaudan
    ],

    # ---------------- BUSINESS SERVICES ----------------

    "Business Services — Global": [
        # US
        "ADP", "PAYX", "SPGI", "MCO", "VRSK", "BAH",
        # Europe
        "WKL.AS",      # Wolters Kluwer
        "REL.L",       # RELX
        "BNZL.L",      # Bunzl
        "RTO.L",       # Rentokil
        "EXPN.L",      # Experian
        "ADEN.SW",     # Adecco
    ],

    # ---------------- LOGISTICS / TRANSPORT ----------------

    "Logistics / Transport — Global": [
        # US
        "FDX", "UPS", "GXO", "CHRW", "EXPD", "ODFL", "JBHT",
        # Europe
        "DHL.DE",      # DHL Group (Deutsche Post)
        "MAERSK-B.CO", # Maersk
        "KNIN.SW",     # Kuehne + Nagel
        "DSV.CO",      # DSV
    ],

    # ---------------- PAYMENTS / FINTECH ----------------

    "Payments / Fintech — Global": [
        # US
        "V", "MA", "PYPL", "FIS", "FISV", "GPN", "WU", "FOUR", "TOST",
        # Europe
        "ADYEN.AS",    # Adyen
        "NEXI.MI",     # Nexi
        "WLN.PA",      # Worldline
    ],
}


def list_universes():
    return list(UNIVERSES.keys())


def get_universe(name: str):
    return UNIVERSES.get(name, [])
