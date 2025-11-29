# 💻 Dell Laptop Vergleichs-Tool

Ein automatisches Python-Tool zum Extrahieren und Vergleichen von Dell Laptop-Daten aus PDF-Dokumenten.

## ✨ Funktionen

- 📄 **PDF-Import:** Extrahiert Daten aus PDF-Dateien
- 🖼️ **Bild-Extraktion:** Speichert Produktbilder organisiert
- 💰 **Preis-Analyse:** Automatische Bewertung basierend auf Rabatten
- 📊 **Multi-Format Export:** CSV, Excel und Markdown-Berichte
- ⭐ **Bewertungssystem:** Visualisierung mit Emojis

## 🚀 Verwendung

```bash
# Virtual Environment aktivieren
source venv/bin/activate

# Script ausführen
python extract_laptop_data.py
```

## 📈 Ergebnisse

Das Skript generiert automatisch:
- `output/laptops_vergleich.csv` - CSV-Tabelle
- `output/laptops_vergleich.xlsx` - Excel-Datei mit Formatierung
- `output/laptops_vergleich.md` - **[Vollständiger Markdown-Report](output/laptops_vergleich.md)** 📋
- `images/` - Extrahierte Bilder aus den PDFs

## 📊 Aktueller Stand

- 💻 **Laptops analysiert:** 11
- 📸 **Bilder extrahiert:** 896
- 💰 **Durchschnittspreis:** €889.08
- 🏆 **Beste Angebote:** 5 Laptops mit >60% Rabatt
- 🟢 **Sehr gute Deals:** 5 Laptops
- 🟡 **Faire Deals:** 4 Laptops
- 🔴 **Zu teuer:** 1 Laptop

## 🎯 Preis-Bewertungssystem

- 🟢 **Sehr gut:** >60% Rabatt
- 🟢 **Gut:** 50-60% Rabatt
- 🟡 **Fair:** 40-50% Rabatt
- 🟡 **Akzeptabel:** 30-40% Rabatt
- 🔴 **Zu teuer:** <30% Rabatt

## 💡 Installation

### Abhängigkeiten

```bash
# Virtual Environment erstellen
python3 -m venv venv

# Aktivieren
source venv/bin/activate  # Linux/Mac
# oder
venv\Scripts\activate  # Windows

# Pakete installieren
pip install pymupdf pandas openpyxl pillow
```

### Benötigte Pakete

- PyMuPDF (fitz) - PDF-Verarbeitung
- pandas - Datenanalyse
- Pillow - Bildverarbeitung
- openpyxl - Excel-Export

## 📋 Vollständiger Laptop-Vergleich

**[➡️ Zum vollständigen Vergleichsbericht mit Bildern](output/laptops_vergleich.md)**

Der vollständige Bericht enthält:
- 📸 Produktbilder für jeden Laptop
- 🔍 Detaillierte technische Spezifikationen
- 💰 Preis-Leistungs-Analyse
- 🔗 Verlinkte Schnellvergleich-Tabelle
- ⭐ Bewertungen und Empfehlungen

## 🛠️ Projektstruktur

```
.
├── extract_laptop_data.py    # Hauptskript
├── *.pdf                      # Quell-PDFs (17 Laptop + 5 Cover)
├── cover/                     # Cover-Design PDFs
├── images/                    # Extrahierte Bilder (896 total)
│   ├── Dell Latitude */
│   ├── Dell Precision */
│   └── ...
└── output/                    # Generierte Berichte
    ├── laptops_vergleich.csv
    ├── laptops_vergleich.xlsx
    └── laptops_vergleich.md
```

## 📜 Lizenz

Privates Projekt

---

*Generiert mit Claude Code - [https://claude.com/claude-code](https://claude.com/claude-code)*
