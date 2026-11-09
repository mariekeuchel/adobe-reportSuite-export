# Adobe Analytics SDR-Export

Exportiert eVars, Props, List-Variablen, Events, Segmente und berechnete Metriken
einer oder mehrerer Report Suites in **eine Excel-Datei**, inklusive Beschreibungen,
Allokation, Expiration, Scope und Segmentdefinition als lesbarem Text.

## 1. Einmalige Einrichtung in Adobe

1. **Adobe Developer Console** (developer.adobe.com/console): neues Projekt anlegen,
   dann *Add API* → **Adobe Analytics** → **OAuth Server-to-Server**.
2. Ein **Produktprofil** zuweisen, das
   - Zugriff auf die gewünschten Report Suites hat und
   - **Report Suite Tools / Admin-Rechte** besitzt. Das ist für die 1.4-Admin-Methoden
     (Allokation, Expiration usw.) nötig und damit `includeType=all` alle Segmente liefert.
3. Den technischen Account im **Admin Console**-Produktprofil als **Product Admin**
   bzw. Analytics-Admin führen. Sonst sieht er nur eigene oder geteilte Segmente.
4. *Client ID* und *Client Secret* aus der Developer Console kopieren.

## 2. Lokal einrichten

```bash
pip install -r requirements.txt
cp config.example.json config.json   # Werte eintragen, Datei NICHT ins Git committen
```

`global_company_id` kann leer bleiben, dann wird sie automatisch ermittelt.

## 3. Ausführen

```bash
# verfügbare Report Suites anzeigen
python adobe_sdr_export.py --list-rsids

# Export für die Report Suites aus config.json
python adobe_sdr_export.py

# bestimmte Report Suites, eigener Dateiname
python adobe_sdr_export.py --rsids rsid_prod rsid_app --out sdr_2026-11.xlsx

# alle Segmente/Calc Metrics der Company, nicht nur die der gewählten RS
python adobe_sdr_export.py --all-segments

# ohne 1.4 API, z. B. wenn die Admin-Rechte fehlen
python adobe_sdr_export.py --skip-14
```

## 4. Inhalt der Excel-Datei

| Tab | Quelle | Wichtige Spalten |
|---|---|---|
| Übersicht | beide | Anzahl Komponenten je Report Suite, Exportzeitpunkt |
| eVars | 1.4 + 2.0 | name, description, `allocation_type`, `expiration_type`, `expiration_custom_days`, `merchandising_syntax`, enabled, description_2_0 |
| Props | 1.4 + 2.0 | pathing, list prop, participation, case sensitivity |
| Listvars | 1.4 | Delimiter, Allokation, Expiration |
| Events | 1.4 + 2.0 | `type`, `serialization`, `participation`, `polarity`, visibility |
| Segmente | 2.0 | **scope_container** (Hit/Visit/Visitor), **definition_lesbar**, verwendete_variablen, verwendete_segmente, owner, tags, shares, definition_json |
| Calc Metrics | 2.0 | **formel_lesbar**, verwendete Komponenten/Segmente, owner, tags |
| Dimensionen (alle) / Metriken (alle) | 2.0 | vollständige Liste inkl. Standard-Dimensionen, Kategorie, Tags |

Jeder Tab hat am Ende leere Spalten **Business-Definition, Fachlicher Owner, Status, Kommentar**
für deinen Zusatzkontext.

Beispiel für `definition_lesbar`:

```
Visit-Container [ (Device [evar5] equals 'mobile' AND (Login [event1] exists OR ...)
  AND EXCLUDE ( Hit-Container [ Page [page] contains 'test' ] )) ]
```

## Hinweise

- Wenn die **1.4-Aufrufe fehlschlagen** (403/401), läuft das Skript trotzdem weiter und
  füllt eVars, Props und Events aus der 2.0 API, dann aber ohne Allokation und Expiration.
  Die Warnung steht im Log.
- **Segmente mit `--rsids`** liefert nur Segmente, die auf diesen Report Suites *angelegt*
  wurden. Segmente sind aber RS-übergreifend nutzbar. Für eine vollständige Liste
  `--all-segments` verwenden.
- Bei sehr exotischen Segment-Konstrukten (z. B. Zeitbeschränkungen in Sequenzen) wird der
  betreffende Teil als JSON in die lesbare Definition übernommen. Das komplette Original
  steht immer in `definition_json`.
- Bei Rate-Limits (HTTP 429) wartet das Skript automatisch und versucht es erneut.
