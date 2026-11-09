# Adobe Analytics SDR-Export

Exportiert eVars, Props, List-Variablen, Events, Segmente und berechnete Metriken
einer oder mehrerer Report Suites in **eine Excel-Datei**, inklusive Beschreibungen,
eVar-Allokation und -Expiration, Scope und Segmentdefinition als lesbarem Text.

> **Hinweis:** Adobe hat die Analytics 1.4 API im August 2026 abgeschaltet
> (`api.omniture.com` antwortet nicht mehr). Das Skript nutzt deshalb nur noch die
> 2.0 API. Diese liefert Konfigurationsdetails **nur für eVars**. Für Props, Events
> und List-Variablen gibt es nur Name, Beschreibung und Kategorie
> ([Migration Guide](https://developer.adobe.com/analytics-apis/docs/2.0/guides/migration)).

## 1. Einmalige Einrichtung in Adobe

1. **Adobe Developer Console** (developer.adobe.com/console): neues Projekt anlegen,
   dann *Add API* → **Adobe Analytics** → **OAuth Server-to-Server**.
2. Ein **Produktprofil** zuweisen, das
   - Zugriff auf die gewünschten Report Suites hat und
   - **Admin-Rechte** besitzt, damit `includeType=all` alle Segmente liefert.
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
```

## 4. Inhalt der Excel-Datei

| Tab | Quelle | Wichtige Spalten |
|---|---|---|
| Übersicht | 2.0 | Anzahl Komponenten je Report Suite, davon eVars mit Einstellungen, Exportzeitpunkt |
| eVars | 2.0 | name, description, `allocationType`, `expirationType`, `expirationCustomDays`, `dataType` (text/counter), `merchandisingSyntax`, `bindingEvents`, `ecomAllocationType`, `attributionModel` |
| Props | 2.0 | name, description, Kategorie, pathable (keine Prop-Einstellungen in der 2.0 API) |
| Listvars | 2.0 | name, description (Delimiter, Allokation, Expiration nicht in der 2.0 API) |
| Events | 2.0 | name, description, type, polarity, Kategorie (Serialisierung/Participation nicht in der 2.0 API) |
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

- Liefert die 2.0 API **keine eVar-Einstellungen** (z. B. weil Adobe einen
  `expansion`-Wert nicht kennt), läuft das Skript trotzdem weiter und schreibt eine
  Warnung ins Log. Die Übersicht zeigt in „eVars mit Einstellungen", wie viele eVars
  Allokation/Expiration enthalten.
- Einstellungen für Props, Events und List-Variablen sind über keine aktuelle Adobe-API
  mehr abrufbar, nur noch in der Oberfläche unter *Admin > Report Suites*.
- **Segmente mit `--rsids`** liefert nur Segmente, die auf diesen Report Suites *angelegt*
  wurden. Segmente sind aber RS-übergreifend nutzbar. Für eine vollständige Liste
  `--all-segments` verwenden.
- Bei sehr exotischen Segment-Konstrukten (z. B. Zeitbeschränkungen in Sequenzen) wird der
  betreffende Teil als JSON in die lesbare Definition übernommen. Das komplette Original
  steht immer in `definition_json`.
- Bei Rate-Limits (HTTP 429) wartet das Skript automatisch und versucht es erneut.
