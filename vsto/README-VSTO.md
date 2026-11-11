# VSTO-Version: Antworten + Archivieren (klassisches Outlook, Windows)

Klassisches COM-/VSTO-Add-in für **Outlook Desktop (Windows)**. Verschiebt die beantwortete Original-E-Mail nach dem Senden automatisch in den Ordner **„Archiv“** (Fallback „Archive“).

## Features

| Ort | Element | Funktion |
|---|---|---|
| **Hauptfenster** (Mail-Liste) | Button „Antworten + Archivieren“ direkt neben der Antworten-Gruppe | Erstellt Antwort auf die markierte Mail; nach dem Senden wird die Original-Mail archiviert |
| **Lesefenster** (geöffnete Mail) | Button „Antworten + Archivieren“ direkt neben der Antworten-Gruppe | Gleicher Ablauf für die geöffnete Mail |
| **Verfassen-Fenster** (Antwort/Neue Mail) | Eigener Tab „SENDEN & ARCHIVIEREN“ mit Checkbox „Original nach Senden archivieren“ | Checkbox anhaken = Original-Mail wird beim Senden archiviert; bei Antworten über die Buttons automatisch angehakt; auch bei normal geöffneten Antworten nutzbar (findet die Original-Mail über den Konversationsindex im Posteingang, max. ~300 neueste Mails) |

Kontextmenü-Einträge (Rechtsklick) werden von aktuellen Microsoft-365-Builds (RibbonX-`contextMenus`) nicht mehr übernommen – der Button im Hauptfenster ist der Ersatz.

## Voraussetzungen

- **Visual Studio 2022** mit Workload **„Office/SharePoint-Entwicklung“**
- **.NET Framework 4.6 – 4.8** (in Windows enthalten)
- Klassisches **Outlook für Windows** (getestet mit Microsoft 365, Build 2610; nicht „neues Outlook“ – das unterstützt keine COM-/VSTO-Add-ins)
- Im Postfach ein Ordner namens **„Archiv“** (oder Name in `ThisAddIn.cs` anpassen: `ARCHIVE_FOLDER_NAME`)

## Projekt anlegen (einmalig)

Die Designerdateien eines VSTO-Projekts erzeugt Visual Studio automatisch – daher legt man das Projekt aus der Vorlage an und fügt die Quellcodedatei ein:

1. Visual Studio → **Neues Projekt** → Vorlage **„Outlook-VSTO-Add-In“** (C#) → Projektname **`OutlookArchiveOnReply`** (wichtig: exakt dieser Name, wegen des Namensraums).
2. Inhalt von **`vsto/ThisAddIn.cs`** komplett in die automatisch erstellte `ThisAddIn.cs` kopieren (alles ersetzen). Die Datei enthält **alles**: die Add-in-Logik (Klasse `ThisAddIn`) und die Ribbon-Klasse (`Ribbon1`) inklusive der Ribbon-XML-Konstanten als Inline-Strings. Eine separate `Ribbon1.cs`/`Ribbon1.xml` wird **nicht** benötigt.
3. Falls der Build `System.Windows.Forms` vermisst: Verweise → Verweis hinzufügen → **System.Windows.Forms** anhaken (wird für den Hinweis-Dialog der Compose-Checkbox benötigt).
4. `Strg+Umschalt+B` → 0 Fehler → **F5** startet Outlook mit dem Add-in (Debug-Modus).

## Bauen und installieren

1. **`F5`** → Outlook startet mit dem Add-in. Alternativ: **Build → Projektmappe erstellen** erzeugt in `bin\Debug\` die **`.dll`** plus `.vsto`-Manifest.
2. Dauerhaft installieren: `bin\Debug\` → `OutlookArchiveOnReply.vsto` doppelklicken → ClickOnce-Installation bestätigen. Für Verteilung: Projekt-Eigenschaften → **Veröffentlichen** → `setup.exe` erzeugen.
3. Deinstallation: Outlook schließen → Systemsteuerung → Programme → „OutlookArchiveOnReply“ entfernen.

## Funktionsweise

- **Buttons** (Hauptfenster/Lesefenster) → `OnArchiveReply` → `StartArchiveReply`: erstellt per `MailItem.Reply()` die Antwort, merkt sich das Zuordnungspaar (Antwort → Original) und zeigt das Antwortfenster.
- **Checkbox** (Verfassen-Fenster) → `SetArchivePending`: hakt die Archivierung für die aktuelle Antwort an (oder entfernt sie); `FindOriginalForReply` lokalisiert die Original-Mail über den Konversationsindex (Original-Index = Antwort-Index minus 5 Bytes, hexadezimal) im Posteingang.
- Schließen der Antwort **ohne Senden** entfernt die Zuordnung (`Close`-Event).
- `Application.ItemSend`: Wird eine gemerkte Antwort gesendet, wird das Original per `MailItem.Move()` in den Archiv-Ordner verschoben (rekursive Suche nach „Archiv“/„Archive“). Andere gesendete Mails bleiben unberührt.

## Wichtige Erkenntnisse aus der Praxis (Fehlerbehebung)

- **Ribbon-IDs müssen in Outlook über ALLE Ribbons hinweg eindeutig sein** (Explorer, Read, Compose). Doppelte IDs (z. B. zwei Gruppen mit `id="ArchiveGroup"`) führen dazu, dass das zweite Ribbon still fehlschlägt – ohne Fehlermeldung. Deshalb: `archiveExplorerGroup`, `archiveReadGroup`, `archiveComposeGroup` usw.
- **`GetCustomUI` wird je ribbonID aufgerufen**: `"Microsoft.Outlook.Explorer"` (Hauptfenster), `"Microsoft.Outlook.Mail.Read"` (Lesefenster), `"Microsoft.Outlook.Mail.Compose"` (Verfassen-Fenster). Diagnose: Haltepunkt auf dem `switch`.
- **Ribbon-XML wird nur beim Outlook-Start geladen** – nach jeder XML-Änderung Debuggen beenden und F5 (alter OUTLOOK.EXE-Prozess vorher im Task-Manager beenden, sonst startet Outlook nicht).
- **Aktuelle Microsoft-365-Builds übernehmen keine RibbonX-Kontextmenüs** (`<contextMenus>` wird ignoriert) – stattdessen Button im Explorer-Ribbon verwenden.
- **Vereinfachtes Menüband (Simplified Ribbon)** kann eigene Gruppen auf eingebauten Tabs verbergen; eigene Tabs erscheinen im ⋯-Overflow. Umschalter unten rechts am Menüband.
- **„Ribbon1 nicht gefunden“** beim Build: Datei liegt nicht im Projekt (ausgegraute Datei → „In Projekt einschließen“) oder Buildvorgang nicht „Kompilieren“.
- **Icons**: `imageMso`-Namen sind versionabhängig; `MailReplyAll` ist breit verfügbar. Ungültige Namen zeigen nur kein Icon, brechen das Ribbon nicht.

## Anpassungen

- **Zielordner**: `ARCHIVE_FOLDER_NAME` / `ARCHIVE_FOLDER_NAME_EN` in `ThisAddIn.cs`.
- **Tab-/Gruppen-/Button-Texte und Positionen**: die drei XML-Konstanten `EXPLORER_XML`, `READ_XML`, `COMPOSE_XML` in `Ribbon1`.
- **Suchtiefe für Original-Mail**: `MAX_SCAN_ITEMS` in `ThisAddIn.cs`.

## Hinweise

- COM-/VSTO-Add-ins laufen **nur im klassischen Outlook für Windows**. Im „neuen Outlook“ und unter macOS werden sie nicht geladen – dort siehe das Web-Add-in im Hauptverzeichnis dieses Repos.
- Beim ersten Start fragt ClickOnce nach Vertrauen → mit **„Installieren“** bestätigen.
