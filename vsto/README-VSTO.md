# VSTO-Version: Antworten + Archivieren (klassisches Outlook, Windows)

Klassisches COM-/VSTO-Add-in für **Outlook Desktop (Windows)**. Fügt in einer gelesenen E-Mail eine Schaltfläche **„Antworten + Archivieren“** hinzu: öffnet das Antwortformular und verschiebt nach dem Senden die Original-E-Mail in den Ordner **„Archiv“** (Fallback „Archive“).

## Voraussetzungen

- **Visual Studio 2022** mit Workload **„Office/SharePoint-Entwicklung“** (Installer → Workloads → „Office/SharePoint-Entwicklung“, enthält die VSTO-Projektvorlagen)
- **.NET Framework 4.6 – 4.8** (in Windows enthalten)
- Klassisches **Outlook für Windows** (nicht „neues Outlook“ – das unterstützt keine COM-/VSTO-Add-ins)
- Im Postfach ein Ordner namens **„Archiv“** (oder Name in `ThisAddIn.cs` oben anpassen: `ARCHIVE_FOLDER_NAME`)

## Projekt anlegen (einmalig)

Die Designerdateien eines VSTO-Projekts erzeugt Visual Studio automatisch – daher legt man das Projekt aus der Vorlage an und fügt die Quellcodedateien aus diesem Ordner ein:

1. Visual Studio → **Neues Projekt** → Vorlage **„Outlook-VSTO-Add-In“** (C#) → Projektname **`OutlookArchiveOnReply`** (wichtig: exakt dieser Name, wegen des Namensraums).
2. Im Projektmappen-Explorer:
   - Inhalt von **`ThisAddIn.cs`** (aus diesem Ordner) in die automatisch erstellte `ThisAddIn.cs` kopieren (alles ersetzen).
   - Rechtsklick auf Projekt → **Hinzufügen → Neues Element → „Klasse“** → Name `Ribbon1.cs` → Inhalt aus **`vsto/Ribbon1.cs`** einfügen (alles ersetzen).
   - `Ribbon1.xml` wird nicht benötigt – das Ribbon-XML ist als Konstante direkt in `Ribbon1.cs` enthalten (robuster, keine Probleme mit eingebetteten Ressourcen).
3. Sicherstellen, dass `Ribbon1.cs` im Projektmappen-Explorer unter dem Projekt liegt, **nicht ausgegraut** ist (falls doch: Rechtsklick → „In Projekt einschließen“) und als Buildvorgang **„Kompilieren“** hat (Datei anklicken → F4).

## Bauen und installieren

1. **`F5`** drücken → Visual Studio startet Outlook mit dem Add-in (Debug-Modus). Alternativ: **Build → Projektmappe erstellen** erzeugt im Ordner `bin\Debug\` die **`.dll`** plus `.vsto`-Manifest und ClickOnce-Setup.
2. Dauerhaft installieren (für Debug-Build): `bin\Debug\` → `OutlookArchiveOnReply.vsto` doppelklicken → ClickOnce-Installation. Für Verteilung: Projekt-Eigenschaften → **Veröffentlichen** → Setup (`setup.exe`) erzeugen.
3. Nach dem Start von Outlook: E-Mail öffnen → im Menüband der Lesemail erscheint die Gruppe **„Archivierung“** mit **„Antworten + Archivieren“**.

## Funktionsweise

- `Ribbon1.cs` liefert das Ribbon-XML inline und fügt den Button in den Standard-Tab **„TabReadMessage“** (Lesemail) ein.
- Klick → `Ribbon1.OnArchiveReply` → `ThisAddIn.StartArchiveReply`: erstellt per `MailItem.Reply()` die Antwort, merkt sich das Zuordnungspaar (Antwort → Original) und zeigt das Antwortfenster.
- Schließt der Nutzer die Antwort **ohne Senden**, wird die Zuordnung beim `Close`-Event wieder entfernt.
- `Application.ItemSend`: Wird die gemerkte Antwort gesendet, wird das Original per `MailItem.Move()` in den Archiv-Ordner verschoben (sucht „Archiv“/„Archive“ im gesamten Postfach, rekursiv).
- Andere, normale gesendete Mails werden nicht verändert.

## Fehlerbehebung (bekannt aus der Praxis)

- **„Typ- oder Namespacename ‚Ribbon1‘ wurde nicht gefunden“**: `Ribbon1.cs` liegt nicht im Projekt (im Projektmappen-Explorer prüfen; ausgraue Datei → „In Projekt einschließen“) oder hat nicht den Buildvorgang „Kompilieren“. Notlösung: Klasse in `ThisAddIn.cs` verschieben.
- **Outlook startet bei F5 nicht**: alter OUTLOOK.EXE-Prozess läuft noch → im Task-Manager beenden, dann `Umschalt+F5` und erneut F5.
- **Button erscheint nicht**: Ribbon-XML wurde nicht geladen. Diagnose: Haltepunkt in `GetCustomUI` – muss beim Öffnen einer Mail mit `ribbonID == "Microsoft.Outlook.Mail.Read"` aufgerufen werden.
- **Button ohne Icon**: `imageMso`-Name je nach Outlook-Version nicht verfügbar – `MailReplyAll` ist breit verfügbar, oder `imageMso` weglassen.

## Anpassungen

- **Zielordner**: `ARCHIVE_FOLDER_NAME` / `ARCHIVE_FOLDER_NAME_EN` in `ThisAddIn.cs`.
- **Button-Icon**: `imageMso` in der XML-Konstante in `Ribbon1.cs` (z. B. `MailReplyAll`, `ArchiveFolder`). Je nach Outlook-Version sind nicht alle Namen verfügbar; `imageMso` kann auch ganz weggelassen werden.

## Hinweise

- COM-/VSTO-Add-ins laufen **nur im klassischen Outlook für Windows** (ab Outlook 2013). Im „neuen Outlook“ und unter macOS werden sie nicht geladen – dort siehe das Web-Add-in im Hauptverzeichnis dieses Repos.
- Beim ersten Start fragt ClickOnce nach Vertrauen → mit **„Installieren“** bestätigen.
- Deinstallation: Outlook schließen → Systemsteuerung → Programme → „OutlookArchiveOnReply“ entfernen.
