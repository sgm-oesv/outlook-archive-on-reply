# Antworten & Archivieren (Outlook-Add-in)

Outlook-Add-in, das das Beantworten von E-Mails um eine Archiv-Funktion erweitert:

1. In einer **gelesenen E-Mail** auf **„Antworten & Archivieren“** klicken.
2. Das normale Antwortformular öffnet sich.
3. Beim **Senden der Antwort** wird die ursprüngliche E-Mail automatisch in den Ordner **„Archiv“** (Fallback: „Archive“) verschoben.

## Projektstruktur

```
outlook-archive-on-reply/
├── manifest.xml              # Add-in-Manifest (XML)
├── src/
│   ├── function-file.html    # Lädt Office.js + commands.js
│   └── commands.js            # Ribbon-Button + OnMessageSend-Handler (EWS)
├── scripts/
│   └── generate-icons.py      # Erzeugt die Platzhalter-Icons
└── assets/                   # Icons (16/32/64/80 px)
```

Falls `assets/` fehlt: `python3 scripts/generate-icons.py`

## Funktionsweise

- **Ribbon-Button** (`MessageReadCommandSurface`): Speichert `itemId` und `conversationId` der Original-Mail über `Office.addin.setStartupVariables` und öffnet per `displayReplyForm` das Antwortformular.
- **`OnMessageSend`-Event** (`LaunchEvent`, SendMode `SoftBlock`): Wird beim Senden ausgelöst. Stimmt die Conversation-ID mit der gemerkten Mail überein, wird die Original-Mail per EWS (`FindFolder` + `MoveItem`) in den Archiv-Ordner verschoben. Danach wird das Senden erlaubt.
- Der Handler läuft asynchron; das Senden wird erst nach dem Verschieben freigegeben (max. einige Sekunden, danach Timeout durch Outlook).

## Voraussetzungen

- Outlook (Windows/Mac, klassisch) oder Outlook im Web – `makeEwsRequestAsync` ist in **neuem Outlook unter Windows evtl. nicht verfügbar**; dort müsste stattdessen Microsoft Graph (`/messages/{id}/move`) verwendet werden.
- Ordner **„Archiv“** im Postfach (oder Name in `ARCHIVE_FOLDER_NAMES` in `src/commands.js` anpassen).
- Berechtigung `ReadWriteMailbox` ist im Manifest gesetzt.

## Lokal testen (Sideload)

1. HTTPS-Server mit gültigem Zertifikat starten, z. B.:

   ```bash
   npx office-addin-dev-certs install
   npx office-addin-dev-server "https://localhost:3000"
   ```

2. Manifest in Outlook Sideloaden:
   - **Windows (klassisch)**: Datei → Info → Add-Ins verwalten → „Meine Add-Ins“ → „Add-Ins hochladen“ → `manifest.xml` auswählen.
   - **Outlook im Web**: Einstellungen → Add-Ins → „Meine Add-Ins“ → „Benutzerdefiniertes Add-In hinzufügen“ → `manifest.xml`.

3. E-Mail öffnen → Button **„Antworten & Archivieren“** → Antwort senden → Original-Mail liegt im Archiv-Ordner.

## Anpassungen

- **Zielordner**: `ARCHIVE_FOLDER_NAMES` in `src/commands.js`.
- **Statt EWS Graph verwenden**: In `moveToArchiveFolder` einen Graph-Call `POST /me/messages/{id}/move` einbauen (erfordert SSO/`getCallbackTokenAsync` mit REST-Id via `Office.context.mailbox.convertToRestId`).
