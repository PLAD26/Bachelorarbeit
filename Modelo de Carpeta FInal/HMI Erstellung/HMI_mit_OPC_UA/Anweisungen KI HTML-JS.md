# Anweisungen für die KI: HTML-HMI mit Anbindung an CODESYS

## Ziel

Ein HMI in HTML entwerfen und die Datei `server.js` (Vorlage) so anpassen, dass das HTML Variablen einer CODESYS-SPS über OPC UA lesen und schreiben kann.

Dazu erhältst du **zwei Listen**:

- **Liste 1 – R&I-Elemente:** die Elemente, die im R&I vorkommen, jeweils mit ihrer SPS-Variable.
- **Liste 2 – Vorgeschlagene Elemente:** die zusätzlichen HMI-Elemente (Schaltflächen, Fenster, Bildschirmseiten, Anzeigen usw.), jeweils mit Name, einer Beschreibung ihrer Funktion und ihrer SPS-Variable.

Lieferumfang:
1. `public/<name>.html` – das entworfene HMI.
2. `server.js` – die Vorlage, in der **nur** der Abschnitt `ZONA MODIFICABLE` angepasst ist.

## Architektur (nicht ändern)

```
Browser (HTML)  <--Socket.IO-->  Node.js (server.js)  <--OPC UA-->  CODESYS
```

Das HTML kommuniziert nie direkt mit der SPS. Es verwendet ausschließlich die Tag-Namen, die in der Tabelle `TAGS` in `server.js` definiert sind.

## Listen, die du erhältst

### Liste 1 – R&I-Elemente

Ein Element pro Zeile. Format: `Name im HMI: Pfad.in.der.SPS`

```
Tank A: Simulation.Level_A
NS1: GVL.N1.OUT
```

Diese Elemente **bilden den Anlagenzustand ab** (Motoren ein oder aus, Füllstände, Sensoren usw.). Zeichne sie im HMI als grafische Elemente, die den Wert ihrer Variable **anzeigen**. Weise ihnen keine Schreibaktionen zu, es sei denn, die Beschreibung eines Elements aus Liste 2 sieht dies vor.

### Liste 2 – Vorgeschlagene Elemente

Jedes Element hat einen Namen, eine Beschreibung und eine Variable:

```
Boton_Manual_NS1
  - Descripción: Botón que activa el modo manual del motor NS1 desde el HMI.
  - Variable: GVL.N1.MANUAL
```

- Die **Beschreibung** gibt an, was das Element tun soll und wie es sich im HMI verhalten soll (zum Beispiel momentan wirkender Taster, Schalter mit Einrastfunktion usw.). Nutze sie, um Gestaltung und Logik des Elements festzulegen.
- Die **Variable** ist diejenige, die in `TAGS` mit dem Element verknüpft wird.
- Wenn die Beschreibung nicht klärt, wie sich eine Schaltfläche verhalten soll (momentan wirkend oder mit Einrastfunktion), wähle die logischste Variante und halte sie in den Entwurfsannahmen fest.

### Umsetzung in die Tabelle `TAGS`

- Der Text vor dem Doppelpunkt (Liste 1) bzw. der Name des Elements (Liste 2) wird zu `name`, ohne Leerzeichen (`Tank A` → `Tank_A`).
- Der Variablenpfad wird unverändert zu `plc`.
- Für **jedes Element beider Listen** wird **eine `TAGS`-Zeile** erzeugt. Jeder `name` muss eindeutig sein. Zwei verschiedene Elemente dürfen denselben `plc`-Pfad verwenden (zum Beispiel eine Anzeige aus Liste 1 und eine Schaltfläche aus Liste 2, die auf dieselbe Variable zeigen), aber jedes mit eigenem `name`.
- Wenn derselbe Name in beiden Listen vorkommt, ändere einen davon und weise den Benutzer darauf hin.
- Sonst nichts: Typ und Zugriff erkennt der Server selbst.

## Regeln für `server.js`

1. Ändere **nur** den Bereich zwischen `>>> ZONA MODIFICABLE` und `<<< ZONA MODIFICABLE`.
2. Ändere nichts zwischen `>>> NO TOCAR` und `<<< NO TOCAR`.
3. Innerhalb des änderbaren Bereichs passt du ausschließlich an:
   - die Tabelle `TAGS` (eine Zeile pro Element der beiden Listen),
   - `NS`, `DEVICE`, `APPLICATION`, wenn der Benutzer sie angibt,
   - `ENDPOINT`, `USER`, `PASS`, `USE_LOGIN`, `HTTP_PORT`, wenn der Benutzer sie angibt.
4. Wenn der Benutzer keinen Verbindungswert angibt, belasse den Wert aus der Vorlage.

### Format jeder `TAGS`-Zeile

```js
{ name: "NOMBRE_HTML", plc: "RUTA.EN.EL.PLC" },
```

| Feld | Bedeutung |
|---|---|
| `name` | Im HTML verwendeter Name. Eindeutig, ohne Leerzeichen. |
| `plc` | Pfad der Variable in CODESYS, ohne das Geräte-Präfix. Z. B.: `GVL.N1.OUT` |
| `readOnly` | (Optional) `true`, wenn das HMI diesen Tag nie beschreiben darf. |

Empfehlung: Setze `readOnly: true` bei den Elementen aus Liste 1 (zeigen nur den Zustand an) und bei den Elementen aus Liste 2, die laut Beschreibung nur lesend sind.

**Gib den Datentyp und den Zugriff weder an noch nimm sie an.** Der Server erkennt sie beim Start, indem er die SPS ausliest (Typ: BOOL, REAL, INT usw.; Lese-/Schreibberechtigung), und zeigt sie in einer Tabelle in der Konsole an.

## Vertrag des HTML

Immer einbinden:

```html
<script src="/socket.io/socket.io.js"></script>
<script>
  const socket = io();

  // Werte von der SPS EMPFANGEN (einmal pro Tag beim Verbinden und bei jeder Änderung)
  socket.on("value", ({ name, value }) => {
    // das grafische Element aktualisieren, das zu "name" gehört
  });

  // (Optional) erkannte Informationen zu jedem Tag: [{ name, type, access }]
  // type: Boolean, Float, Int16, ...   access: read, write, readwrite
  socket.on("tags", (tags) => { /* ... */ });

  // Einen Wert an die SPS SENDEN
  // socket.emit("set", { name: "NOMBRE_TAG", value: true });
</script>
```

Regeln:
- Verwende **exakt** die Namen aus der Spalte `name` von `TAGS`. Erfinde keine Tags.
- Alle lesbaren Tags kommen über dasselbe Ereignis `"value"`; unterscheide sie anhand von `name`.
- Nimm den Typ nicht an: Prüfe beim Anzeigen des Wertes mit `typeof value` (`"boolean"` oder `"number"`). Eine Zahl wird mit `.toFixed(1)` angezeigt, wenn sie dezimal ist.
- Zum Schreiben: Sende `true`/`false` an einen booleschen Tag und eine Zahl an einen numerischen; der Server konvertiert in den exakten SPS-Typ.
- Schreibe nur in die Tags der Elemente aus **Liste 2**, deren Beschreibung angibt, dass das HMI sie steuern soll (Schaltflächen, Sollwerte). Die Elemente aus Liste 1 werden nur gelesen.
- Wenn der Server einen Schreibvorgang ablehnt, meldet er dies in der Konsole (Tag ohne Berechtigung in der SPS oder als `readOnly` markiert).
- Momentan wirkender Taster: `set true` bei `mousedown` und `set false` bei `mouseup` und `mouseleave`. Für Touchscreens ergänze `touchstart` und `touchend`.
- Zeige den Verbindungsstatus mit den Socket.IO-Ereignissen `connect` und `disconnect` an.
- Alle Elemente aus Liste 2 müssen im HMI mit dem Verhalten erscheinen, das ihre Beschreibung vorgibt.
- Das HTML muss eine einzelne Datei mit eingebettetem CSS und JavaScript sein, ohne externe Bibliotheken, es sei denn, der Benutzer verlangt etwas anderes.

## Format der Antwort der KI

1. Das vollständige HTML.
2. Der geänderte Code.
3. Die nötigen Hinweise an den Benutzer.
