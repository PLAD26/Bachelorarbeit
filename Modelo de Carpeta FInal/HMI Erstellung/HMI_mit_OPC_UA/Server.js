// =====================================================================
// server.js – VORLAGE  |  HMI (HTML) <-> CODESYS (OPC UA) mit Node.js
// =====================================================================
// REGEL FÜR DIE KI:
//   - Ändere NUR den Abschnitt "ÄNDERBARER BEREICH" (Konstanten und TAGS-Tabelle).
//   - Ändere NICHTS im Abschnitt "NICHT ÄNDERN".
//   - Datentyp und Schreibberechtigung jeder Variable ERKENNT der Server
//     beim Start selbst, indem er die SPS ausliest. Sie müssen weder
//     angegeben noch angenommen werden.
//   - Das HTML kommuniziert mit diesem Server über Socket.IO:
//       Server -> HTML : "value" { name, value }          (Wert eines Tags)
//       Server -> HTML : "tags"  [{ name, type, access }]  (erkannte Infos, optional)
//       HTML -> Server : "set"   { name, value }          (ein Tag schreiben)
//     "name" ist immer der in der TAGS-Tabelle definierte Tag-Name.
// =====================================================================

const express = require("express");
const http = require("http");
const { Server } = require("socket.io");
const {
  OPCUAClient, AttributeIds, DataType, TimestampsToReturn,
  MessageSecurityMode, SecurityPolicy, UserTokenType,
} = require("node-opcua");

// =====================================================================
// >>> ÄNDERBARER BEREICH  (Anfang)
// =====================================================================

// --- Verbindung zur SPS ---
const ENDPOINT = "opc.tcp://127.0.0.1:4840";   // IP der SPS
const USER = "Ali";                      
const PASS = "Versuch123.";                   
const USE_LOGIN = true;                        // false = anonymer Zugriff

// --- Aufbau der NodeId (mit "Copy NodeId" in UaExpert prüfen) ---
const NS = 4;                                  // Namespace
const DEVICE = "CODESYS Control Win V3 x64";   // Name des Geräts/der Runtime
const APPLICATION = "Application";             // Name der Applikation

// --- HMI-Port ---
const HTTP_PORT = 3000;

// --- TAG-TABELLE: eine Zeile pro SPS-Variable ---
//   name     : Name, den das HTML verwendet (eindeutig, ohne Leerzeichen)
//   plc      : Pfad der Variable in der SPS, so wie er in CODESYS erscheint
//   readOnly : (optional) true = das HMI kann diesen Tag nie schreiben,
//              auch wenn die SPS es erlauben würde
// Der Typ (BOOL, REAL, INT...) und die Berechtigungen werden automatisch erkannt.
const TAGS = [
  { name: "Tank_A", plc: "Simulation.Level_A" },
  { name: "NS1",    plc: "GVL.N1.OUT" },
];

// =====================================================================
// <<< ÄNDERBARER BEREICH  (Ende)
// =====================================================================


// =====================================================================
// >>> NICHT ÄNDERN  (Anfang)
// =====================================================================

const SUPPORTED = ["Boolean", "SByte", "Byte", "Int16", "UInt16", "Int32", "UInt32", "Float", "Double", "String"];
const READ_BIT = 1;    // OPC UA AccessLevel: CurrentRead
const WRITE_BIT = 2;   // OPC UA AccessLevel: CurrentWrite

const nodeIdOf = (tag) => `ns=${NS};s=|var|${DEVICE}.${APPLICATION}.${tag.plc}`;

// Wandelt den vom HTML empfangenen Wert in den in der SPS erkannten Typ um
function convert(tag, value) {
  switch (tag.type) {
    case "Boolean": return value === true || value === 1 || value === "true" || value === "1";
    case "Float":
    case "Double":  return Number(value);
    case "String":  return String(value);
    default:        return Math.trunc(Number(value));
  }
}

const app = express();
const server = http.createServer(app);
const io = new Server(server);
app.use(express.static("public"));

(async () => {
  try {
    // Validierung der TAGS-Tabelle
    const names = new Set();
    for (const t of TAGS) {
      if (!t.name || !t.plc) throw new Error("Jeder Tag benötigt 'name' und 'plc'");
      if (names.has(t.name)) throw new Error(`Doppelter Tag: ${t.name}`);
      names.add(t.name);
    }

    const client = OPCUAClient.create({
      applicationName: "HMI-Node",
      endpointMustExist: false,
      securityMode: MessageSecurityMode.None,
      securityPolicy: SecurityPolicy.None,
      connectionStrategy: { initialDelay: 1000, maxRetry: 3 },
    });

    await client.connect(ENDPOINT);
    console.log("OPC UA: verbunden mit", ENDPOINT);

    const session = USE_LOGIN
      ? await client.createSession({ type: UserTokenType.UserName, userName: USER, password: PASS })
      : await client.createSession();
    console.log("OPC UA: Sitzung erstellt", USE_LOGIN ? `(Benutzer ${USER})` : "(anonym)");

    // ---- Automatische Erkennung von Typ und Berechtigungen ----
    const nodesToRead = [];
    for (const t of TAGS) {
      nodesToRead.push({ nodeId: nodeIdOf(t), attributeId: AttributeIds.DataType });
      nodesToRead.push({ nodeId: nodeIdOf(t), attributeId: AttributeIds.UserAccessLevel });
    }
    const results = await session.read(nodesToRead);

    const active = [];   // gültige Tags
    const byName = {};
    TAGS.forEach((t, i) => {
      const dtDv = results[2 * i];
      const alDv = results[2 * i + 1];

      if (dtDv.statusCode.isNotGood() || !dtDv.value || !dtDv.value.value) {
        console.error(`✘ ${t.name}: existiert nicht in der SPS (${dtDv.statusCode.toString()}) -> ${nodeIdOf(t)}`);
        return;
      }
      const dtId = dtDv.value.value;                       // NodeId des Datentyps
      const typeName = dtId.namespace === 0 ? DataType[dtId.value] : undefined;
      if (!SUPPORTED.includes(typeName)) {
        console.error(`✘ ${t.name}: Typ nicht unterstützt (${typeName || dtId.toString()}); wird ignoriert`);
        return;
      }

      let level = READ_BIT;                                // standardmäßig nur Lesen
      if (alDv.statusCode.isGood() && alDv.value && typeof alDv.value.value === "number") {
        level = alDv.value.value;
      } else {
        console.warn(`! ${t.name}: Zugriffsberechtigung konnte nicht gelesen werden; wird als nur lesbar behandelt`);
      }

      t.type = typeName;
      t.canRead = !!(level & READ_BIT);
      t.canWrite = !!(level & WRITE_BIT) && !t.readOnly;
      active.push(t);
      byName[t.name] = t;
    });

    console.log("\nErkannte Tags:");
    console.table(active.map((t) => ({
      name: t.name,
      plc: t.plc,
      typ: t.type,
      zugriff: t.canRead && t.canWrite ? "Lesen/Schreiben" : t.canWrite ? "Schreiben" : "Lesen",
    })));

    const tagInfo = active.map((t) => ({
      name: t.name,
      type: t.type,
      access: t.canRead && t.canWrite ? "readwrite" : t.canWrite ? "write" : "read",
    }));

    // ---- Abonnement aller lesbaren Tags ----
    const lastValues = {};
    const subscription = await session.createSubscription2({
      requestedPublishingInterval: 200,
      requestedLifetimeCount: 100,
      requestedMaxKeepAliveCount: 10,
      maxNotificationsPerPublish: 100,
      publishingEnabled: true,
      priority: 10,
    });

    for (const tag of active.filter((t) => t.canRead)) {
      try {
        const mon = await subscription.monitor(
          { nodeId: nodeIdOf(tag), attributeId: AttributeIds.Value },
          { samplingInterval: 200, discardOldest: true, queueSize: 1 },
          TimestampsToReturn.Both
        );
        let warned = false;
        mon.on("changed", (dv) => {
          if (dv.statusCode.isNotGood() || !dv.value || dv.value.value === null) {
            if (!warned) {
              console.warn(`Tag "${tag.name}" ohne gültigen Wert (${dv.statusCode.toString()})`);
              warned = true;
            }
            return;
          }
          lastValues[tag.name] = dv.value.value;
          io.emit("value", { name: tag.name, value: dv.value.value });
        });
      } catch (err) {
        console.error(`"${tag.name}" konnte nicht überwacht werden:`, err.message);
      }
    }

    // ---- Kommunikation mit dem Browser ----
    io.on("connection", (socket) => {
      console.log("Browser verbunden:", socket.id);

      socket.emit("tags", tagInfo);
      for (const [name, value] of Object.entries(lastValues)) {
        socket.emit("value", { name, value });
      }

      socket.on("set", async ({ name, value }) => {
        const tag = byName[name];
        if (!tag) {
          console.warn(`Schreibzugriff abgelehnt: Tag "${name}" existiert nicht oder ist ungültig`);
          return;
        }
        if (!tag.canWrite) {
          console.warn(`Schreibzugriff abgelehnt: "${name}" hat keine Schreibberechtigung`);
          return;
        }
        try {
          const status = await session.write({
            nodeId: nodeIdOf(tag),
            attributeId: AttributeIds.Value,
            value: { value: { dataType: DataType[tag.type], value: convert(tag, value) } },
          });
          if (status.isNotGood()) console.error(`Fehler beim Schreiben von ${name}:`, status.toString());
        } catch (err) {
          console.error(`Ausnahme beim Schreiben von ${name}:`, err.message);
        }
      });

      socket.on("disconnect", () => console.log("Browser getrennt:", socket.id));
    });

    process.on("SIGINT", async () => {
      console.log("\nOPC-UA-Verbindung wird geschlossen...");
      try {
        await subscription.terminate();
        await session.close();
        await client.disconnect();
      } finally {
        process.exit(0);
      }
    });

    server.listen(HTTP_PORT, () => console.log(`HMI: http://localhost:${HTTP_PORT}`));
  } catch (err) {
    console.error("Fehler beim Start:", err.message);
    process.exit(1);
  }
})();

// =====================================================================
// <<< NICHT ÄNDERN  (Ende)
// =====================================================================
