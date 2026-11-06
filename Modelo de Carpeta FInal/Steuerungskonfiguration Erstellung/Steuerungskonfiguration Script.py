# -*- coding: utf-8 -*-
"""
Programm : generador_steuerungskonfig.py
Zweck    : Liest die JSON-Datei (Steuerungskonfiguration & E/A-Belegung) und die GSDML-/XML-Dateien
           aus dem Ordner "GSD" und erzeugt daraus ein CODESYS-Skript
           (steuerungskonfig_generator.py), das die Steuerungskonfiguration und die
           E/A-Belegung direkt in das gerade GEÖFFNETE CODESYS-Projekt einfügt
           (es wird kein neues Projekt angelegt). Zusätzlich wird das ST-Programm PLC_PRG
           erzeugt, das die realen Ein-/Ausgänge mit den Funktionsbausteinen verbindet
           (Sensoren: Eingang -> Baustein-Pin, Aktoren: Baustein-Pin -> Ausgang) und über
           die Variable "Anlage" zwischen Prozess und Simulation umschaltet.
Start    : normales Python 3 (nicht CODESYS!)
               python generador_steuerungskonfig.py          -> öffnet ein Fenster (Standard)
               python generador_steuerungskonfig.py MeineDatei.json [-o skript.py] [-d GSD]
                                                              -> ohne Fenster (Kommandozeile)
Benötigt : nur die Python-Standardbibliothek (tkinter ist bei Python für Windows dabei)
Ordner   : Dieses Programm, der Ordner "GSD" und (ohne Angabe) die JSON-Datei
           liegen im selben Ordner.

JSON (Aufbau):
  "plc"           : {"type": Name der SPS im Projekt,
                     "instance": Gerätename im Geräte-Repository (optional)}
  "ethernet"      : {"type": Typ Ethernet,   "instance": gewünschter Name}
  "pn_controller" : {"type": Typ Controller, "instance": gewünschter Name}
  "devices"       : Liste von Geräten
        {"instance": gewünschter Name, "dap": technischer Name,
         "modules": [{"instance": gewünschter Name, "module": technischer Name}, ...]}
        (Module gehören zum jeweiligen Gerät)
  "io"            : Liste von E/A-Zuweisungen
        {"device_or_submodule": Modul/Gerät, "exact_channel": Kanal, "variable": Variablenname}
  "addresses"     : (für das ST-Programm PLC_PRG)
        "Sensors"   : Liste {"variable": "L100.IN",  "address": "%IX0.0"}  ->  L100.IN := %IX0.0;
        "Actuators" : Liste {"variable": "Y100.OUT", "address": "%QX0.0"}  ->  %QX0.0 := Y100.OUT;
  Weitere Abschnitte werden ignoriert.
"""
from __future__ import print_function
import argparse
import contextlib
import json
import os
import re
import sys
import xml.etree.ElementTree as ET

# ======================================================================
# FESTE KONSTANTEN (ändern sich nie)
# ======================================================================
STEUERUNG_SUCHTEXT = "CODESYS Control Win V3"   # Standard, wenn "plc.instance" im JSON fehlt
STEUERUNG_VERSION = "3.5.21.20"          # CODESYS V3.5 SP21 Patch 2
GSD_NEU_INSTALLIEREN = True              # Standardwert: GSDML immer (neu) installieren (im Fenster umschaltbar)
ORDNER_NAME = "GSD"             # Ordner mit den Gerätebeschreibungen (XML)
STANDARD_AUSGABE = "steuerungskonfig_generator.py"
PROGRAMM_NAME = "PLC_PRG"               # Name des erzeugten ST-Programms
SCHALTER_NAME = "Anlage"                # BOOL-Variable: True = reale E/A, False = Simulation
SIMULATION_AUFRUF = "Simulation();"     # Aufruf im ELSE-Zweig der Sensor-Zuweisungen
ADRESSE_MUSTER = re.compile(r"^%[IQM][XBWDL]\d+(\.\d+)?$")   # z. B. %IX0.0, %QW1


# ======================================================================
# HILFSFUNKTIONEN
# ======================================================================
def norm(text):
    """Vergleichsform: ohne Leerzeichen, '_' und '-', kleingeschrieben."""
    return re.sub(r"[\s_\-]+", "", text).lower()


def zelle(wert):
    """Wert als Text; Leerzeichen am Anfang/Ende (auch geschützte) werden entfernt."""
    if wert is None:
        return ""
    if isinstance(wert, float) and wert.is_integer():
        wert = int(wert)
    return str(wert).replace("\xa0", " ").strip()


def lit(text):
    """Python-2/3-kompatibles Unicode-Literal (u"...") für das erzeugte Skript."""
    return "u" + json.dumps(text)


# ======================================================================
# JSON LESEN
# ======================================================================
def _abschnitt(daten, schluessel):
    """Liefert einen Abschnitt der JSON-Datei, der ein Objekt (dict) sein muss."""
    wert = daten.get(schluessel)
    if not isinstance(wert, dict):
        raise ValueError("Fehler: Abschnitt '%s' fehlt in der JSON-Datei oder ist kein Objekt." % schluessel)
    return wert


def _liste(daten, schluessel, pflicht=False):
    """Liefert eine Liste der JSON-Datei; jedes Element muss ein Objekt (dict) sein."""
    wert = daten.get(schluessel, [])
    if not isinstance(wert, list) or not all(isinstance(e, dict) for e in wert):
        raise ValueError("Fehler: '%s' muss in der JSON-Datei eine Liste von Objekten sein." % schluessel)
    if pflicht and not wert:
        raise ValueError("Fehler: '%s' enthält keinen Eintrag." % schluessel)
    return wert


def json_lesen(pfad):
    """Liest die JSON-Datei und liefert dieselben Strukturen wie früher die Excel-Datei:
    (sps, plc_suchtext, geraete, module, belegung, sensoren, aktoren)."""
    try:
        with open(pfad, "r", encoding="utf-8-sig") as f:
            daten = json.load(f)
    except json.JSONDecodeError as fehler:
        raise ValueError("Fehler: Ungültiges JSON (Zeile %d, Spalte %d): %s"
                         % (fehler.lineno, fehler.colno, fehler.msg))
    if not isinstance(daten, dict):
        raise ValueError("Fehler: Die JSON-Datei muss ein Objekt als Wurzel haben.")

    # Steuerung, Ethernet, Controller  (früher Blatt 1, Zeilen 1-3)
    plc = _abschnitt(daten, "plc")
    sps = zelle(plc.get("type"))                                  # Name der SPS im Projekt
    plc_suchtext = zelle(plc.get("instance")) or STEUERUNG_SUCHTEXT   # Gerät im Repository
    ethernet = _abschnitt(daten, "ethernet")
    ethernet_typ, ethernet_name = zelle(ethernet.get("type")), zelle(ethernet.get("instance"))
    controller = _abschnitt(daten, "pn_controller")
    controller_typ, controller_name = zelle(controller.get("type")), zelle(controller.get("instance"))
    for label, wert in (("plc/type (SPS-Name)", sps),
                        ("ethernet/type und instance", ethernet_typ and ethernet_name),
                        ("pn_controller/type und instance", controller_typ and controller_name)):
        if not wert:
            raise ValueError("Fehler: %s fehlt in der JSON-Datei." % label)

    # Geräte: (Name, Suchtext, Elterngerät)
    geraete = [
        (ethernet_name, ethernet_typ, sps),
        (controller_name, controller_typ, ethernet_name),
    ]
    # Module: (Name, technischer Modulname, Elterngerät, technischer Name des Elterngeräts)
    module = []
    for nr, dev in enumerate(_liste(daten, "devices", pflicht=True), start=1):
        name, dap = zelle(dev.get("instance")), zelle(dev.get("dap"))
        if not name or not dap:
            raise ValueError("Fehler: devices[%d]: 'instance' und 'dap' sind Pflichtfelder." % nr)
        geraete.append((name, dap, controller_name))
        for mnr, mod in enumerate(_liste(dev, "modules"), start=1):
            m_name, m_tech = zelle(mod.get("instance")), zelle(mod.get("module"))
            if not (m_name and m_tech):
                print("  Warnung: devices[%d].modules[%d] unvollständig, wird übersprungen." % (nr, mnr))
                continue
            module.append((m_name, m_tech, name, dap))

    # E/A-Belegung (früher Blatt 2)
    belegung = []
    for nr, eintrag in enumerate(_liste(daten, "io"), start=1):
        modul = zelle(eintrag.get("device_or_submodule"))
        kanal = zelle(eintrag.get("exact_channel"))
        variable = zelle(eintrag.get("variable"))
        if not (modul or kanal or variable):
            continue
        if not (modul and kanal and variable):
            print("  Warnung: io[%d] unvollständig, wird übersprungen." % nr)
            continue
        belegung.append((modul, kanal, variable))

    # Adressen für PLC_PRG (optional): Sensoren und Aktoren
    adressen = daten.get("addresses", {})
    if not isinstance(adressen, dict):
        raise ValueError("Fehler: 'addresses' muss in der JSON-Datei ein Objekt sein.")
    sensoren, aktoren = [], []
    for schluessel, ziel in (("Sensors", sensoren), ("Actuators", aktoren)):
        for nr, eintrag in enumerate(_liste(adressen, schluessel), start=1):
            variable, adresse = zelle(eintrag.get("variable")), zelle(eintrag.get("address"))
            if not (variable or adresse):
                continue
            if not (variable and adresse):
                print("  Warnung: addresses/%s[%d] unvollständig, wird übersprungen." % (schluessel, nr))
                continue
            if not ADRESSE_MUSTER.match(adresse):
                print("  Warnung: addresses/%s[%d]: Adresse '%s' hat kein gültiges Format (z. B. %%IX0.0)."
                      % (schluessel, nr, adresse))
            ziel.append((variable, adresse))

    return sps, plc_suchtext, geraete, module, belegung, sensoren, aktoren


# ======================================================================
# XML-ORDNER LESEN UND PRÜFEN
# ======================================================================
def xml_dateien(ordner):
    gefunden = []
    for wurzel, _, dateien in os.walk(ordner):
        for d in sorted(dateien):
            if d.lower().endswith(".xml"):
                gefunden.append(os.path.abspath(os.path.join(wurzel, d)))
    return sorted(gefunden)


def lokal(tag):
    return tag.rsplit("}", 1)[-1]


def xml_infos(pfade):
    """Sammelt alle Texte und Modulnamen der GSDML-Dateien (nur für Warnungen)."""
    texte, modulnamen = set(), set()
    for p in pfade:
        try:
            wurzel = ET.parse(p).getroot()
        except ET.ParseError as fehler:
            print("  Warnung: XML nicht lesbar (%s): %s" % (os.path.basename(p), fehler))
            continue
        tid = {}
        for e in wurzel.iter():
            if lokal(e.tag) == "Text":
                tid[e.get("TextId")] = e.get("Value", "")
                texte.add(norm(e.get("Value", "")))
        for m in wurzel.iter():
            if lokal(m.tag) == "ModuleItem":
                for n in m.iter():
                    if lokal(n.tag) == "Name":
                        modulnamen.add(norm(tid.get(n.get("TextId"), n.get("TextId") or "")))
                        break
    return texte, modulnamen


def pruefen(geraete, module, texte, modulnamen):
    """Warnt vor Namen, die in keiner XML vorkommen (z. B. Tippfehler)."""
    for name, suchtext, _ in geraete[2:]:
        if not any(norm(suchtext) in t for t in texte):
            print("  Warnung: Gerät '%s' (%s) kommt in keiner XML vor - Tippfehler?" % (name, suchtext))
    for name, modul, eltern, _ in module:
        if norm(modul) not in modulnamen:
            print("  Hinweis: Modul '%s' (%s) steht in keiner XML. Falls das Gerät es automatisch "
                  "anlegt, wird es ignoriert; sonst gibt es einen Fehler." % (name, modul))


# ======================================================================
# VORLAGE DES ERZEUGTEN CODESYS-SKRIPTS (IronPython 2.7)
# ======================================================================
VORLAGE = r'''# -*- coding: utf-8 -*-
"""
Programm : steuerungskonfig_generator.py   (AUTOMATISCH ERZEUGT, nicht von Hand ändern)
Zweck    : Fügt in das geöffnete CODESYS-Projekt (V3.5 SP21 Patch 2) die
           Steuerungskonfiguration ein und weist den E/A-Kanälen Variablennamen zu
           (E/A-Belegung). Alle Daten stammen aus der JSON-Datei.
Start    : Zuerst das Projekt in CODESYS öffnen, dann
           Tools -> Scripting -> Skriptdatei ausführen
Hinweis  : Das Projekt wird NICHT automatisch gespeichert (danach Strg+S).
Hinweis  : Läuft im Scripting Engine (IronPython 2.7), nicht in normalem Python.
"""
from __future__ import print_function
import os
import re

# ======================================================================
# KONFIGURATION (aus JSON-Datei und Ordner "GSD")
# ======================================================================
# Gerätebeschreibungen (GSDML/XML) aus dem Ordner "GSD"
GSD_DATEIEN = [
@@GSD_DATEIEN@@
]
GSD_NEU_INSTALLIEREN = @@GSD_NEU@@    # True = GSDML immer installieren

# Steuerung: (Name im Projekt, Suchtext im Geräte-Repository)
STEUERUNG = (@@PLC_NAME@@, @@PLC_SUCHTEXT@@)
STEUERUNG_VERSION = @@PLC_VERSION@@

# Geräte: (Name im Projekt, Suchtext im Geräte-Repository, Name des Elterngeräts)
GERAETE = [
@@GERAETE@@
]

# Module: (Name im Projekt, Modulname laut GSDML, Name des Elterngeräts, Suchtext des Elterngeräts)
MODULE = [
@@MODULE@@
]

# E/A-Belegung: (Gerät oder Untermodul, Kanal, Variablenname)
E_A_BELEGUNG = [
@@E_A@@
]

# ST-Programm: Sensoren (Variable des Bausteins, reale Adresse) -> Variable := Adresse
SENSOREN = [
@@SENSOREN@@
]
# Aktoren (Variable des Bausteins, reale Adresse) -> Adresse := Variable
AKTOREN = [
@@AKTOREN@@
]
PROGRAMM_NAME = @@PROGRAMM_NAME@@
SCHALTER_NAME = @@SCHALTER_NAME@@          # True = reale E/A, False = Simulation
SIMULATION_AUFRUF = @@SIMULATION_AUFRUF@@

# Von Geräten automatisch angelegte Untermodule: Gerätename -> Menge normalisierter Namen
AUTO_MODULE = {}


# ======================================================================
# GERÄTE-REPOSITORY
# ======================================================================
def _norm(text):
    """Vergleichsform: ohne Leerzeichen, '_' und '-', kleingeschrieben."""
    return re.sub(r"[\s_\-]+", "", str(text)).lower()


def _feld(objekt, *namen):
    """Liest das erste vorhandene Attribut (Schreibweise je nach Version verschieden)."""
    for n in namen:
        wert = getattr(objekt, n, None)
        if wert is not None:
            return wert
    return None


def alle_geraete():
    return list(device_repository.get_all_devices())


def geraet_name(g):
    return str(g.device_info.name)


def versions_schluessel(g):
    """Wandelt '3.5.21.20' in (3, 5, 21, 20) um, damit Versionen sortierbar sind."""
    text = str(_feld(g.device_id, "version", "Version"))
    try:
        return tuple(int(t) for t in text.split("."))
    except ValueError:
        return (0,)


def geraet_suchen(suchtext, version=""):
    """Sucht ein Gerät im Repository (bei mehreren Versionen die vorgegebene oder die neueste)."""
    ziel = _norm(suchtext)
    treffer = [g for g in alle_geraete() if ziel in _norm(geraet_name(g))]
    if not treffer:
        raise Exception("Gerät nicht im Repository gefunden: " + suchtext)
    genau = [g for g in treffer if _norm(geraet_name(g)) == ziel]
    kandidaten = genau if genau else treffer
    if version:
        passend = [g for g in kandidaten
                   if str(_feld(g.device_id, "version", "Version")) == version]
        if passend:
            kandidaten = passend
        else:
            print("  Hinweis: Version %s für '%s' nicht gefunden, nehme neueste" % (version, suchtext))
    auswahl = max(kandidaten, key=versions_schluessel)
    print("  Gewählt: %s | %s" % (geraet_name(auswahl), auswahl.device_id))
    return auswahl


def _gsd_bereits_installiert(pfad):
    """Prüft anhand des Dateinamens, ob die GSDML schon im Repository ist."""
    stamm = os.path.splitext(os.path.basename(pfad))[0]
    kern = re.sub(r"^GSDML-V[\d.]+-", "", stamm)     # 'GSDML-V2.34-Buerkert-...' -> 'Buerkert-...'
    return any(kern in str(g.device_id) for g in alle_geraete())


def gsd_installieren():
    """Installiert die GSDML-Dateien im Geräte-Repository (wie Tools -> Geräte-Repository)."""
    from System import Guid
    wandler = Guid("{C633F245-876F-45E8-AAB4-3FBD994C08B8}")   # Konverter für Fremdformate
    try:
        quellen = list(device_repository.sources)
    except Exception:
        quellen = [device_repository.sources[0]]
    for pfad in GSD_DATEIEN:
        if not os.path.exists(pfad):
            print("GSDML-Datei nicht gefunden:", pfad)
            continue
        if not GSD_NEU_INSTALLIEREN and _gsd_bereits_installiert(pfad):
            print("GSDML bereits installiert:", os.path.basename(pfad))
            continue
        letzter_fehler = None
        for quelle in quellen:
            try:
                if device_repository.import_device(pfad, quelle, wandler) is not None:
                    print("GSDML installiert:", os.path.basename(pfad))
                    break
            except Exception as fehler:
                letzter_fehler = fehler
        else:
            print("GSDML-Installation fehlgeschlagen (evtl. keine GSDML):",
                  os.path.basename(pfad), "|", letzter_fehler)


# ======================================================================
# STEUERUNGSKONFIGURATION
# ======================================================================
def objekt_finden(projekt, name):
    treffer = projekt.find(name, True)
    if not treffer:
        raise Exception("Objekt nicht gefunden: " + name)
    return treffer[0]


def kinder_namen(objekt):
    """Normalisierte Namen aller Unterobjekte eines Geräts (für die Erkennung automatischer Module)."""
    namen = set()
    try:
        kinder = objekt.get_children(True)
    except Exception:
        return namen
    for k in kinder:
        try:
            namen.add(_norm(k.get_name()))
        except Exception:
            pass
    return namen


def geraet_anhaengen(eltern, name, suchtext, version=""):
    """Hängt ein Gerät an das Elternobjekt an (entspricht 'Gerät anhängen')."""
    did = geraet_suchen(suchtext, version).device_id
    return eltern.add(name, _feld(did, "type", "Type"), _feld(did, "id", "Id"),
                      _feld(did, "version", "Version"))


def gsd_module_lesen(pfad, dap_id):
    """Liest (ModuleItem-ID, Name) der für den Zugangspunkt zulässigen Module aus der GSDML.
    Gibt None zurück, wenn der Zugangspunkt in dieser Datei nicht vorkommt."""
    import clr
    clr.AddReference("System.Xml")
    from System.Xml import XmlDocument
    dok = XmlDocument()
    dok.Load(pfad)
    dap = dok.SelectSingleNode("//*[local-name()='DeviceAccessPointItem'][@ID='%s']" % dap_id)
    if dap is None:
        return None
    texte = {}
    for t in dok.SelectNodes("//*[local-name()='Text']"):
        texte[t.GetAttribute("TextId")] = t.GetAttribute("Value")
    erlaubt = set()
    for r in dap.SelectNodes(".//*[local-name()='ModuleItemRef']"):
        erlaubt.add(r.GetAttribute("ModuleItemTarget"))
    ergebnis = []
    for m in dok.SelectNodes("//*[local-name()='ModuleItem']"):
        modul_id = m.GetAttribute("ID")
        if erlaubt and modul_id not in erlaubt:
            continue
        name_knoten = m.SelectSingleNode(".//*[local-name()='Name']")
        tid = name_knoten.GetAttribute("TextId") if name_knoten is not None else ""
        ergebnis.append((modul_id, texte.get(tid, tid)))
    return ergebnis


def modul_ist_automatisch(projekt, name, modul_name, eltern_name):
    """True, wenn das Gerät dieses Modul schon selbst angelegt hat (dann wird es ignoriert)."""
    vorhanden = AUTO_MODULE.get(eltern_name, set())
    if _norm(name) in vorhanden or _norm(modul_name) in vorhanden:
        return True
    return bool(projekt.find(name, True))      # Objekt mit diesem Namen existiert schon


def modul_anhaengen(eltern, name, modul_name, geraet_suchtext):
    """Hängt ein Modul an ein Gerät an; die zugehörige GSDML wird in allen XML gesucht."""
    did = geraet_suchen(geraet_suchtext).device_id
    kennung = _feld(did, "id", "Id")
    dap_id = str(kennung).rsplit(" - ", 1)[-1]                 # z. B. 'DIM 20'
    for pfad in GSD_DATEIEN:
        if not os.path.exists(pfad):
            continue
        module = gsd_module_lesen(pfad, dap_id)
        if module is None:
            continue
        treffer = [i for i, n in module if _norm(n) == _norm(modul_name)]
        if treffer:
            return eltern.add(name, _feld(did, "type", "Type"), kennung,
                              _feld(did, "version", "Version"), treffer[0])
    raise Exception("Modul nicht in den GSDML-Dateien gefunden: " + modul_name)


# ======================================================================
# E/A-BELEGUNG
# ======================================================================
def _kanal_in(param, kanal, tiefe=0):
    """Sucht rekursiv einen abbildbaren E/A-Kanal (z. B. Bit0 unter 'Eingänge')."""
    if tiefe > 6:
        return None
    namen = [str(getattr(param, a)) for a in ("name", "visible_name")
             if getattr(param, a, None)]
    if kanal in namen and getattr(param, "is_mappable_io", False):
        return param
    if getattr(param, "has_sub_elements", False):
        for i in range(param.Count):
            gefunden = _kanal_in(param[i], kanal, tiefe + 1)
            if gefunden is not None:
                return gefunden
    return None


def variable_zuweisen(geraet, kanal, variable):
    """Schreibt den Variablennamen in die Spalte 'Variable' des E/A-Abbilds."""
    for c in geraet.connectors:
        if str(c.connector_role) != "Child":
            continue
        for param in c.host_parameters:
            gefunden = _kanal_in(param, kanal)
            if gefunden is not None:
                gefunden.io_mapping.variable = variable
                return True
    return False


# ======================================================================
# ST-PROGRAMM PLC_PRG
# ======================================================================
def st_deklaration_bauen():
    return "PROGRAM %s\nVAR\n    %s : BOOL;\nEND_VAR\n" % (PROGRAMM_NAME, SCHALTER_NAME)


def st_code_bauen():
    """Erzeugt den ST-Code: reale Eingänge -> Bausteine, Bausteine -> reale Ausgänge."""
    teile = []
    if SENSOREN:
        zeilen = ["IF %s THEN" % SCHALTER_NAME]
        zeilen += ["    %s := %s;" % (variable, adresse) for variable, adresse in SENSOREN]
        zeilen += ["ELSE", "    " + SIMULATION_AUFRUF, "END_IF"]
        teile.append("\n".join(zeilen))
    if AKTOREN:
        zeilen = ["IF %s THEN" % SCHALTER_NAME]
        zeilen += ["    %s := %s;" % (adresse, variable) for variable, adresse in AKTOREN]
        zeilen += ["END_IF"]
        teile.append("\n".join(zeilen))
    return "\n\n".join(teile) + "\n"


def anwendung_finden(projekt):
    """Sucht die Application unterhalb der Steuerung."""
    sps = objekt_finden(projekt, STEUERUNG[0])
    kinder = list(sps.get_children(True))
    for o in kinder:
        if getattr(o, "is_application", False):
            return o
    for o in kinder:
        if o.get_name() == "Application":
            return o
    raise Exception("Application unterhalb von '%s' nicht gefunden." % STEUERUNG[0])


def programm_erstellen(projekt):
    """Legt das ST-Programm an (bzw. überschreibt ein vorhandenes gleichen Namens)."""
    vorhanden = projekt.find(PROGRAMM_NAME, True)
    if vorhanden:
        pou = vorhanden[0]
        print("  Programm existiert bereits, Inhalt wird ersetzt:", PROGRAMM_NAME)
    else:
        pou = anwendung_finden(projekt).create_pou(PROGRAMM_NAME, PouType.Program,
                                                   ImplementationLanguages.st)
    pou.textual_declaration.replace(st_deklaration_bauen())
    pou.textual_implementation.replace(st_code_bauen())


# ======================================================================
# HAUPTPROGRAMM
# ======================================================================
def objekt_holen_oder_anlegen(projekt, name, suchtext, eltern_name, version=""):
    """Nimmt ein schon vorhandenes Objekt oder hängt das Gerät neu an."""
    vorhanden = projekt.find(name, True)
    if vorhanden:
        print("  Bereits vorhanden, wird verwendet:", name)
        return vorhanden[0]
    neu = geraet_anhaengen(objekt_finden(projekt, eltern_name) if eltern_name else projekt,
                           name, suchtext, version)
    return neu if neu is not None else objekt_finden(projekt, name)


def erstellen():
    # 1) Das aktuell geöffnete Projekt verwenden
    projekt = projects.primary
    if projekt is None:
        raise Exception("Kein CODESYS-Projekt geöffnet. Bitte zuerst ein Projekt öffnen.")
    try:
        print("Projekt:", projekt.path)
    except Exception:
        pass

    # 2) Gerätebeschreibungen installieren
    gsd_installieren()

    # 3) Steuerung: vorhandene SPS verwenden, sonst neu hinzufügen
    print("Steuerung:", STEUERUNG[0])
    try:
        objekt_holen_oder_anlegen(projekt, STEUERUNG[0], STEUERUNG[1], None, STEUERUNG_VERSION)
    except Exception as fehler:
        print("  FEHLER bei", STEUERUNG[0], ":", fehler)

    # 4) Steuerungskonfiguration: Geräte anhängen und deren automatische Module merken
    for name, suchtext, eltern_name in GERAETE:
        print("Füge hinzu:", name)
        try:
            neu = objekt_holen_oder_anlegen(projekt, name, suchtext, eltern_name)
            AUTO_MODULE[name] = kinder_namen(neu)
            if AUTO_MODULE[name]:
                print("  Vorhandene Untermodule:", ", ".join(sorted(AUTO_MODULE[name])))
        except Exception as fehler:
            print("  FEHLER bei", name, ":", fehler)

    #    Module aus der JSON-Datei; bereits vorhandene/automatische werden ignoriert
    for name, modul_name, eltern_name, geraet_suchtext in MODULE:
        try:
            if modul_ist_automatisch(projekt, name, modul_name, eltern_name):
                print("Modul wird ignoriert (existiert bereits):", name)
                continue
            print("Füge Modul hinzu:", name)
            modul_anhaengen(objekt_finden(projekt, eltern_name), name, modul_name, geraet_suchtext)
        except Exception as fehler:
            print("  FEHLER bei Modul", name, ":", fehler)

    # 5) E/A-Belegung: Variablennamen zuweisen
    for geraet_bez, kanal, variable in E_A_BELEGUNG:
        try:
            ok = variable_zuweisen(objekt_finden(projekt, geraet_bez), kanal, variable)
        except Exception as fehler:
            print("  FEHLER bei", geraet_bez, kanal, ":", fehler)
            ok = False
        print("E/A-Belegung %s.%s -> %s : %s" % (geraet_bez, kanal, variable,
                                                  "OK" if ok else "FEHLER"))

    # 6) ST-Programm PLC_PRG: reale E/A <-> Bausteine
    if SENSOREN or AKTOREN:
        print("Erzeuge ST-Programm:", PROGRAMM_NAME)
        try:
            programm_erstellen(projekt)
            print("  %d Sensor-Zuweisung(en), %d Aktor-Zuweisung(en)" % (len(SENSOREN), len(AKTOREN)))
        except Exception as fehler:
            print("  FEHLER bei", PROGRAMM_NAME, ":", fehler)
    else:
        print("Keine Sensoren/Aktoren in 'addresses' - PLC_PRG wird nicht erzeugt.")

    print("Fertig. Das Projekt wurde nicht gespeichert - bitte mit Strg+S speichern.")

erstellen()
'''


def skript_bauen(sps, plc_suchtext, geraete, module, belegung, sensoren, aktoren, gsd_pfade,
                 neu_installieren=GSD_NEU_INSTALLIEREN):
    """Setzt die Daten in die Vorlage ein."""
    def tupel(zeile):
        return "    (" + ", ".join(lit(x) for x in zeile) + "),"

    ersetzungen = {
        "@@GSD_DATEIEN@@": "\n".join("    " + lit(p) + "," for p in gsd_pfade),
        "@@GSD_NEU@@": str(bool(neu_installieren)),
        "@@PLC_NAME@@": lit(sps),
        "@@PLC_SUCHTEXT@@": lit(plc_suchtext),
        "@@PLC_VERSION@@": lit(STEUERUNG_VERSION),
        "@@GERAETE@@": "\n".join(tupel(g) for g in geraete),
        "@@MODULE@@": "\n".join(tupel(m) for m in module),
        "@@E_A@@": "\n".join(tupel(e) for e in belegung),
        "@@SENSOREN@@": "\n".join(tupel(s) for s in sensoren),
        "@@AKTOREN@@": "\n".join(tupel(a) for a in aktoren),
        "@@PROGRAMM_NAME@@": lit(PROGRAMM_NAME),
        "@@SCHALTER_NAME@@": lit(SCHALTER_NAME),
        "@@SIMULATION_AUFRUF@@": lit(SIMULATION_AUFRUF),
    }
    text = VORLAGE
    for schluessel, wert in ersetzungen.items():
        text = text.replace(schluessel, wert)
    return text


# ======================================================================
# KERNFUNKTION (von Fenster und Kommandozeile gemeinsam genutzt)
# ======================================================================
def erzeugen(json_pfad, ordner, skript_pfad, neu_installieren=GSD_NEU_INSTALLIEREN):
    """Liest JSON + XML-Ordner und schreibt das CODESYS-Skript. Fehler -> ValueError."""
    if not os.path.isfile(json_pfad):
        raise ValueError("JSON-Datei nicht gefunden: " + json_pfad)
    if not os.path.isdir(ordner):
        raise ValueError("Ordner '%s' nicht gefunden: %s" % (ORDNER_NAME, ordner))
    gsd = xml_dateien(ordner)
    if not gsd:
        print("Warnung: Keine XML-Dateien in", ordner)

    print("JSON  :", json_pfad)
    print("XML   : %d Datei(en) in %s" % (len(gsd), ordner))
    print("GSD_NEU_INSTALLIEREN =", bool(neu_installieren))
    sps, plc_suchtext, geraete, module, belegung, sensoren, aktoren = json_lesen(json_pfad)
    texte, modulnamen = xml_infos(gsd)
    pruefen(geraete, module, texte, modulnamen)

    skript_pfad = os.path.abspath(skript_pfad)
    os.makedirs(os.path.dirname(skript_pfad), exist_ok=True)
    code = skript_bauen(sps, plc_suchtext, geraete, module, belegung, sensoren, aktoren, gsd, neu_installieren)
    with open(skript_pfad, "w", encoding="utf-8") as f:
        f.write(code)
    print("Fertig: %d Geräte, %d Module, %d E/A-Zuweisungen, %d Sensoren, %d Aktoren"
          % (len(geraete), len(module), len(belegung), len(sensoren), len(aktoren)))
    print("Skript :", skript_pfad)
    print("Jetzt das Zielprojekt in CODESYS öffnen und: Tools -> Scripting -> Skriptdatei ausführen")
    return skript_pfad


# ======================================================================
# FENSTER (tkinter)
# ======================================================================
UNGUELTIG = '<>:"/\\|?*'


def fenster(hier):
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Steuerungskonfiguration & E/A-Belegung - Skript-Generator")
    root.minsize(760, 560)
    rahmen = ttk.Frame(root, padding=12)
    rahmen.pack(fill="both", expand=True)
    rahmen.columnconfigure(1, weight=1)

    # Eingabevariablen mit Standardwerten
    v_json = tk.StringVar()
    v_ordner = tk.StringVar(value=os.path.join(hier, ORDNER_NAME))
    v_sname = tk.StringVar(value=STANDARD_AUSGABE)
    v_sdir = tk.StringVar(value=hier)
    v_neu = tk.BooleanVar(value=GSD_NEU_INSTALLIEREN)

    # Standard-JSON: erste .json neben dem Programm (falls vorhanden)
    try:
        js = sorted(f for f in os.listdir(hier) if f.lower().endswith(".json"))
        if js:
            v_json.set(os.path.join(hier, js[0]))
    except OSError:
        pass

    def wahl_datei():
        p = filedialog.askopenfilename(title="JSON-Datei wählen", initialdir=hier,
                                       filetypes=[("JSON", "*.json"), ("Alle Dateien", "*.*")])
        if p:
            v_json.set(os.path.normpath(p))

    def wahl_ordner(var, titel):
        p = filedialog.askdirectory(title=titel, initialdir=var.get() or hier)
        if p:
            var.set(os.path.normpath(p))

    def zeile(r, text, var, befehl, breite=None):
        ttk.Label(rahmen, text=text).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Entry(rahmen, textvariable=var, width=breite or 60).grid(row=r, column=1, sticky="ew", pady=4)
        if befehl:
            ttk.Button(rahmen, text="Durchsuchen...", command=befehl).grid(row=r, column=2, padx=(8, 0))

    def gruppe(r, text):
        ttk.Label(rahmen, text=text, font=("TkDefaultFont", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(12, 0))

    gruppe(0, "Eingabe")
    zeile(1, "JSON-Datei:", v_json, wahl_datei)
    zeile(2, "Ordner GSDML:", v_ordner, lambda: wahl_ordner(v_ordner, "Ordner mit den XML-Dateien"))
    gruppe(3, "Erzeugtes Skript (fügt alles in das geöffnete CODESYS-Projekt ein)")
    zeile(4, "Dateiname:", v_sname, None)
    zeile(5, "Speicherort:", v_sdir, lambda: wahl_ordner(v_sdir, "Speicherort des Skripts"))

    # Schalter: XML überschreiben (-> GSD_NEU_INSTALLIEREN im erzeugten Skript)
    v_text = tk.StringVar()

    def schalter_text(*_):
        v_text.set("AN  -> GSDML_NEU_INSTALLIEREN = True  (GSDML werden bei jedem Lauf neu installiert)"
                   if v_neu.get() else
                   "AUS -> GSDML_NEU_INSTALLIEREN = False (bereits installierte GSDML bleiben unverändert)")

    ttk.Checkbutton(rahmen, text="GSDML-Dateien überschreiben", variable=v_neu,
                    command=schalter_text).grid(row=6, column=0, sticky="w", pady=(8, 0))
    ttk.Label(rahmen, textvariable=v_text).grid(row=6, column=1, columnspan=2, sticky="w", pady=(8, 0))
    schalter_text()

    log = tk.Text(rahmen, height=14, state="disabled", wrap="word")
    log.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
    rahmen.rowconfigure(9, weight=1)

    class LogSchreiber(object):
        """Leitet print()-Ausgaben in das Textfeld um."""
        def write(self, text):
            log.configure(state="normal")
            log.insert("end", text)
            log.see("end")
            log.configure(state="disabled")
            root.update_idletasks()

        def flush(self):
            pass

    def namen_pruefen(name, endung, bezeichnung):
        name = name.strip()
        if not name:
            raise ValueError("%s: Dateiname fehlt." % bezeichnung)
        if any(c in UNGUELTIG for c in name):
            raise ValueError("%s: Der Dateiname enthält ungültige Zeichen (%s)." % (bezeichnung, UNGUELTIG))
        if not name.lower().endswith(endung):
            name += endung
        return name

    def starten():
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")
        try:
            json_datei = v_json.get().strip()
            if not json_datei:
                raise ValueError("Bitte eine JSON-Datei auswählen.")
            sname = namen_pruefen(v_sname.get(), ".py", "Skript")
            sdir = v_sdir.get().strip()
            if not sdir:
                raise ValueError("Bitte einen Speicherort angeben.")
            v_sname.set(sname)
            ziel = os.path.join(sdir, sname)
            if os.path.exists(ziel) and not messagebox.askyesno(
                    "Überschreiben?", "Das Skript existiert bereits:\n%s\n\nÜberschreiben?" % ziel):
                return
            with contextlib.redirect_stdout(LogSchreiber()):
                erzeugen(json_datei, v_ordner.get().strip(), ziel, v_neu.get())
            messagebox.showinfo("Fertig", "Skript erzeugt:\n%s" % ziel)
        except Exception as fehler:          # Eingabefehler und unerwartete Fehler anzeigen
            LogSchreiber().write("FEHLER: %s\n" % fehler)
            messagebox.showerror("Fehler", str(fehler))

    leiste = ttk.Frame(rahmen)
    leiste.grid(row=7, column=0, columnspan=3, sticky="e", pady=(14, 0))
    ttk.Button(leiste, text="Skript erzeugen", command=starten).pack(side="right")
    ttk.Button(leiste, text="Schließen", command=root.destroy).pack(side="right", padx=(0, 8))
    ttk.Label(rahmen, text="Protokoll:").grid(row=8, column=0, sticky="w", pady=(10, 0))

    root.mainloop()


# ======================================================================
# HAUPTPROGRAMM
# ======================================================================
def main():
    hier = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) == 1:               # ohne Argumente: Fenster öffnen
        fenster(hier)
        return

    ap = argparse.ArgumentParser(description="Erzeugt das CODESYS-Skript aus der JSON-Datei.")
    ap.add_argument("json", help="JSON-Datei")
    ap.add_argument("-o", "--ausgabe", default=os.path.join(hier, STANDARD_AUSGABE),
                    help="Pfad des erzeugten CODESYS-Skripts")
    ap.add_argument("--nicht-ueberschreiben", action="store_true",
                    help="GSD_NEU_INSTALLIEREN = False im erzeugten Skript (Standard: True)")
    ap.add_argument("-d", "--dispositivos", default=os.path.join(hier, ORDNER_NAME),
                    help="Ordner mit den XML-Dateien")
    args = ap.parse_args()
    try:
        erzeugen(args.json, args.dispositivos, args.ausgabe, not args.nicht_ueberschreiben)
    except ValueError as fehler:
        sys.exit(str(fehler))


if __name__ == "__main__":
    main()
