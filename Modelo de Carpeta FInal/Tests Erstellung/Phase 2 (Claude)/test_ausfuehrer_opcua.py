#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEST-AUSFÜHRER FÜR CODESYS (OPC UA)  -  VERSION "BASE" (Leitfaden v2)
=====================================================================

Programm mit grafischer Oberfläche (tkinter), das sich per OPC UA mit CODESYS verbindet, die
in einer Excel-Datei definierten Tests ausführt und das Ergebnis in die Spalte "Testergebnis"
einer neuen Excel-Datei schreibt. Ein Meldungsfenster zeigt in Echtzeit, was das Programm tut.

-----------------------------------------------------------------------------------------
ANWEISUNGEN FÜR DIE KI, DIE DIESE DATEI AN KONKRETE TESTS ANPASST (Leitfaden v2)
-----------------------------------------------------------------------------------------
Das Programm muss NICHT neu geschrieben werden: Man geht von diesem Code aus und ändert nur
das Nötige. Die KI darf JEDEN Teil ändern (änderbarer Bereich und Motor), OHNE den Benutzer
um Bestätigung zu bitten. Jede Änderung wird mit einem Kommentar "# [KI-ÄNDERUNG] was und
warum" markiert und im KI-ÄNDERUNGSPROTOKOLL weiter unten vermerkt. Die Bereichsmarkierungen
werden nicht gelöscht.

- Die PINS (Signale und Sollwert jedes Tests) entnimmt DIE KI der Excel-Datei und schreibt sie
  direkt in die Tabelle TESTS (Abschnitt 2). Das Programm interpretiert den Text der Excel-Datei
  NICHT, um Signale zu gewinnen.
- Ursache "Start" im Testvorgang -> in TESTS steht START. Das Programm setzt Ablauf.Start = 1
  mit einer positiven Flanke und beobachtet danach, ob das Akzeptanzkriterium erfüllt wird.
- Braucht ein Test eine manuelle Aktion des Benutzers, füllt die KI das Feld "manuell" dieses
  Tests aus, und das Programm sagt es AUSDRÜCKLICH im Meldungsfenster
  ("⚠ MANUELLE AKTION ERFORDERLICH").
- Keine Signale erfinden und die Verbindungsdaten nicht ändern, ohne dass der Benutzer sie
  vorgibt.

KI-ÄNDERUNGSPROTOKOLL
 - (noch keine: die KI, die diese Datei an eine konkrete Excel-Datei anpasst, füllt es hier aus)

-----------------------------------------------------------------------------------------
VERHALTEN DES MOTORS (Zusammenfassung, damit er bei der Anpassung nicht kaputtgeht)
-----------------------------------------------------------------------------------------
- Die Tests werden in der Reihenfolge der Tabelle TESTS (= Reihenfolge der Excel-Datei)
  nacheinander ausgeführt.
- Ein Test beginnt so:
      Ursache = Signale -> sobald ALLE den angegebenen Wert haben (=1). Hat der Test "manuell",
                           wird vor dem Warten der Hinweis auf die manuelle Aktion angezeigt.
      Ursache = START   -> das Programm drückt Ablauf.Start und der Test beginnt in diesem
                           Moment.
                           (Mischt die Ursache Start und Signale: drückt und wartet zusätzlich
                           auf die Signale.)
- Danach wird das Akzeptanzkriterium überwacht, bis ZUERST eines der folgenden Ereignisse
  eintritt:
      a) alle Signale sind gleichzeitig erfüllt          -> "OK"
      b) der Testvorgang des nächsten Tests ist erfüllt   -> Liste der nicht erfüllten Variablen
  Hat der nächste Test die Ursache START (oder keine Pins), kann diese Ursache NICHT beobachtet
  werden: Der laufende Test wird nur mit a) oder mit der Frage nach der Minute Wartezeit
  abgeschlossen.
- Vergeht 1 Minute ohne Fortschritt, wird der Benutzer gefragt, ob weiter gewartet werden soll
  (mit Erinnerung an die ausstehende manuelle Aktion, falls vorhanden). Antwortet er mit Nein:
  Der laufende Test behält seine nicht erfüllten Variablen (oder "Failed", falls er nicht
  begonnen hatte) und alle übrigen Tests werden "Failed".
- Ein mit "ohne_pins" markierter Test wird nicht ausgeführt: Es wird im Meldungsfenster
  gewarnt und er bleibt "Failed".
- Die OPC-UA-Knoten lauten  <CODESYS-Präfix> + SIGNAL_PREFIX + Signal  (z. B. GVL.L1.SO).
- ABLAUF.START (Leitfaden §5.4): wird OHNE das Präfix GVL. und nur für einen Augenblick
  geschrieben (False -> True -> False, positive Flanke).
      * Jeder Test mit der Ursache START drückt Ablauf.Start einmal, wenn er zum aktuellen
        Test wird.
      * RÜCKFALLREGEL (nur wenn KEIN Test der Excel-Datei die Ursache START hat):
        Startregel beim Verbinden, mit einer einzigen Lesung der Signale:
          - Testvorgang von Test 1 NICHT erfüllt          -> Ablauf.Start wird sofort aktiviert.
          - Testvorgang UND gesamtes Akzeptanzkriterium erfüllt
                                                          -> wird NICHT aktiviert; der Test
                                                             bleibt "OK" und der nächste wird
                                                             genauso geprüft.
          - Testvorgang erfüllt, aber Akzeptanzkriterium unvollständig
                                                          -> Ablauf.Start wird an dieser Stelle
                                                             aktiviert.
          - Testvorgang eines späteren Tests nicht erfüllt -> wird nicht aktiviert; es wird auf
                                                             diesen Test gewartet.
- Beim Laden der Excel-Datei wird gegen die Tabelle TESTS geprüft: Stimmt die Anzahl der Tests
  nicht überein, ist das ein FEHLER (es wird nichts ausgeführt); erscheint ein Pin nicht im
  Text seiner Zelle, ist das eine WARNUNG.
- Schlägt die OPC-UA-Verbindung fehl (oder fehlen Knoten), wird KEINE geänderte Excel-Datei
  erzeugt.

Abhängigkeiten:
    pip install asyncua openpyxl
(tkinter ist unter Windows/macOS in Python enthalten; unter Linux: apt install python3-tk)
"""

import os
import queue
import re
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from openpyxl import load_workbook
from openpyxl.styles import Alignment

# #############################################################################
# VON DER KI ÄNDERBARER BEREICH  (Abschnitte 1 und 2)
# #############################################################################

# =============================================================================
# 1. KONFIGURATION
# =============================================================================
OPC_URL = "opc.tcp://127.0.0.1:4840"      # Adresse des OPC-UA-Servers von CODESYS
OPC_USER = "Ali"                          # OPC-UA-Benutzer
OPC_PASSWORD = "Versuch123."              # OPC-UA-Passwort
OPC_CONNECT_TIMEOUT_S = 10                # Maximale Zeit für den Verbindungsaufbau

# OPC-UA-Knoten von CODESYS:  ns=<Index>;s=|var|<Gerätename>.Application.<Variable>
# - Der Namespace-Index ist meist 4 (in UaExpert sichtbar).
# - Der Gerätename ist der im CODESYS-Baum angezeigte
#   (z. B. "CODESYS Control Win V3 x64" für die simulierte SPS unter Windows).
OPC_NAMESPACE_INDEX = 4
CODESYS_DEVICE_NAME = "CODESYS Control Win V3 x64"
NODE_PREFIX = f"|var|{CODESYS_DEVICE_NAME}.Application."

# Präfix, das den Signalen der Tests vorangestellt wird (Leitfaden §3: L1.SO -> GVL.L1.SO).
SIGNAL_PREFIX = "GVL."
# Variable, die die Automatisierung startet. Wird OHNE das Präfix GVL. geschrieben (Leitfaden §4.4).
START_VARIABLE = "Ablauf.Start"

POLL_INTERVAL_S = 0.1        # In welchem Abstand die Signale gelesen werden
WAIT_PROMPT_S = 60           # Nach dieser Zeit ohne Fortschritt wird der Benutzer gefragt (1 Minute)
START_PULSE_S = 0.5          # Dauer des Impulses von Ablauf.Start (positive Flanke)

# Texte, die in die Spalte Testergebnis geschrieben werden
RESULT_OK = "OK"
RESULT_FAILED = "Failed"
RESULT_UNMET_PREFIX = "Nicht erfüllte Variablen: "

# =============================================================================
# 2. EINLESEN DER EXCEL-DATEI
# =============================================================================
# Woran jede Spalte erkannt wird: Die Überschrift (in Kleinbuchstaben) muss mit einem dieser
# Texte BEGINNEN. Die Überschriften dürfen zusätzlichen Text enthalten, z. B.
# "Testvorgang (Aktivierung der Causes)". Die Spalte "nr" ist optional.
COLUMN_PREFIXES = {
    "nr":       ("test nr", "#", "nr", "nr."),
    "trigger":  ("testvorgang",),
    "criteria": ("akzeptanz",),
    "result":   ("testergebnis",),
}

# -----------------------------------------------------------------------------
# TABELLE DER TESTS (Leitfaden v2 §3, Schritt 2). DIE KI schreibt hier die PINS jedes Tests,
# die sie selbst aus der Excel-Datei gelesen hat. Das Programm entnimmt dem Text der Excel-Datei
# KEINE Signale.
#
# Jeder Eintrag ist ein Dictionary:
#   "num"      Nummer des Tests.
#   "zeile"    Zeile der Excel-Datei, in der dieser Test steht (dort wird Testergebnis
#              geschrieben).
#   "trigger"  die Ursache (Testvorgang):  START  (Ursache "Start": das Programm drückt
#              Ablauf.Start mit einer positiven Flanke)  oder eine Liste von Tupeln
#              (Name_ohne_Präfix, Sollwert). Mischt die Ursache Start und Signale:
#              [START, ("L1.SO", 1)].
#   "criteria" Liste von Tupeln (Name_ohne_Präfix, Sollwert) des Akzeptanzkriteriums.
#   "manuell"  (optional) Text der manuellen Aktion, die der Benutzer ausführen muss. Das
#              Programm zeigt ihn im Fenster als "MANUELLE AKTION ERFORDERLICH" an, wenn der
#              Test zum aktuellen Test wird. Bei einer reinen START-Ursache ohne Wirkung.
#   "ohne_pins" (optional, anstelle von trigger/criteria) Grund, warum die Pins nicht sicher
#              ermittelt werden konnten. Der Test wird nicht ausgeführt und bleibt "Failed".
#
# Die Namen stehen OHNE das Präfix GVL. und genau so geschrieben wie in der Excel-Datei.
# Die Werte sind 0/1 (oder True/False, oder ganze Zahlen). Pins dürfen nicht erfunden werden.
#
# Beispiel für das Format (Werte nur zur Veranschaulichung, KEINE einer echten Excel-Datei):
#   TESTS = [
#       {"num": 1, "zeile": 2, "trigger": START,
#        "criteria": [("N1.OUT", 1), ("N1.LOCK", 0)]},
#       {"num": 2, "zeile": 3, "trigger": [("L3.SO", 1)],
#        "criteria": [("N1.OUT", 0), ("N1.LOCK", 1), ("Y1.LOCK", 0)],
#        "manuell": "Füllen Sie den Behälter, bis der Sensor L3.SO aktiviert wird."},
#       {"num": 3, "zeile": 4,
#        "ohne_pins": "Der Text der Excel-Datei enthält kein erkennbares Signal."},
#   ]
# -----------------------------------------------------------------------------
START = "START"      # Markierung für die Ursachen "Start" (Impuls auf Ablauf.Start)

TESTS = []           # <- DIE KI FÜLLT DIESE TABELLE mit den Pins der Excel-Datei (Leitfaden v2 §3, Schritt 2)

# #############################################################################
# ENDE DES ÄNDERBAREN BEREICHS
# #############################################################################
# MOTOR - NICHT ÄNDERN (Leitfaden v2 §2.1: die KI DARF ihn bei Bedarf ändern, ohne um
# Bestätigung zu bitten; dann muss sie einen Kommentar "# [KI-ÄNDERUNG] was und warum"
# hinterlassen, es im KI-ÄNDERUNGSPROTOKOLL vermerken und diese Markierungen nicht löschen).
# #############################################################################

# =============================================================================
# DATENMODELL UND EINLESEN DER EXCEL-DATEI
# =============================================================================
@dataclass
class Test:
    """Ein Test: eine Zeile der Excel-Datei + seine Pins (aus der Tabelle TESTS)."""
    number: int                  # Nummer des Tests
    row: int                     # Zeile in der Excel-Datei (zum Schreiben des Ergebnisses)
    trigger_text: str            # Originaltext des Testvorgangs (aus der Excel-Datei)
    criteria_text: str           # Originaltext des Akzeptanzkriteriums (aus der Excel-Datei)
    trigger: list                # [(Signal, Wert), ...], die den Test starten (ohne die Markierung START)
    criteria: list               # [(Signal, Wert), ...], die erwartet werden
    result: str = None           # Ergebnis (OK / Liste der Variablen / Failed)
    manual: str = None           # Text der manuellen Aktion, die der Benutzer ausführen muss (oder None)
    start: bool = False          # True, wenn die Ursache "Start" enthält (Impuls auf Ablauf.Start)
    no_pins: str = None          # Grund, falls der Test keine Pins hat (wird nicht ausgeführt)

    @property
    def observable(self):
        """True, wenn sich seine Ursache durch Lesen von Signalen beobachten lässt (eine Start-Ursache nicht)."""
        return not self.no_pins and not self.start and bool(self.trigger)


def _check_pin(item, num, where):
    """Prüft einen Pin (Name, Wert) der Tabelle TESTS und gibt ihn als Tupel zurück."""
    ok = (isinstance(item, (tuple, list)) and len(item) == 2
          and isinstance(item[0], str) and item[0].strip()
          and isinstance(item[1], (int, bool)))
    if not ok:
        raise ValueError(f"TESTS, Test {num}: ungültiger Pin in '{where}': {item!r} "
                         "(muss ein Tupel (Name, Wert) sein).")
    name = item[0].strip()
    if SIGNAL_PREFIX and name.startswith(SIGNAL_PREFIX):
        raise ValueError(f"TESTS, Test {num}: der Pin '{name}' trägt das Präfix {SIGNAL_PREFIX}; "
                         "er muss ohne Präfix geschrieben werden.")
    return (name, item[1])


def _build_test(entry):
    """Wandelt einen Eintrag der Tabelle TESTS in ein Test-Objekt um (und prüft sein Format)."""
    if not isinstance(entry, dict):
        raise ValueError(f"TESTS: Jeder Eintrag muss ein Dictionary sein, nicht {entry!r}.")
    num, zeile = entry.get("num"), entry.get("zeile")
    if not isinstance(num, int) or not isinstance(zeile, int):
        raise ValueError(f"TESTS: Der Eintrag {entry!r} benötigt 'num' und 'zeile' (ganze Zahlen).")
    manual = entry.get("manuell") or None
    if manual is not None and not isinstance(manual, str):
        raise ValueError(f"TESTS, Test {num}: 'manuell' muss ein Text sein.")
    if entry.get("ohne_pins"):
        return Test(num, zeile, "", "", [], [], no_pins=str(entry["ohne_pins"]))

    raw = entry.get("trigger")
    if isinstance(raw, str) and raw == START:
        start, trigger = True, []
    elif isinstance(raw, (list, tuple)) and raw:
        start = any(isinstance(i, str) and i == START for i in raw)
        trigger = [_check_pin(i, num, "trigger")
                   for i in raw if not (isinstance(i, str) and i == START)]
    else:
        raise ValueError(f"TESTS, Test {num}: 'trigger' fehlt (START oder eine Liste von Pins) "
                         "oder 'ohne_pins'.")
    raw_criteria = entry.get("criteria", [])
    if not isinstance(raw_criteria, (list, tuple)):
        raise ValueError(f"TESTS, Test {num}: 'criteria' muss eine Liste von Pins sein.")
    criteria = [_check_pin(i, num, "criteria") for i in raw_criteria]
    return Test(num, zeile, "", "", trigger, criteria, manual=manual, start=start)


def locate_columns(ws):
    """Sucht die Kopfzeile. Gibt (Kopfzeile, {'nr','trigger','criteria','result': Spaltennr.}) zurück."""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10)):
        cols = {}
        for cell in row:
            if isinstance(cell.value, str):
                txt = cell.value.strip().lower()
                for key, prefixes in COLUMN_PREFIXES.items():
                    if key not in cols and txt.startswith(tuple(p.lower() for p in prefixes)):
                        cols[key] = cell.column
                        break
        if {"trigger", "criteria", "result"} <= cols.keys():
            return row[0].row, cols
    raise ValueError(
        "Die Spalten Testvorgang / Akzeptanzkriterium / Testergebnis wurden nicht gefunden "
        "(COLUMN_PREFIXES prüfen)."
    )


def load_tests(ws):
    """
    Liest die Testzeilen der Excel-Datei und verbindet sie mit den Pins der Tabelle TESTS (die
    Pins werden NICHT aus dem Text der Excel-Datei entnommen). Gibt (Liste_der_Tests, Spalten,
    Warnungen) zurück.
    Es ist ein FEHLER, wenn die Anzahl der Tests der Excel-Datei nicht mit der der Tabelle
    übereinstimmt.
    """
    header_row, cols = locate_columns(ws)
    rows = []                                              # [(Zeile, Ursachentext, Kriterientext)]
    for r in range(header_row + 1, ws.max_row + 1):
        trig_txt = ws.cell(r, cols["trigger"]).value
        if trig_txt is None or str(trig_txt).strip() == "":
            continue                                       # leere Zeile
        crit_txt = ws.cell(r, cols["criteria"]).value or ""
        rows.append((r, str(trig_txt), str(crit_txt)))
    if not rows:
        raise ValueError("Die Excel-Datei enthält keine Tests.")
    if not TESTS:
        raise ValueError("Die Tabelle TESTS ist leer: Die KI muss sie mit den Pins der "
                         "Excel-Datei füllen (Leitfaden v2 §3, Schritt 2).")

    tests = [_build_test(entry) for entry in TESTS]
    if len(tests) != len(rows):
        raise ValueError(f"Die Excel-Datei hat {len(rows)} Tests, aber die Tabelle TESTS hat "
                         f"{len(tests)}. Der Code wurde für eine andere Excel-Datei erzeugt: "
                         "Bitten Sie die KI, ihn neu zu erzeugen.")

    warnings = []                                          # nicht blockierende Warnungen
    for t, (r, trig_txt, crit_txt) in zip(tests, rows):
        t.trigger_text, t.criteria_text = trig_txt, crit_txt
        if t.row != r:
            warnings.append(f"Test {t.number}: Die Tabelle TESTS nennt Zeile {t.row}, in der "
                            f"Excel-Datei steht er aber in Zeile {r}.")
        if t.no_pins:
            continue
        if t.start and not re.search(r"\bstart\b", trig_txt, re.IGNORECASE):
            warnings.append(f"Test {t.number}: Die Tabelle nennt die Ursache Start, aber der "
                            f"Testvorgang der Excel-Datei enthält nicht das Wort 'Start': {trig_txt!r}.")
        for sig, _ in t.trigger:
            if sig not in trig_txt:
                warnings.append(f"Test {t.number}: Der Pin {sig} (Testvorgang) erscheint nicht im "
                                f"Text der Excel-Datei: {trig_txt!r}.")
        for sig, _ in t.criteria:
            if sig not in crit_txt:
                warnings.append(f"Test {t.number}: Der Pin {sig} (Akzeptanzkriterium) erscheint nicht "
                                "im Text der Excel-Datei.")
        if t.manual and t.start and not t.trigger:
            warnings.append(f"Test {t.number}: Das Feld 'manuell' hat bei einer reinen "
                            "Start-Ursache keine Wirkung (der Impuls erfolgt sofort).")
    return tests, cols, warnings


# =============================================================================
# OPC-UA-ZUGRIFF
# =============================================================================
def signal_node_id(signal):
    """Leitfaden §3: Das Signal 'L1.SO' wird als 'GVL.L1.SO' gelesen/geschrieben (SIGNAL_PREFIX + Signal)."""
    return f"ns={OPC_NAMESPACE_INDEX};s={NODE_PREFIX}{SIGNAL_PREFIX}{signal}"


# Leitfaden §4.4: Die Startvariable wird OHNE das Präfix GVL. angesprochen.
START_NODE_ID = f"ns={OPC_NAMESPACE_INDEX};s={NODE_PREFIX}{START_VARIABLE}"


def create_client():
    """Erzeugt den OPC-UA-Client mit Benutzer und Passwort (noch ohne Verbindung)."""
    from asyncua.sync import Client
    client = Client(OPC_URL, timeout=OPC_CONNECT_TIMEOUT_S)
    client.set_user(OPC_USER)
    client.set_password(OPC_PASSWORD)
    return client


class OpcIO:
    """Kapselt den OPC-UA-Client: Lesen von Signalen und Impuls von Ablauf.Start."""

    def __init__(self, client):
        self.client = client
        self._nodes = {}

    def _node(self, node_id):
        if node_id not in self._nodes:
            self._nodes[node_id] = self.client.get_node(node_id)
        return self._nodes[node_id]

    def read(self, signals):
        """Liest mehrere Signale -> {Signal: Wert}."""
        signals = list(dict.fromkeys(signals))           # ohne Duplikate, gleiche Reihenfolge
        # asyncua.sync.Client hat KEIN get_values(): Jeder Knoten wird mit get_value() gelesen
        return {s: self._node(signal_node_id(s)).get_value() for s in signals}

    def validate(self, signals):
        """Prüft, ob alle Knoten existieren. Gibt {Knoten: Fehler} der fehlgeschlagenen zurück."""
        errors = {}
        for s in signals:
            try:
                self._node(signal_node_id(s)).get_value()
            except Exception as exc:                      # noqa: BLE001
                errors[signal_node_id(s)] = str(exc)
        try:
            self._node(START_NODE_ID).get_value()
        except Exception as exc:                          # noqa: BLE001
            errors[START_NODE_ID] = str(exc)
        return errors

    def pulse_start(self):
        """
        Ablauf.Start mit positiver Flanke (Leitfaden v2 §5.4): Es wird sichergestellt, dass der
        Wert False ist, dann nur für einen Augenblick True und wieder False.
        """
        from asyncua import ua
        node = self._node(START_NODE_ID)
        if node.get_value():                              # False sicherstellen: Die Flanke muss echt sein
            node.set_value(ua.DataValue(ua.Variant(False, ua.VariantType.Boolean)))
            time.sleep(START_PULSE_S)
        node.set_value(ua.DataValue(ua.Variant(True, ua.VariantType.Boolean)))
        time.sleep(START_PULSE_S)
        node.set_value(ua.DataValue(ua.Variant(False, ua.VariantType.Boolean)))


# =============================================================================
# LOGIK ZUR AUSFÜHRUNG DER TESTS
# =============================================================================
class StopRequested(Exception):
    """Der Benutzer hat während der Ausführung 'Anhalten' gedrückt (oder das Fenster geschlossen)."""


def condition_met(conditions, values):
    """True, wenn ALLE Signale (Signal, Wert) mit dem Gelesenen übereinstimmen."""
    return all(values.get(sig) == expected for sig, expected in conditions)


def unmet_list(conditions, values):
    """Gibt die NICHT erfüllten Bedingungen zurück: [(Signal, erwartet, gelesen)]."""
    return [(s, e, values.get(s)) for s, e in conditions if values.get(s) != e]


def format_unmet(unmet):
    """Text für Testergebnis, wenn Signale fehlschlagen (Leitfaden §5)."""
    items = "; ".join(f"{s} (erwartet {e}, gelesen {int(v) if isinstance(v, bool) else v})"
                      for s, e, v in unmet)
    return f"{RESULT_UNMET_PREFIX}{items}"


def signals_of(*condition_lists):
    """Signalnamen mehrerer Bedingungslisten, ohne Duplikate."""
    return list(dict.fromkeys(s for conds in condition_lists for s, _ in conds))


MANUAL_MARK = "⚠ MANUELLE AKTION ERFORDERLICH"


def describe_signals(conditions):
    """Text 'S1 = 1, S2 = 0' einer Liste von Bedingungen."""
    return ", ".join(f"{s} = {v}" for s, v in conditions)


def describe_trigger(test):
    """Text der Ursache eines Tests für das Fenster und die Abschlusstabelle."""
    if test.no_pins:
        return "(ohne Pins)"
    parts = (["Start"] if test.start else []) + [f"{s} = {v}" for s, v in test.trigger]
    return ", ".join(parts)


def warnung_manuelle_aktion(log, num, text, wartet_auf):
    """
    Ausdrücklicher Hinweis auf eine manuelle Aktion im Meldungsfenster (Leitfaden v2 §5.8).
    Alle Hinweise auf manuelle Aktionen verwenden dieses Format; das Fenster hebt sie hervor
    (Etikett "manual").
    """
    balken = "=" * 60
    log("", "manual")
    log(balken, "manual")
    log(f"{MANUAL_MARK} — Test {num}", "manual")
    log(f"Was zu tun ist: {text}", "manual")
    if wartet_auf:
        log(f"Das Programm wartet auf: {wartet_auf}", "manual")
    log("Das Programm wartet weiter; in diesem Fenster muss nichts gedrückt werden.", "manual")
    log(balken, "manual")


class TestRunner:
    """
    Zustandsautomat, der die Tests in zeitlicher Reihenfolge durchläuft.

    Anfangsphase (beim Verbinden) - NUR wenn kein Test die Ursache START hat (Rückfallregel):
      Die Tests werden ab dem ersten geprüft:
        * Test 1 mit NICHT erfülltem Testvorgang -> Ablauf.Start wird aktiviert.
        * Testvorgang erfüllt und das GESAMTE Akzeptanzkriterium erfüllt
                                              -> Test "OK", ohne Ablauf.Start zu aktivieren,
                                                 und der nächste Test wird geprüft.
        * Testvorgang erfüllt, aber Akzeptanzkriterium unvollständig
                                              -> Ablauf.Start wird aktiviert.
        * Testvorgang (eines Tests nach dem 1.) nicht erfüllt -> Start wird nicht aktiviert;
                                                 es wird gewartet.

    Normale Phase - Zustände des aktiven Tests (self.idx):
      - WARTEN    : Hat er "manuell", wird beim Eintritt der Hinweis auf die manuelle Aktion
                    angezeigt. Ist seine Ursache START, wird Ablauf.Start gedrückt (positive
                    Flanke). Hat er Signale, werden sie gelesen, bis sie den angegebenen Wert
                    haben (=1).
      - ANALYSIEREN: Sein Akzeptanzkriterium wird gelesen. Abgeschlossen wird mit dem, was
                    ZUERST eintritt:
            a) alle Signale sind erfüllt                -> OK
            b) der Testvorgang des nächsten ist erfüllt -> Liste der nicht erfüllten Signale
               (nur wenn dieser Testvorgang beobachtbar ist: eine START-Ursache zählt nicht)
      Vergeht WAIT_PROMPT_S ohne Fortschritt, wird der Benutzer gefragt, ob weiter gewartet
      werden soll.
    """

    def __init__(self, io, tests, ask_continue, log, stop_event):
        self.io = io
        self.tests = tests
        self.ask_continue = ask_continue      # Funktion(Nachricht) -> True/False
        self.log = log                        # Funktion(Text, Etikett=None): schreibt ins Fenster
        self.stop_event = stop_event          # threading.Event: Der Benutzer hat 'Anhalten' gedrückt
        self.idx = 0
        self.analyzing = False
        self.t_ref = time.monotonic()         # Bezugspunkt für die Wartezeit
        self.last_unmet = []
        self._last_wait_msg = None            # damit die Meldung 'Warte...' nicht wiederholt wird
        self._entered = None                  # Test, der bereits als aktueller "eingetreten" ist (manueller Hinweis gegeben)
        self._start_done = False              # Wurde Ablauf.Start für den aktuellen Test schon gedrückt?

    def _check_stop(self):
        if self.stop_event.is_set():
            raise StopRequested()

    # ---- ein Test ohne Pins wird nicht ausgeführt ---------------------------
    def _skip(self, test):
        test.result = RESULT_FAILED
        self.log(f"[WARNUNG] Test {test.number}: keine Pins definiert ({test.no_pins}). "
                 f"Er wird nicht ausgeführt; er bleibt {RESULT_FAILED}.")
        self.idx += 1
        self.analyzing = False
        self.t_ref = time.monotonic()

    # ---- Anfangsphase (Rückfallregel): Muss Ablauf.Start aktiviert werden? --
    def _initial_phase(self):
        n = len(self.tests)
        first = True                          # erster Test, der ausgeführt werden kann
        while self.idx < n:
            self._check_stop()
            test = self.tests[self.idx]
            if test.no_pins:
                self._skip(test)
                continue
            vals = self.io.read(signals_of(test.trigger, test.criteria))
            trigger_ok = condition_met(test.trigger, vals)

            if not trigger_ok:
                if first:
                    self.log("Der Testvorgang von Test 1 ist nicht erfüllt: Ablauf.Start wird aktiviert.")
                    self.io.pulse_start()
                else:
                    self.log(f"Der Testvorgang von Test {test.number} ist noch nicht erfüllt: "
                             "Ablauf.Start wird nicht aktiviert, es wird auf diesen Test gewartet.")
                return

            unmet = unmet_list(test.criteria, vals)
            if unmet:
                self.log(f"Test {test.number}: Der Testvorgang ist erfüllt, das Akzeptanzkriterium "
                         f"aber nicht (es fehlen {len(unmet)} Signale): Ablauf.Start wird aktiviert.")
                self.io.pulse_start()
                return

            # Testvorgang und Akzeptanzkriterium erfüllt: Ablauf.Start ist nicht nötig.
            test.result = RESULT_OK
            self.log(f"Test {test.number}: Testvorgang und Akzeptanzkriterium bereits erfüllt -> "
                     f"{RESULT_OK}. Ablauf.Start wird nicht aktiviert; weiter mit dem nächsten Test.")
            first = False
            self.idx += 1

    # ---- ein Test wird zum aktuellen: Hinweis auf manuelle Aktion (Leitfaden v2 §5.8) -----
    def _enter(self, test):
        if self._entered is test:
            return
        self._entered = test
        self._start_done = False
        if test.manual and test.trigger:      # bei einer reinen START-Ursache ohne Wirkung
            warnung_manuelle_aktion(self.log, test.number, test.manual, describe_signals(test.trigger))

    # ---- Ursache "Start": positive Flanke auf Ablauf.Start (Leitfaden v2 §5.4) ----------
    def _pulse_for(self, test):
        self.log(f"Test {test.number}: Ursache Start -> Impuls auf {START_VARIABLE} "
                 "(False -> True -> False).")
        self.io.pulse_start()
        self._start_done = True
        self.log(f"   Impuls auf {START_VARIABLE} ausgeführt.")

    # ---- der Testvorgang ist erfüllt (oder Start gedrückt): Die Analyse beginnt ------
    def _begin(self, test):
        self.analyzing = True
        self.t_ref = time.monotonic()
        self.log("")
        self.log(f">>> Test {test.number}/{self.tests[-1].number} wird ausgeführt: "
                 f"{test.trigger_text}")

    # ---- Abschluss eines Tests ----------------------------------------------
    def _close(self, result):
        test = self.tests[self.idx]
        test.result = result
        shown = result if len(result) < 110 else result[:107] + "..."
        self.log(f"   -> Test {test.number} abgeschlossen: {shown}")
        self.idx += 1
        self.analyzing = False
        self.t_ref = time.monotonic()

    # ---- der Benutzer beschließt, nicht weiter zu warten --------------------
    def _abort(self):
        test = self.tests[self.idx]
        if self.analyzing and self.last_unmet:
            # Er hatte bereits begonnen und wurde nicht erfüllt: Die fehlgeschlagenen Variablen werden gemeldet.
            test.result = format_unmet(self.last_unmet)
        else:
            test.result = RESULT_FAILED
        # Alle restlichen Tests konnten nicht geprüft werden -> Failed
        for t in self.tests[self.idx + 1:]:
            t.result = RESULT_FAILED
        self.log(f"   -> Der Benutzer hat beschlossen, nicht weiter zu warten: restliche Tests = {RESULT_FAILED}")
        self.idx = len(self.tests)

    # ---- Hauptschleife ------------------------------------------------------
    def run(self):
        n = len(self.tests)
        self._check_stop()

        if any(t.start for t in self.tests):
            self.log(f"Es gibt Tests mit der Ursache Start: {START_VARIABLE} wird bei jedem von "
                     "ihnen gedrückt (die Startregel beim Verbinden wird nicht angewendet).")
        else:
            self._initial_phase()
        self.t_ref = time.monotonic()

        while self.idx < n:
            self._check_stop()
            test = self.tests[self.idx]
            if test.no_pins:
                self._skip(test)
                continue
            nxt = self.tests[self.idx + 1] if self.idx + 1 < n else None
            nxt_obs = nxt if (nxt is not None and nxt.observable) else None

            if not self.analyzing:
                # --- WARTEN / aktuellen Test starten ---
                self._enter(test)
                if test.start and not self._start_done:
                    if not test.trigger:               # reine Start-Ursache
                        self._begin(test)
                        self._pulse_for(test)          # der Test beginnt mit der positiven Flanke
                        continue                       # das Kriterium wird schon in diesem Zyklus ausgewertet
                    self._pulse_for(test)              # Start + Signale: drückt und wartet auf die Signale
                vals = self.io.read(signals_of(test.trigger))
                if condition_met(test.trigger, vals):
                    if test.manual:
                        self.log(f"✔ Manuelle Aktion erkannt: {describe_signals(test.trigger)}", "ok")
                    self._begin(test)
                    continue                           # das Kriterium schon in diesem Zyklus auswerten
                waiting_for = f"Test {test.number}: " + describe_signals(test.trigger)
                if self._last_wait_msg != test.number:
                    self._last_wait_msg = test.number
                    self.log(f"Warte auf den Testvorgang von Test {test.number} ({waiting_for})...")
            else:
                # --- ANALYSIEREN des Akzeptanzkriteriums des aktuellen Tests ---
                vals = self.io.read(signals_of(test.criteria, nxt_obs.trigger if nxt_obs else []))
                unmet = unmet_list(test.criteria, vals)
                if not unmet:                                        # a) alles erfüllt
                    self._close(RESULT_OK)
                    continue
                if nxt_obs and condition_met(nxt_obs.trigger, vals):  # b) der nächste ist eingetreten
                    self._close(format_unmet(unmet))
                    continue
                self.last_unmet = unmet
                waiting_for = ("Testvorgang von Test %d: " % nxt_obs.number +
                               describe_signals(nxt_obs.trigger)) if nxt_obs else \
                              f"Erfüllung des Akzeptanzkriteriums von Test {test.number}"

            # --- Ist die Wartezeit ohne Fortschritt abgelaufen? ---
            if time.monotonic() - self.t_ref >= WAIT_PROMPT_S:
                reminder = ""
                if not self.analyzing and test.manual and test.trigger:
                    reminder = f"{MANUAL_MARK} ausstehend (Test {test.number}): {test.manual}"
                self.log(f"Seit 1 Minute wird gewartet auf: {waiting_for}. Der Benutzer wird gefragt...")
                if reminder:
                    self.log(reminder, "manual")
                question = f"Seit 1 Minute wird gewartet auf: {waiting_for}."
                if reminder:
                    question += f"\n\n{reminder}"
                if self.ask_continue(question + "\n\nMöchten Sie weiter warten?"):
                    self.log("Der Benutzer hat beschlossen, weiter zu warten.")
                    self.t_ref = time.monotonic()
                else:
                    self._abort()
                    break

            time.sleep(POLL_INTERVAL_S)


# =============================================================================
# GESAMTE SITZUNG: Excel -> Verbindung -> Tests -> Speichern
# =============================================================================
def write_results(wb, ws, tests, cols, out_path):
    """Schreibt Testergebnis und speichert die neue Excel-Datei (das Original bleibt unberührt)."""
    for t in tests:
        cell = ws.cell(t.row, cols["result"])
        # Blieb ein Test durch einen Fehler (z. B. Verbindungsverlust) ohne Ergebnis -> Failed.
        cell.value = t.result if t.result is not None else RESULT_FAILED
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(out_path)


def summary_lines(tests):
    """Abschlusstabelle #, Testvorgang, Testergebnis (als Textzeilen)."""
    lines = ["=" * 78, f"{'#':<4}{'Testvorgang':<30}Testergebnis", "-" * 78]
    for t in tests:
        trig = describe_trigger(t)
        res = t.result if t.result is not None else RESULT_FAILED
        lines.append(f"{t.number:<4}{trig:<30}{res}")
    lines.append("=" * 78)
    return lines


def run_session(in_path, out_path, log, ask_continue, stop_event):
    """
    Führt den gesamten Ablauf aus. Gibt (ok, Nachricht) zurück.

    Wichtige Regel: Die geänderte Excel-Datei wird NUR erzeugt, wenn die Verbindung zum
    OPC-UA-Server erfolgreich war und die Knoten existieren. Schlägt die Verbindung fehl, wird
    nichts erzeugt.
    """
    # --- 1) Eingabe-Excel einlesen ---
    log(f"Excel-Datei wird gelesen: {in_path}")
    try:
        wb = load_workbook(in_path)
        ws = wb.active
        tests, cols, warnings = load_tests(ws)
    except Exception as exc:                                  # noqa: BLE001
        log(f"[FEHLER] Die Excel-Datei konnte nicht gelesen werden: {exc}")
        return False, f"Die Excel-Datei konnte nicht gelesen werden:\n{exc}"

    # Die Pins jedes Tests (aus der Tabelle TESTS) werden zur Überprüfung angezeigt
    log(f"Excel-Datei geladen: {len(tests)} Tests. Die Pins stammen aus der Tabelle TESTS des Codes.")
    for w in warnings:
        log(f"  [WARNUNG] {w}")
    if warnings:
        log("  [WARNUNG] Die Excel-Datei scheint sich von der zu unterscheiden, mit der der Code "
            "erzeugt wurde; bitten Sie die KI, ihn neu zu erzeugen.")
    for t in tests:
        if t.no_pins:
            log(f"  Test {t.number}: OHNE PINS ({t.no_pins}) -> wird nicht ausgeführt, bleibt {RESULT_FAILED}.")
            continue
        log(f"  Test {t.number}: startet mit {describe_trigger(t)}")
        log(f"      prüft ({len(t.criteria)}): {t.criteria}")
        if t.manual:
            log(f"      manuelle Aktion: {t.manual}")
        if not t.criteria:
            log(f"  [WARNUNG] Test {t.number} hat keine Signale im Akzeptanzkriterium.")

    # --- 2) OPC-UA-Verbindung (Benutzer und Passwort) ---
    log("")
    log(f"Es wird versucht, sich mit {OPC_URL} als '{OPC_USER}' zu verbinden...")
    client = None
    try:
        client = create_client()
        client.connect()
    except Exception as exc:                                  # noqa: BLE001
        log(f"[FEHLER] Die Verbindung zum OPC-UA-Server war nicht möglich: {exc}")
        log("Es wurde keine geänderte Excel-Datei erzeugt.")
        return False, (f"Die Verbindung zum OPC-UA-Server ({OPC_URL}) war nicht möglich:\n{exc}\n\n"
                       "Es wurde keine Excel-Datei erzeugt.")
    log("Verbindung zum OPC-UA-Server hergestellt.")

    # --- 3) Knoten prüfen und die Tests ausführen ---
    status_msg = None
    try:
        io = OpcIO(client)
        log("Es wird geprüft, ob alle Signale in CODESYS existieren...")
        all_signals = sorted({s for t in tests for s, _ in t.trigger + t.criteria})
        errors = io.validate(all_signals)
        if errors:
            log("[FEHLER] Diese Knoten konnten nicht gelesen werden (NODE_PREFIX / "
                "OPC_NAMESPACE_INDEX / SIGNAL_PREFIX / START_VARIABLE prüfen):")
            for nid in errors:
                log(f"   - {nid}")
            log("Es wurde keine geänderte Excel-Datei erzeugt.")
            return False, ("Einige OPC-UA-Knoten konnten nicht gelesen werden (siehe Protokollfenster).\n\n"
                           "Es wurde keine Excel-Datei erzeugt.")
        log(f"Die {len(all_signals)} Signale und {START_VARIABLE} wurden gefunden.")
        log("")
        log("Die Ausführung der Tests beginnt...")

        runner = TestRunner(io, tests, ask_continue, log, stop_event)
        try:
            runner.run()
            log("")
            log("Alle Tests wurden verarbeitet.")
        except StopRequested:
            status_msg = (f"Ausführung vom Benutzer angehalten. Tests ohne Ergebnis bleiben "
                          f"{RESULT_FAILED}.")
            log(status_msg)
        except Exception as exc:                              # noqa: BLE001
            status_msg = (f"Fehler während der Ausführung: {exc}. Tests ohne Ergebnis bleiben "
                          f"{RESULT_FAILED}.")
            log(f"[FEHLER] {status_msg}")
    finally:
        log("OPC-UA-Verbindung wird geschlossen...")
        try:
            client.disconnect()
        except Exception:                                     # noqa: BLE001
            pass

    # --- 4) Geänderte Excel-Datei speichern (die Verbindung war erfolgreich) ---
    try:
        write_results(wb, ws, tests, cols, out_path)
    except Exception as exc:                                  # noqa: BLE001
        log(f"[FEHLER] Die Excel-Datei konnte nicht gespeichert werden: {exc}")
        return False, (f"Die Excel-Datei konnte nicht gespeichert werden:\n{exc}\n\n"
                       "Prüfen Sie, ob die Datei nicht in Excel geöffnet ist.")
    for line in summary_lines(tests):
        log(line)
    log(f"Excel-Datei mit Ergebnissen gespeichert unter: {out_path}")
    msg = f"Excel-Datei mit Ergebnissen gespeichert unter:\n{out_path}"
    if status_msg:
        msg = status_msg + "\n\n" + msg
    return True, msg


# =============================================================================
# GRAFISCHE OBERFLÄCHE
# =============================================================================
class App:
    """
    Hauptfenster. Die Tests laufen in einem eigenen Thread, damit das Fenster nicht einfriert;
    der Thread kommuniziert mit dem Fenster über eine Warteschlange (self.q):
        ("log", Text)                    -> Zeile zum Fenster hinzufügen
        ("ask", Nachricht, Antwortqueue) -> dem Benutzer eine Ja/Nein-Frage stellen
        ("msg", Typ, Titel, Text)        -> ein Meldungsfenster anzeigen
        ("finished",)                    -> der Thread ist beendet
    """

    def __init__(self, root):
        self.root = root
        root.title("Test-Ausführer CODESYS (OPC UA)")
        root.geometry("900x660")
        root.minsize(720, 480)

        self.q = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.closing = False

        self.var_in = tk.StringVar()      # Pfad der Eingabe-Excel-Datei
        self.var_name = tk.StringVar()    # Name der geänderten Excel-Datei
        self.var_dir = tk.StringVar()     # Ordner, in dem gespeichert wird

        # --- Oberer Bereich: Dateiauswahl ---
        frm = ttk.LabelFrame(root, text="Dateien", padding=10)
        frm.pack(fill="x", padx=10, pady=(10, 5))
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Excel-Datei mit den Tests:").grid(row=0, column=0, sticky="w", pady=4)
        self.ent_in = ttk.Entry(frm, textvariable=self.var_in)
        self.ent_in.grid(row=0, column=1, sticky="ew", padx=6)
        self.btn_in = ttk.Button(frm, text="Durchsuchen...", command=self.choose_input)
        self.btn_in.grid(row=0, column=2)

        ttk.Label(frm, text="Name der geänderten Excel-Datei:").grid(row=1, column=0, sticky="w", pady=4)
        self.ent_name = ttk.Entry(frm, textvariable=self.var_name)
        self.ent_name.grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(frm, text="Speichern im Ordner:").grid(row=2, column=0, sticky="w", pady=4)
        self.ent_dir = ttk.Entry(frm, textvariable=self.var_dir)
        self.ent_dir.grid(row=2, column=1, sticky="ew", padx=6)
        self.btn_dir = ttk.Button(frm, text="Durchsuchen...", command=self.choose_dir)
        self.btn_dir.grid(row=2, column=2)

        # --- Steuerungsschaltflächen ---
        bar = ttk.Frame(root)
        bar.pack(fill="x", padx=10, pady=5)
        self.btn_start = ttk.Button(bar, text="Tests starten", command=self.start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(bar, text="Anhalten", command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=8)
        self.lbl_status = ttk.Label(bar, text="Bereit.")
        self.lbl_status.pack(side="left", padx=12)

        # --- Textfenster mit dem Protokoll des Geschehens ---
        logf = ttk.LabelFrame(root, text="Protokoll", padding=6)
        logf.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.txt = ScrolledText(logf, wrap="word", state="disabled", height=16,
                                font=("Consolas", 10))
        self.txt.pack(fill="both", expand=True)
        self.txt.tag_configure("error", foreground="#c62828")
        self.txt.tag_configure("test", foreground="#1565c0", font=("Consolas", 10, "bold"))
        self.txt.tag_configure("ok", foreground="#2e7d32")
        self.txt.tag_configure("manual", foreground="#8a3b00", background="#fff3cd",
                               font=("Consolas", 10, "bold"))

        self._inputs = [self.ent_in, self.btn_in, self.ent_name, self.ent_dir, self.btn_dir]

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.append_log("Wählen Sie die Excel-Datei mit den Tests aus und drücken Sie 'Tests starten'.")
        self._poll()

    # ---- Dateiauswahl -------------------------------------------------------
    def choose_input(self):
        path = filedialog.askopenfilename(
            parent=self.root, title="Wählen Sie die Excel-Datei mit den Tests aus",
            filetypes=[("Excel (*.xlsx)", "*.xlsx"), ("Alle Dateien", "*.*")])
        if path:
            self.var_in.set(path)
            # Automatischer Vorschlag für Name und Ordner (änderbar)
            self.var_name.set(Path(path).stem + "_Ergebnisse.xlsx")
            self.var_dir.set(str(Path(path).parent))

    def choose_dir(self):
        folder = filedialog.askdirectory(
            parent=self.root, title="Ordner, in dem die geänderte Excel-Datei gespeichert wird",
            initialdir=self.var_dir.get() or None)
        if folder:
            self.var_dir.set(folder)

    # ---- Protokollfenster ---------------------------------------------------
    def append_log(self, text, tag=None):
        """Fügt dem Fenster eine Zeile mit Uhrzeit hinzu (nur aus dem Hauptthread)."""
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {text}\n" if text else "\n"
        tags = ()
        if tag == "manual":                                   # Hinweis auf manuelle Aktion (Leitfaden v2 §5.8)
            tags = ("manual",)
        elif tag == "ok":
            tags = ("ok",)
        elif "[FEHLER]" in text or "Es wurde keine" in text:
            tags = ("error",)
        elif text.startswith(">>>"):
            tags = ("test",)
        elif text.startswith("   -> Test") and text.rstrip().endswith(RESULT_OK):
            tags = ("ok",)
        self.txt.configure(state="normal")
        self.txt.insert("end", line, tags)
        self.txt.see("end")
        self.txt.configure(state="disabled")

    # ---- Validierung und Start ----------------------------------------------
    def start(self):
        in_path = self.var_in.get().strip()
        name = self.var_name.get().strip()
        folder = self.var_dir.get().strip()

        if not in_path or not os.path.isfile(in_path):
            messagebox.showerror("Eingabe-Excel-Datei", "Wählen Sie eine gültige Excel-Datei aus.", parent=self.root)
            return
        if not in_path.lower().endswith(".xlsx"):
            messagebox.showerror("Eingabe-Excel-Datei", "Die Datei muss eine .xlsx-Datei sein.", parent=self.root)
            return
        if not name or re.search(r'[\\/:*?"<>|]', name):
            messagebox.showerror("Name der Excel-Datei",
                                 'Geben Sie einen gültigen Dateinamen ein (ohne \\ / : * ? " < > |).',
                                 parent=self.root)
            return
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
            self.var_name.set(name)
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Speicherordner", "Wählen Sie einen vorhandenen Ordner aus.", parent=self.root)
            return

        out_path = os.path.join(folder, name)
        if os.path.abspath(out_path).lower() == os.path.abspath(in_path).lower():
            messagebox.showerror("Name der Excel-Datei",
                                 "Die geänderte Excel-Datei darf das Original nicht überschreiben.\n"
                                 "Ändern Sie den Namen oder den Ordner.", parent=self.root)
            return
        if os.path.exists(out_path) and not messagebox.askyesno(
                "Datei existiert bereits", f"'{name}' existiert bereits in diesem Ordner.\nMöchten Sie sie überschreiben?",
                parent=self.root):
            return

        # Oberfläche vorbereiten und den Arbeitsthread starten
        self.stop_event.clear()
        self._set_running(True)
        self.append_log("-" * 60)
        self.worker = threading.Thread(target=self._worker, args=(in_path, out_path), daemon=True)
        self.worker.start()

    def stop(self):
        if self.worker and self.worker.is_alive():
            self.append_log("Wird angehalten... (die Excel-Datei wird gespeichert, falls bereits eine Verbindung besteht)")
            self.stop_event.set()
            self.btn_stop.configure(state="disabled")

    def _set_running(self, running):
        state = "disabled" if running else "normal"
        for w in self._inputs:
            w.configure(state=state)
        self.btn_start.configure(state=state)
        self.btn_stop.configure(state="normal" if running else "disabled")
        self.lbl_status.configure(text="Läuft..." if running else "Bereit.")

    # ---- Arbeitsthread (berührt Widgets NICHT direkt: verwendet die Warteschlange) --------
    def _worker(self, in_path, out_path):
        def log(text, tag=None):
            self.q.put(("log", text, tag))

        def ask_continue(msg):
            reply = queue.Queue()
            self.q.put(("ask", msg, reply))
            while True:                       # auf die Antwort warten, 'Anhalten' im Blick
                try:
                    return reply.get(timeout=0.2)
                except queue.Empty:
                    if self.stop_event.is_set():
                        raise StopRequested()

        try:
            ok, msg = run_session(in_path, out_path, log, ask_continue, self.stop_event)
            self.q.put(("msg", "info" if ok else "error",
                        "Tests abgeschlossen" if ok else "Nicht abgeschlossen", msg))
        except Exception as exc:                              # noqa: BLE001
            log(f"[FEHLER] Unerwarteter Fehler: {exc}")
            self.q.put(("msg", "error", "Unerwarteter Fehler", str(exc)))
        finally:
            self.q.put(("finished",))

    # ---- Verarbeitung der Warteschlange im Hauptthread ----------------------
    def _poll(self):
        self.root.after(50, self._poll)       # zuerst neu einplanen (die Dialoge sind modal)
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    tag = item[2] if len(item) > 2 else None
                    self.append_log(item[1], tag)
                    if tag == "manual" and item[1].startswith("⚠"):
                        self.root.bell()                  # Ton, wenn ein manueller Hinweis erscheint
                elif kind == "ask":
                    answer = messagebox.askyesno("Weiter warten", item[1], parent=self.root)
                    item[2].put(answer)
                elif kind == "msg":
                    show = messagebox.showinfo if item[1] == "info" else messagebox.showerror
                    show(item[2], item[3], parent=self.root)
                elif kind == "finished":
                    self._set_running(False)
                    if self.closing:
                        self.root.destroy()
                        return
        except queue.Empty:
            pass

    def on_close(self):
        if self.worker and self.worker.is_alive():
            if messagebox.askyesno("Beenden", "Es laufen Tests.\nAnhalten und beenden?", parent=self.root):
                self.closing = True
                self.stop_event.set()
                self.lbl_status.configure(text="Wird geschlossen...")
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
