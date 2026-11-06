# =========================================================
# Programm : SFC-Generator (JSON-Version mit Alternativen und Parallelzweigen)
# Autor    : Ali Duarte Perez Luna
# Datum    : 23.10.2026
# Zweck    : Liest eine JSON-Anweisungsdatei ein und erzeugt daraus eine
#            XML-Datei mit einem SFC-Programm für CODESYS.
#            Das Programm heißt immer "Ablauf" (der "program_name" im JSON
#            wird bewusst ignoriert).
#            Unterstützt: lineare Abläufe, Alternativverzweigungen
#            (selectionDivergence), Alternativzusammenführungen
#            (selectionConvergence), Parallelverzweigungen
#            (simultaneousDivergence), Parallelzusammenführungen
#            (simultaneousConvergence) und Sprünge (jumpStep).
#
#            Parallelzweige im JSON ("transitions"):
#              * "from": "StepA", "to": ["StepB", "StepC"]
#                  -> Parallelverzweigung: Mit EINER Bedingung werden beide
#                     Schritte gleichzeitig aktiv.
#              * "from": ["StepB", "StepC"], "to": "StepD"
#                  -> Parallelzusammenführung: StepD wird erst aktiv, wenn
#                     ALLE Schritte der Liste aktiv sind und die Bedingung wahr ist.
# =========================================================
import json  # Einlesen der JSON-Anweisungsdatei
import os  # Dateipfade und Betriebssystemfunktionen
import sys  # Zugriff auf Kommandozeilenargumente (Programmname)
import uuid  # Erzeugung eindeutiger IDs (UUID)
from datetime import datetime  # Zeitstempel für die XML-Datei
from xml.sax.saxutils import escape  # Maskierung von Sonderzeichen für XML
import tkinter as tk  # GUI-Grundbibliothek (Tkinter)
from tkinter import ttk, filedialog, messagebox  # Widgets, Dateidialoge und Meldungsfenster

# Fester Programmname im CODESYS-Projekt (unabhängig vom JSON)
PROGRAMM_NAME = "Ablauf"

# Erlaubte SFC-Qualifizierer für Aktionen (IEC 61131-3)
GUELTIGE_QUALIFIZIERER = {"N", "R", "S", "P", "L", "D", "SD", "DS", "SL"}


# Eigene Fehlerklasse: Die Meldung wird dem Benutzer im Fehlerfenster angezeigt
class SFCFehler(Exception):
    pass


# ---------------------------------------------------------
# HILFSFUNKTIONEN FÜR DIE JSON-AUSWERTUNG
# ---------------------------------------------------------
def aktion_zerlegen(aktion_text):
    """
    Zerlegt einen Aktionseintrag aus dem JSON in (Qualifizierer, Aktionsname).
    Beispiele:
        "Y100.OFF"     -> ("N", "Y100.OFF")   (kein Buchstabe davor = Qualifizierer N)
        "R U300.AUT"   -> ("R", "U300.AUT")
        "S U300.EXT"   -> ("S", "U300.EXT")
        "N Sollwert"   -> ("N", "Sollwert")
    """
    text = str(aktion_text).strip()
    teile = text.split(None, 1)  # am ersten Leerzeichen trennen

    # Steht vor dem Namen ein gültiger Qualifizierer, wird er verwendet
    if len(teile) == 2 and teile[0].upper() in GUELTIGE_QUALIFIZIERER:
        return teile[0].upper(), teile[1].strip()

    # Sonst gilt der Standard-Qualifizierer N
    return "N", text


# ---------------------------------------------------------
# AUFBAU DES SFC-KÖRPERS (Schritte, Transitionen, Verzweigungen, Sprünge)
# ---------------------------------------------------------
class SFCAufbau:
    """
    Läuft den Ablaufgraphen (Schritte + Transitionen) von der Initialschritt
    aus ab und erzeugt die XML-Elemente in der Reihenfolge, in der CODESYS sie
    erwartet.

    Regeln:
      * Ein Schritt mit EINER ausgehenden Transition -> normale Folge.
      * Ein Schritt mit MEHREREN ausgehenden Transitionen -> Alternativ-
        verzweigung (selectionDivergence), je Transition ein Zweig.
      * Zeigt eine Transition auf einen Schritt, der in "steps" VOR dem
        aktuellen Schritt (oder an derselben Stelle) steht, ist es ein SPRUNG
        (jumpStep).
      * Zeigt eine Transition vorwärts auf einen Schritt, der von mehreren
        Transitionen erreicht wird, werden die Zweige in einer Alternativ-
        zusammenführung (selectionConvergence) geschlossen.
      * Hat eine Transition als Ziel eine LISTE von Schritten, entsteht eine
        Parallelverzweigung (simultaneousDivergence): alle Schritte der Liste
        werden gleichzeitig aktiv.
      * Eine Transition mit einer LISTE von Quellschritten ist eine
        Parallelzusammenführung (simultaneousConvergence): die Zweige laufen
        parallel, bis alle Quellschritte erreicht sind; erst dann folgt die
        gemeinsame Transition.
    """

    def __init__(self, schritte_dict, schritt_reihenfolge, ausgehend, konv_eingaenge,
                 parallel_zusammenfuehrungen=None):
        self.schritte_dict = schritte_dict  # Schrittname -> Daten (Name, Aktionen)
        self.index = {n: i for i, n in enumerate(schritt_reihenfolge)}  # Position in "steps"
        self.ausgehend = ausgehend  # Schrittname -> Liste der ausgehenden Transitionen
        self.konv_eingaenge = konv_eingaenge  # Schrittname -> Anzahl der Vorwärts-Eingänge (nur >= 2)
        self.teile = []  # erzeugte XML-Elemente
        self.naechste_id = 0  # fortlaufende localId
        self.erzeugt = set()  # bereits erzeugte Schritte
        self.offen = {}  # Konvergenzziel -> bisher eingetroffene Transitions-IDs
        self.div_index = 0  # Position der zuletzt erzeugten Verzweigung in self.teile
        self.branch_zaehler = 0  # Zähler für die Namen "Branch0", "Branch1", ...
        # Parallelzusammenführungen: Zusammenführungs-Nr. -> {'von': [...], 'nach': ..., 'condition': ...}
        self.par_zus = parallel_zusammenfuehrungen or {}
        # Quellschritt einer Parallelzusammenführung -> Nr. der Zusammenführung
        self.par_quelle = {s: nr for nr, z in self.par_zus.items() for s in z['von']}
        # Zusammenführungs-Nr. -> bisher eingetroffene Schritt-IDs der parallelen Zweige
        self.par_offen = {}

    # Vergibt die nächste fortlaufende localId
    def neue_id(self):
        wert = self.naechste_id
        self.naechste_id += 1
        return wert

    # -- Schritt (mit Aktionsblock) ----------------------------------
    def schritt_erzeugen(self, name, eingang_ref, ist_initial):
        if name in self.erzeugt:
            raise SFCFehler(f"Der Schritt '{name}' würde mehrfach erzeugt. Bitte prüfen Sie die "
                            f"Transitionen (Rücksprünge müssen auf einen früher in 'steps' "
                            f"stehenden Schritt zeigen).")
        self.erzeugt.add(name)
        schritt = self.schritte_dict[name]

        schritt_id = self.neue_id()
        init_attribut = ' initialStep="true"' if ist_initial else ''
        verbindung_ein = ('<connectionPointIn />' if ist_initial else
                          f'<connectionPointIn>\n                <connection refLocalId="{eingang_ref}" />\n              </connectionPointIn>')
        # Ausgang für Aktionen nur bei Schritten, die Aktionen besitzen
        hat_ausgang_aktion = (not ist_initial) and bool(schritt['actions'])
        verbindung_aus_aktion = '\n              <connectionPointOutAction formalParameter="x" />' if hat_ausgang_aktion else ''
        wahr_falsch = 'TRUE' if ist_initial else 'FALSE'

        self.teile.append(f"""            <step localId="{schritt_id}"{init_attribut} name="{escape(name)}">
              <position x="0" y="0" />
              {verbindung_ein}
              <connectionPointOut formalParameter="sfc" />{verbindung_aus_aktion}
              <addData>
                <data name="http://www.3s-software.com/plcopenxml/sfc/element" handleUnknown="implementation">
                  <attributes xmlns="">
                    <attribute guid="38391c6d-6d4a-42f8-8ee7-9f45e5adafa8">{escape(name)}</attribute>
                    <attribute guid="01580b27-6378-448b-8ecb-0e4b795b58d6">FALSE</attribute>
                    <attribute guid="bc882c11-1e91-4dd8-a6b8-2075724ed18b">0</attribute>
                    <attribute guid="6844a48e-46c2-4cc8-a185-a478f3b99cc0">{wahr_falsch}</attribute>
                    <attribute guid="62e1754b-7629-4e63-9cec-10ae0c536f1f">FALSE</attribute>
                  </attributes>
                </data>
              </addData>
            </step>""")

        # Aktionsblock aufbauen, falls der Schritt Aktionen hat
        if schritt['actions']:
            aktionen_str = ""
            for qualifizierer, aktions_name in schritt['actions']:
                aktion_id = self.neue_id()
                aktionen_str += f"""              <action localId="{aktion_id}" qualifier="{qualifizierer}" duration="" indicator="">
                <relPosition x="0" y="0" />
                <reference name="{escape(aktions_name)}" />
                <connectionPointOut />
                <addData>
                  <data name="http://www.3s-software.com/plcopenxml/sfc/element" handleUnknown="implementation">
                    <attributes xmlns="">
                      <attribute guid="38391c6d-6d4a-42f8-8ee7-9f45e5adafa8">{escape(aktions_name)}</attribute>
                      <attribute guid="01580b27-6378-448b-8ecb-0e4b795b58d6">FALSE</attribute>
                      <attribute guid="bc882c11-1e91-4dd8-a6b8-2075724ed18b">0</attribute>
                      <attribute guid="e174fc0d-80b0-4a9e-a530-ca239c249a50">{qualifizierer}</attribute>
                    </attributes>
                  </data>
                </addData>
              </action>\n"""
            block_id = self.neue_id()
            self.teile.append(f"""            <actionBlock localId="{block_id}">
              <position x="0" y="0" />
              <connectionPointIn>
                <connection refLocalId="{schritt_id}" />
              </connectionPointIn>
{aktionen_str}            </actionBlock>""")

        return schritt_id

    # -- Transition (Bedingung + Transition) --------------------------
    def transition_erzeugen(self, quelle_ref, bedingung):
        eingangsvariable_id = self.neue_id()
        transition_id = self.neue_id()
        self.teile.append(f"""            <inVariable localId="{eingangsvariable_id}">
              <position x="0" y="0" />
              <connectionPointOut />
              <expression>{escape(bedingung)}</expression>
            </inVariable>
            <transition localId="{transition_id}">
              <position x="0" y="0" />
              <connectionPointIn>
                <connection refLocalId="{quelle_ref}" formalParameter="sfc" />
              </connectionPointIn>
              <condition>
                <connectionPointIn>
                  <connection refLocalId="{eingangsvariable_id}" />
                </connectionPointIn>
              </condition>
              <addData>
                <data name="http://www.3s-software.com/plcopenxml/sfc/element" handleUnknown="implementation">
                  <attributes xmlns="">
                    <attribute guid="38391c6d-6d4a-42f8-8ee7-9f45e5adafa8">{escape(bedingung)}</attribute>
                    <attribute guid="01580b27-6378-448b-8ecb-0e4b795b58d6">FALSE</attribute>
                    <attribute guid="bc882c11-1e91-4dd8-a6b8-2075724ed18b">0</attribute>
                    <attribute guid="62e1754b-7629-4e63-9cec-10ae0c536f1f">FALSE</attribute>
                    <attribute guid="8294df19-5962-4dee-a874-1051dabb0e3e">0</attribute>
                  </attributes>
                </data>
              </addData>
            </transition>""")
        return transition_id

    # -- Alternativverzweigung ----------------------------------------
    def verzweigung_erzeugen(self, schritt_id, anzahl_zweige):
        div_id = self.neue_id()
        self.div_index = len(self.teile)  # Position des Elements, falls es später benannt werden muss
        ausgaenge = '              <connectionPointOut formalParameter="sfc" />\n' * anzahl_zweige
        self.teile.append(f"""            <selectionDivergence localId="{div_id}">
              <position x="0" y="0" />
              <connectionPointIn>
                <connection refLocalId="{schritt_id}" />
              </connectionPointIn>
{ausgaenge}              <addData>
                <data name="http://www.3s-software.com/plcopenxml/sfc/element" handleUnknown="implementation">
                  <attributes xmlns="">
                    <attribute guid="01580b27-6378-448b-8ecb-0e4b795b58d6">FALSE</attribute>
                    <attribute guid="23bdaa98-72ec-41f7-817b-9dede5697086">FALSE</attribute>
                  </attributes>
                </data>
              </addData>
            </selectionDivergence>""")
        return div_id

    # -- Alternativzusammenführung -------------------------------------
    def zusammenfuehrung_erzeugen(self, transitions_ids):
        konv_id = self.neue_id()
        eingaenge = ""
        for t_id in transitions_ids:
            eingaenge += f"""              <connectionPointIn>
                <connection refLocalId="{t_id}" formalParameter="sfc" />
              </connectionPointIn>
"""
        self.teile.append(f"""            <selectionConvergence localId="{konv_id}">
              <position x="0" y="0" />
{eingaenge}              <connectionPointOut />
            </selectionConvergence>""")
        return konv_id

    # -- Sprung ---------------------------------------------------------
    def sprung_erzeugen(self, transition_id, ziel):
        sprung_id = self.neue_id()
        self.teile.append(f"""            <jumpStep localId="{sprung_id}" targetName="{escape(ziel)}">
              <position x="0" y="0" />
              <connectionPointIn>
                <connection refLocalId="{transition_id}" />
              </connectionPointIn>
              <addData>
                <data name="http://www.3s-software.com/plcopenxml/sfc/element" handleUnknown="implementation">
                  <attributes xmlns="">
                    <attribute guid="38391c6d-6d4a-42f8-8ee7-9f45e5adafa8">{escape(ziel)}</attribute>
                    <attribute guid="01580b27-6378-448b-8ecb-0e4b795b58d6">FALSE</attribute>
                  </attributes>
                </data>
              </addData>
            </jumpStep>""")
        return sprung_id

    # -- Parallelverzweigung ---------------------------------------------
    # Wird direkt hinter einer Transition erzeugt; jeder Ausgang führt direkt in einen Schritt.
    def parallel_verzweigung_erzeugen(self, transition_id, anzahl_zweige):
        div_id = self.neue_id()
        ausgaenge = '              <connectionPointOut formalParameter="sfc" />\n' * anzahl_zweige
        self.teile.append(f"""            <simultaneousDivergence name="" localId="{div_id}">
              <position x="0" y="0" />
              <connectionPointIn>
                <connection refLocalId="{transition_id}" />
              </connectionPointIn>
{ausgaenge}              <addData>
                <data name="http://www.3s-software.com/plcopenxml/sfc/element" handleUnknown="implementation">
                  <attributes xmlns="">
                    <attribute guid="01580b27-6378-448b-8ecb-0e4b795b58d6">FALSE</attribute>
                    <attribute guid="23bdaa98-72ec-41f7-817b-9dede5697086">TRUE</attribute>
                  </attributes>
                </data>
              </addData>
            </simultaneousDivergence>""")
        return div_id

    # -- Parallelzusammenführung -----------------------------------------
    # Die Eingänge sind die letzten SCHRITTE der parallelen Zweige (keine Transitionen).
    def parallel_zusammenfuehrung_erzeugen(self, schritt_ids):
        konv_id = self.neue_id()
        eingaenge = ""
        for s_id in schritt_ids:
            eingaenge += f"""              <connectionPointIn>
                <connection refLocalId="{s_id}" formalParameter="sfc" />
              </connectionPointIn>
"""
        self.teile.append(f"""            <simultaneousConvergence localId="{konv_id}">
              <position x="0" y="0" />
{eingaenge}              <connectionPointOut />
            </simultaneousConvergence>""")
        return konv_id

    # Benennt eine bereits erzeugte Verzweigung nachträglich (z. B. "Branch0").
    # CODESYS exportiert Verzweigungen, deren Zweige alle in Sprüngen enden, mit Namen.
    def verzweigung_benennen(self, teile_index, branch_name):
        alt = '                  <attributes xmlns="">\n'
        neu = alt + f'                    <attribute guid="38391c6d-6d4a-42f8-8ee7-9f45e5adafa8">{escape(branch_name)}</attribute>\n'
        self.teile[teile_index] = self.teile[teile_index].replace(alt, neu, 1)

    # -- Rekursiver Ablauf ----------------------------------------------
    # Erzeugt einen Schritt und alles, was von ihm aus folgt.
    # Rückgabe: (Art, Wert) des Zweigendes:
    #   ("sprung", Sprung-ID)            -> Zweig endete mit einem Sprung
    #   ("wartend", (Transition-ID, Ziel)) -> Sprung wurde zurückgehalten (siehe unten)
    #   ("konv", None) / (None, None)    -> Zweig endete an einer Zusammenführung o. Ä.
    def schritt_verarbeiten(self, name, eingang_ref, ist_initial=False, zurueckhalten=False):
        schritt_id = self.schritt_erzeugen(name, eingang_ref, ist_initial)

        # Letzter Schritt eines parallelen Zweigs: Der Zweig endet hier und wartet an der
        # Parallelzusammenführung auf die übrigen Zweige (siehe parallel_verarbeiten).
        if name in self.par_quelle:
            self.par_offen.setdefault(self.par_quelle[name], []).append(schritt_id)
            return ("parallel", None)

        transitionen = self.ausgehend[name]

        if len(transitionen) == 1:
            # Normale Folge: eine Transition
            return self.zweig_verarbeiten(name, transitionen[0], schritt_id, zurueckhalten)

        # Mehrere Transitionen: Alternativverzweigung
        div_id = self.verzweigung_erzeugen(schritt_id, len(transitionen))
        div_index = self.div_index
        ergebnisse = []
        for i, t in enumerate(transitionen):
            # Der Sprung des ERSTEN Zweigs wird zurückgehalten, weil CODESYS ihn bei
            # lauter Sprung-Zweigen hinter die Zusammenführung setzt (siehe unten).
            ergebnisse.append(self.zweig_verarbeiten(name, t, div_id, zurueckhalten=(i == 0)))

        erstes = ergebnisse[0]
        alle_spruenge = (erstes[0] == "wartend"
                         and all(e[0] == "sprung" for e in ergebnisse[1:]))

        if alle_spruenge:
            # Alle Zweige enden in Sprüngen. CODESYS exportiert das so:
            #   - Verzweigung mit Namen "BranchN"
            #   - Transition des ersten Zweigs -> Eingang 1 der Zusammenführung
            #   - Sprünge der übrigen Zweige  -> weitere Eingänge der Zusammenführung
            #   - hinter der Zusammenführung: Sprung zum Ziel des ersten Zweigs
            self.verzweigung_benennen(div_index, f"Branch{self.branch_zaehler}")
            self.branch_zaehler += 1
            trans_id, ziel_erster = erstes[1]
            eingaenge = [trans_id] + [e[1] for e in ergebnisse[1:]]
            konv_id = self.zusammenfuehrung_erzeugen(eingaenge)
            self.sprung_erzeugen(konv_id, ziel_erster)
            return (None, None)

        # Sonst: zurückgehaltenen Sprung des ersten Zweigs jetzt normal erzeugen
        if erstes[0] == "wartend":
            self.sprung_erzeugen(erstes[1][0], erstes[1][1])

        # Nach allen Zweigen: ist eine Zusammenführung vollständig, wird sie geschlossen
        fertig = [z for z, ids in self.offen.items() if len(ids) == self.konv_eingaenge[z]]
        if len(fertig) > 1:
            raise SFCFehler("Nach der Verzweigung am Schritt '" + name + "' enden die Zweige in mehreren "
                            "verschiedenen Zusammenführungen (" + ", ".join(fertig) + "). Das wird nicht unterstützt.")
        if fertig:
            ziel = fertig[0]
            konv_id = self.zusammenfuehrung_erzeugen(self.offen.pop(ziel))
            return self.schritt_verarbeiten(ziel, konv_id)
        return (None, None)

    # Verarbeitet eine Transition und ihr Ziel (Sprung, Zusammenführung oder nächster Schritt)
    def zweig_verarbeiten(self, quelle, t, quelle_ref, zurueckhalten=False):
        transition_id = self.transition_erzeugen(quelle_ref, t['condition'])
        return self.ziel_verarbeiten(self.index[quelle], transition_id, t['to'], zurueckhalten)

    # Verarbeitet das Ziel einer bereits erzeugten Transition.
    # quell_index = Position des Quellschritts in "steps" (zur Erkennung von Sprüngen)
    # ziel        = Schrittname (str) oder Liste von Schrittnamen (Parallelverzweigung)
    def ziel_verarbeiten(self, quell_index, transition_id, ziel, zurueckhalten=False):
        if isinstance(ziel, list):
            # Mehrere Ziele: Parallelverzweigung, alle Schritte werden gleichzeitig aktiv
            return self.parallel_verarbeiten(transition_id, ziel)

        if self.index[ziel] <= quell_index:
            # Ziel steht in "steps" vor (oder an) der aktuellen Stelle -> Sprung
            if zurueckhalten:
                return ("wartend", (transition_id, ziel))
            return ("sprung", self.sprung_erzeugen(transition_id, ziel))
        if ziel in self.konv_eingaenge:
            # Ziel wird von mehreren Zweigen erreicht -> Zweig endet an der Zusammenführung
            self.offen.setdefault(ziel, []).append(transition_id)
            return ("konv", None)
        return self.schritt_verarbeiten(ziel, transition_id, False, zurueckhalten)

    # Erzeugt eine Parallelverzweigung mit allen ihren Zweigen und schließt sie an der
    # zugehörigen Parallelzusammenführung wieder ab. Danach geht es hinter der
    # Zusammenführung mit der gemeinsamen Transition weiter.
    def parallel_verarbeiten(self, transition_id, ziele):
        div_id = self.parallel_verzweigung_erzeugen(transition_id, len(ziele))

        # Jeder Zweig beginnt direkt mit einem Schritt und läuft bis zu seinem
        # Quellschritt der Parallelzusammenführung.
        for ziel in ziele:
            ergebnis = self.schritt_verarbeiten(ziel, div_id)
            if ergebnis[0] != "parallel":
                raise SFCFehler(f"Der parallele Zweig, der mit dem Schritt '{ziel}' beginnt, "
                                f"endet nicht in einer Parallelzusammenführung. Parallele Zweige "
                                f"müssen mit einer Transition enden, bei der 'from' eine Liste ist.")

        # Vollständige Zusammenführung suchen (alle Quellschritte sind eingetroffen)
        fertig = [nr for nr, ids in self.par_offen.items()
                  if len(ids) == len(self.par_zus[nr]['von'])]
        if len(fertig) != 1:
            raise SFCFehler("Die parallelen Zweige werden nicht in einer gemeinsamen "
                            "Parallelzusammenführung zusammengeführt. Alle Schritte in der "
                            "'from'-Liste müssen aus derselben Parallelverzweigung stammen.")
        nr = fertig[0]
        zus = self.par_zus[nr]
        schritt_ids = self.par_offen.pop(nr)

        konv_id = self.parallel_zusammenfuehrung_erzeugen(schritt_ids)
        t_id = self.transition_erzeugen(konv_id, zus['condition'])
        # Sprung-Erkennung: der am weitesten hinten stehende Quellschritt zählt
        quell_index = max(self.index[s] for s in zus['von'])
        return self.ziel_verarbeiten(quell_index, t_id, zus['nach'])


# Hauptfunktion: liest die JSON-Datei ein, prüft sie und erzeugt die PLCopenXML-Datei
# Parameter: json_pfad = Pfad der JSON-Datei, ausgabe_xml_pfad = Zielpfad der XML-Datei
# Rückgabe: (Erfolg als bool, Meldung als Text)
def sfc_xml_logik_generieren(json_pfad, ausgabe_xml_pfad):
    try:
        return _sfc_xml_erzeugen(json_pfad, ausgabe_xml_pfad)
    except SFCFehler as fehler:
        return False, str(fehler)


def _sfc_xml_erzeugen(json_pfad, ausgabe_xml_pfad):
    # ---------------------------------------------------------
    # 1. JSON-DATEI EINLESEN
    # ---------------------------------------------------------
    try:
        # utf-8-sig verträgt auch eine Datei mit BOM
        with open(json_pfad, "r", encoding="utf-8-sig") as datei:
            daten = json.load(datei)
    except json.JSONDecodeError as e:
        raise SFCFehler(f"Die JSON-Datei ist syntaktisch fehlerhaft:\n{e}")
    except Exception as e:
        raise SFCFehler(f"Fehler beim Öffnen der JSON-Datei:\n{e}")

    if not isinstance(daten, dict):
        raise SFCFehler("Die JSON-Datei muss ein Objekt ({ ... }) auf oberster Ebene enthalten.")

    # ---------------------------------------------------------
    # 2. LOKALE VARIABLEN ("additional_variables")
    # ---------------------------------------------------------
    # lokale_variablen_dict: Variablenname -> Datentyp
    lokale_variablen_dict = {"Start": "BOOL"}  # Unveränderliche Regel 1

    for eintrag in daten.get("additional_variables", []):
        if not isinstance(eintrag, dict) or not str(eintrag.get("name", "")).strip():
            raise SFCFehler("Ein Eintrag in 'additional_variables' hat keinen gültigen 'name'.")
        v_name = str(eintrag["name"]).strip()
        v_typ = str(eintrag.get("type", "BOOL")).strip().upper() or "BOOL"
        lokale_variablen_dict[v_name] = v_typ

    # ---------------------------------------------------------
    # 3. AKTIONSDEFINITIONEN ("action_definitions") mit ST-Code
    # ---------------------------------------------------------
    # globale_aktionen: alle Aktionen mit Name, ST-Code und eindeutiger ID
    globale_aktionen = []
    aktionsnamen = set()
    for eintrag in daten.get("action_definitions", []):
        if not isinstance(eintrag, dict) or not str(eintrag.get("name", "")).strip():
            raise SFCFehler("Ein Eintrag in 'action_definitions' hat keinen gültigen 'name'.")
        a_name = str(eintrag["name"]).strip()
        if a_name in aktionsnamen:
            raise SFCFehler(f"Die Aktion '{a_name}' ist in 'action_definitions' mehrfach definiert.")
        aktionsnamen.add(a_name)
        globale_aktionen.append({
            'name': a_name,
            'code': str(eintrag.get("st_code", "")),
            'id': str(uuid.uuid4()),  # eindeutige ID für die Aktion
        })

    # ---------------------------------------------------------
    # 4. SCHRITTE ("steps")
    # ---------------------------------------------------------
    roh_schritte = daten.get("steps", [])
    if not isinstance(roh_schritte, list) or len(roh_schritte) < 1:
        raise SFCFehler("Das JSON muss in 'steps' mindestens einen Schritt enthalten.")

    # schritte_dict: Schrittname -> Daten (Name, Aktionen)
    # schritt_reihenfolge: Schrittnamen in der Reihenfolge des JSON
    schritte_dict = {}
    schritt_reihenfolge = []

    for position, roh in enumerate(roh_schritte, start=1):
        if not isinstance(roh, dict) or not str(roh.get("name", "")).strip():
            raise SFCFehler(f"Der Schritt Nr. {position} in 'steps' hat keinen gültigen 'name'.")
        schritt_name = str(roh["name"]).strip()
        if schritt_name in schritte_dict:
            raise SFCFehler(f"Der Schrittname '{schritt_name}' kommt in 'steps' mehrfach vor.")

        aktionen = []
        for aktion_text in roh.get("actions", []):
            if not str(aktion_text).strip():
                continue  # leere Einträge ignorieren
            aktionen.append(aktion_zerlegen(aktion_text))

        schritte_dict[schritt_name] = {'name': schritt_name, 'actions': aktionen}
        schritt_reihenfolge.append(schritt_name)

    # Initialschritt an den Anfang setzen (falls "initial_step" angegeben ist)
    initial = str(daten.get("initial_step", "")).strip()
    if initial:
        if initial not in schritte_dict:
            raise SFCFehler(f"Der 'initial_step' '{initial}' existiert nicht in 'steps'.")
        schritt_reihenfolge.remove(initial)
        schritt_reihenfolge.insert(0, initial)
    initial = schritt_reihenfolge[0]
    index = {n: i for i, n in enumerate(schritt_reihenfolge)}

    # ---------------------------------------------------------
    # 5. TRANSITIONEN ("transitions")
    # ---------------------------------------------------------
    # ausgehend: Schrittname -> Liste der ausgehenden Transitionen (in JSON-Reihenfolge)
    ausgehend = {name: [] for name in schritt_reihenfolge}
    gesehen = set()  # (von, nach) zur Erkennung doppelter Transitionen
    # parallel_zusammenfuehrungen: Nr. -> {'von': [Schritte], 'nach': Ziel, 'condition': Bedingung}
    parallel_zusammenfuehrungen = {}
    parallel_quellen = {}  # Quellschritt einer Parallelzusammenführung -> Nr.

    # Hilfsfunktion: wandelt "from"/"to" in eine Liste von Schrittnamen um.
    # Ein einzelner Name (str) ergibt eine Liste mit einem Eintrag; eine JSON-Liste bleibt eine Liste.
    def namen_liste(wert, feld, position):
        if isinstance(wert, list):
            namen = [str(n).strip() for n in wert]
            if not namen:
                raise SFCFehler(f"Transition Nr. {position}: Die Liste in '{feld}' ist leer.")
            if len(set(namen)) != len(namen):
                raise SFCFehler(f"Transition Nr. {position}: Die Liste in '{feld}' enthält einen Schritt mehrfach.")
        else:
            namen = [str(wert).strip()]
        for n in namen:
            if n not in schritte_dict:
                art = "Quellschritt" if feld == "from" else "Zielschritt"
                raise SFCFehler(f"Transition Nr. {position}: Der {art} '{n}' existiert nicht in 'steps'.")
        return namen

    for position, roh in enumerate(daten.get("transitions", []), start=1):
        if not isinstance(roh, dict):
            raise SFCFehler(f"Die Transition Nr. {position} hat ein ungültiges Format.")
        von_liste = namen_liste(roh.get("from", ""), "from", position)
        nach_liste = namen_liste(roh.get("to", ""), "to", position)
        bedingung = str(roh.get("condition", "")).strip()
        von_text = ", ".join(von_liste)
        nach_text = ", ".join(nach_liste)

        if not bedingung:
            raise SFCFehler(f"Transition Nr. {position} (von '{von_text}' nach '{nach_text}') hat keine Bedingung. "
                            f"Jede Transition benötigt eine Transitionsbedingung.")
        if (tuple(von_liste), tuple(nach_liste)) in gesehen:
            raise SFCFehler(f"Die Transition von '{von_text}' nach '{nach_text}' ist mehrfach definiert.")
        gesehen.add((tuple(von_liste), tuple(nach_liste)))

        # Ziel: einzelner Schritt (str) oder Liste (Parallelverzweigung)
        ziel = nach_liste if len(nach_liste) > 1 else nach_liste[0]

        if len(von_liste) > 1:
            # Mehrere Quellschritte: Parallelzusammenführung
            nr = len(parallel_zusammenfuehrungen)
            for q in von_liste:
                if q in parallel_quellen:
                    raise SFCFehler(f"Transition Nr. {position}: Der Schritt '{q}' ist bereits Quelle einer "
                                    f"anderen Parallelzusammenführung.")
                parallel_quellen[q] = nr
            parallel_zusammenfuehrungen[nr] = {'von': von_liste, 'nach': ziel, 'condition': bedingung}
        else:
            ausgehend[von_liste[0]].append({'to': ziel, 'condition': bedingung})

    # Ein Schritt am Ende eines parallelen Zweigs darf keine eigene ausgehende Transition haben
    for q in parallel_quellen:
        if ausgehend[q]:
            raise SFCFehler(f"Der Schritt '{q}' steht in der 'from'-Liste einer Parallelzusammenführung "
                            f"und hat zusätzlich eine eigene ausgehende Transition. Das ist nicht erlaubt.")

    # Jeder Schritt braucht mindestens eine ausgehende Transition (oder endet in einer Parallelzusammenführung)
    for name in schritt_reihenfolge:
        if not ausgehend[name] and name not in parallel_quellen:
            raise SFCFehler(f"Für den Schritt '{name}' fehlt eine Transition in 'transitions'. "
                            f"Ohne Transitionsbedingung kann CODESYS nicht zum nächsten Schritt wechseln.")

    # Vorwärts-Eingänge je Zielschritt zählen (Ziel steht in "steps" NACH der Quelle).
    # Ziele mit mindestens 2 Vorwärts-Eingängen werden als Zusammenführung aufgebaut.
    # Ziele einer Parallelverzweigung (Liste) werden nicht mitgezählt: sie hängen direkt an der Verzweigung.
    vorwaerts_eingaenge = {}
    for quelle in schritt_reihenfolge:
        for t in ausgehend[quelle]:
            if not isinstance(t['to'], list) and index[t['to']] > index[quelle]:
                vorwaerts_eingaenge[t['to']] = vorwaerts_eingaenge.get(t['to'], 0) + 1
    # Transitionen hinter einer Parallelzusammenführung zählen ebenfalls als ein Eingang
    for zus in parallel_zusammenfuehrungen.values():
        if not isinstance(zus['nach'], list) and index[zus['nach']] > max(index[s] for s in zus['von']):
            vorwaerts_eingaenge[zus['nach']] = vorwaerts_eingaenge.get(zus['nach'], 0) + 1
    konv_eingaenge = {z: n for z, n in vorwaerts_eingaenge.items() if n >= 2}

    # ---------------------------------------------------------
    # 6. SFC-KÖRPER AUFBAUEN (Schritte, Alternativen, Parallelzweige, Zusammenführungen, Sprünge)
    # ---------------------------------------------------------
    aufbau = SFCAufbau(schritte_dict, schritt_reihenfolge, ausgehend, konv_eingaenge,
                       parallel_zusammenfuehrungen)
    aufbau.schritt_verarbeiten(initial, None, ist_initial=True)

    # Prüfungen nach dem Aufbau
    if aufbau.par_offen:
        raise SFCFehler("Eine Parallelzusammenführung ist unvollständig: Nicht alle Schritte der "
                        "'from'-Liste wurden über dieselbe Parallelverzweigung erreicht.")
    if aufbau.offen:
        raise SFCFehler("Eine Zusammenführung ist unvollständig: " + ", ".join(
            f"'{z}' wird nur von {len(ids)} von {konv_eingaenge[z]} Zweigen erreicht"
            for z, ids in aufbau.offen.items()) + ". Die Zweige müssen in derselben Verzweigung beginnen.")
    nicht_erreicht = [n for n in schritt_reihenfolge if n not in aufbau.erzeugt]
    if nicht_erreicht:
        raise SFCFehler("Folgende Schritte sind vom Initialschritt aus nicht erreichbar: "
                        + ", ".join(nicht_erreicht))

    # Aktueller Zeitstempel für die XML-Kopfdaten
    jetzt = datetime.now().isoformat()

    # ---------------------------------------------------------
    # 7. XML-BLÖCKE AUFBAUEN
    # ---------------------------------------------------------
    # XML-Abschnitt für die lokalen Variablen zusammensetzen
    lokale_variablen_str = ""
    for v_name, v_typ in lokale_variablen_dict.items():
        lokale_variablen_str += f"""            <variable name="{escape(v_name)}">
              <type>
                <{escape(v_typ)} />
              </type>
            </variable>\n"""

    # aktionen_block: XML der Aktionen (ST-Code)
    # projekt_struktur_kinder: Einträge der Aktionen in der Projektstruktur
    aktionen_block = ""
    projekt_struktur_kinder = ""
    if globale_aktionen:
        aktionen_block = "\n        <actions>"
        for aktion in globale_aktionen:
            aktionen_block += f"""
          <action name="{escape(aktion['name'])}">
            <body>
              <ST>
                <xhtml xmlns="http://www.w3.org/1999/xhtml">{escape(aktion['code'])}</xhtml>
              </ST>
            </body>
            <addData>
              <data name="http://www.3s-software.com/plcopenxml/objectid" handleUnknown="discard">
                <ObjectId>{aktion['id']}</ObjectId>
              </data>
            </addData>
          </action>"""
        aktionen_block += "\n        </actions>"
        # Aktionen in die Projektstruktur eintragen
        for aktion in globale_aktionen:
            projekt_struktur_kinder += f"""\n          <Object Name="{escape(aktion['name'])}" ObjectId="{aktion['id']}" />"""

    # ---------------------------------------------------------
    # 8. VOLLSTÄNDIGE DATEI ZUSAMMENSETZEN
    # ---------------------------------------------------------
    # Kopf der XML-Datei (Projektdaten, Variablen, Aktionen)
    xml_kopf = f"""<?xml version="1.0" encoding="utf-8"?>
<project xmlns="http://www.plcopen.org/xml/tc6_0200">
  <fileHeader companyName="" productName="CODESYS" productVersion="CODESYS V3.5 SP21 Patch 2" creationDateTime="{jetzt}" />
  <contentHeader name="{PROGRAMM_NAME}.project" modificationDateTime="{jetzt}">
    <coordinateInfo>
      <fbd><scaling x="1" y="1" /></fbd>
      <ld><scaling x="1" y="1" /></ld>
      <sfc><scaling x="1" y="1" /></sfc>
    </coordinateInfo>
    <addData>
      <data name="http://www.3s-software.com/plcopenxml/projectinformation" handleUnknown="implementation">
        <ProjectInformation />
      </data>
    </addData>
  </contentHeader>
  <types>
    <dataTypes />
    <pous>
      <pou name="{PROGRAMM_NAME}" pouType="program">
        <interface>
          <localVars>
{lokale_variablen_str}          </localVars>
          <addData>
            <data name="http://www.3s-software.com/plcopenxml/attributes" handleUnknown="implementation">
              <Attributes>
                <Attribute Name="do_checks_for_implicit_code" Value="" />
              </Attributes>
            </data>
          </addData>
        </interface>{aktionen_block}
        <body>
          <SFC>"""

    # Fußteil der XML-Datei (SFC-Einstellungen, Projektstruktur)
    xml_fuss = f"""          </SFC>
        </body>
        <addData>
          <data name="http://www.3s-software.com/plcopenxml/sfcsettings" handleUnknown="implementation">
            <SFCSettings>
              <LibTitle>IecSfc</LibTitle>
              <LibCompany>System</LibCompany>
              <LibVersion>4.4.0.0</LibVersion>
              <LibNamespace>IecSfc</LibNamespace>
              <CalcActiveTransOnly>false</CalcActiveTransOnly>
              <UseDefaults>false</UseDefaults>
              <Flags>
                <EnableLimit Declare="true" Use="false" />
                <Init Declare="true" Use="false" />
                <Reset Declare="true" Use="false" />
                <QuitError Declare="true" Use="false" />
                <Pause Declare="true" Use="false" />
                <Error Declare="true" Use="false" />
                <Trans Declare="true" Use="false" />
                <ErrorStep Declare="true" Use="false" />
                <ErrorPOU Declare="true" Use="false" />
                <CurrentStep Declare="true" Use="false" />
                <Tip Declare="true" Use="false" />
                <TipMode Declare="true" Use="false" />
                <MaxFlags Declare="true" Use="false" />
                <SFCErrorAnalyzation Declare="true" Use="false" />
                <SFCErrorAnalyzationTable Declare="true" Use="false" />
                <CurrentSteps Declare="true" Use="false" />
                <MetaSteps Declare="true" Use="false" />
              </Flags>
            </SFCSettings>
          </data>
          <data name="http://www.3s-software.com/plcopenxml/objectid" handleUnknown="discard">
            <ObjectId>a1b2c3d4-0005-0005-0005-000000000005</ObjectId>
          </data>
        </addData>
      </pou>
    </pous>
  </types>
  <instances>
    <configurations />
  </instances>
  <addData>
    <data name="http://www.3s-software.com/plcopenxml/projectstructure" handleUnknown="discard">
      <ProjectStructure>
        <Object Name="{PROGRAMM_NAME}" ObjectId="a1b2c3d4-0005-0005-0005-000000000005">{projekt_struktur_kinder}
        </Object>
      </ProjectStructure>
    </data>
  </addData>
</project>"""

    # Gesamte XML-Datei zusammensetzen
    finales_xml = xml_kopf + "\n" + "\n".join(aufbau.teile) + "\n" + xml_fuss

    # XML-Datei im UTF-8-Format schreiben; bei einem Fehler wird eine Meldung zurückgegeben
    try:
        with open(ausgabe_xml_pfad, "w", encoding="utf-8") as datei:
            datei.write(finales_xml)
        return True, f"XML erfolgreich gespeichert unter:\n{ausgabe_xml_pfad}"
    except Exception as e:
        raise SFCFehler(f"Fehler beim Speichern der XML-Datei:\n{e}")


# ---------------------------------------------------------
# GRAFISCHE BENUTZEROBERFLÄCHE (GUI)
# ---------------------------------------------------------
# Hauptfenster der Anwendung (erbt von tk.Tk)
class SFCAnwendung(tk.Tk):
    def __init__(self):
        super().__init__()
        # Programmnamen für den Fenstertitel ermitteln
        dateiname = os.path.basename(sys.argv[0])
        self.title(f"SFC-Generator ({dateiname})")
        # Fenstergröße, Innenabstand und feste Größe festlegen
        self.geometry("620x420")
        self.configure(padx=20, pady=20)
        self.resizable(False, False)

        # --- Zeile 1: Anweisungsdatei (JSON) ---
        ttk.Label(self, text="Anweisungsdatei (JSON):").pack(anchor="w", pady=(0, 2))
        # Rahmen (Frame) zur Gruppierung von Eingabefeld und Button
        rahmen_json = ttk.Frame(self)
        rahmen_json.pack(fill="x", pady=(0, 15))
        self.eingabe_json = ttk.Entry(rahmen_json)
        self.eingabe_json.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(rahmen_json, text="Durchsuchen...", command=self.json_suchen).pack(side="right")

        # --- Zeile 2: Name der Ausgabe-XML-Datei ---
        ttk.Label(self, text="Name der Ausgabe-XML-Datei:").pack(anchor="w", pady=(0, 2))
        self.eingabe_xml_name = ttk.Entry(self)
        # Standardname für die Ausgabedatei vorbelegen
        self.eingabe_xml_name.insert(0, "ausgabe.xml")
        self.eingabe_xml_name.pack(fill="x", pady=(0, 15))

        # --- Zeile 3: Ordner zum Speichern der XML-Datei ---
        ttk.Label(self, text="Ordner zum Speichern der XML-Datei:").pack(anchor="w", pady=(0, 2))
        rahmen_ordner = ttk.Frame(self)
        rahmen_ordner.pack(fill="x", pady=(0, 25))

        self.eingabe_ordner = ttk.Entry(rahmen_ordner)
        # Standardordner zuweisen (Desktop oder Dokumente)
        standard_verzeichnis = os.path.join(os.path.expanduser('~'), 'Desktop')
        self.eingabe_ordner.insert(0, standard_verzeichnis)
        self.eingabe_ordner.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(rahmen_ordner, text="Durchsuchen...", command=self.ordner_suchen).pack(side="right")

        # --- Visuelle Trennlinie ---
        ttk.Separator(self, orient='horizontal').pack(fill='x', pady=(0, 15))

        # --- Hauptschaltfläche ---
        schaltflaeche_generieren = ttk.Button(self, text="XML generieren", command=self.generierung_ausfuehren)
        schaltflaeche_generieren.pack()

    # Öffnet den Dateidialog und trägt den gewählten JSON-Pfad ein
    def json_suchen(self):
        datei = filedialog.askopenfilename(
            title="JSON-Datei auswählen",
            filetypes=[("JSON-Dateien", "*.json")]
        )
        if datei:
            self.eingabe_json.delete(0, tk.END)
            self.eingabe_json.insert(0, datei)

    # Öffnet den Ordnerdialog und trägt den gewählten Zielordner ein
    def ordner_suchen(self):
        ordner = filedialog.askdirectory(title="Zielordner auswählen")
        if ordner:
            self.eingabe_ordner.delete(0, tk.END)
            self.eingabe_ordner.insert(0, ordner)

    # Liest die Eingaben, prüft sie und startet die XML-Generierung
    def generierung_ausfuehren(self):
        # Eingaben aus den Feldern lesen (Leerzeichen am Rand entfernen)
        json_pfad = self.eingabe_json.get().strip()
        xml_name = self.eingabe_xml_name.get().strip()
        ordner_pfad = self.eingabe_ordner.get().strip()

        # Vorabprüfungen der Benutzeroberfläche
        if not json_pfad or not os.path.exists(json_pfad):
            messagebox.showerror("Dateifehler", "Bitte wählen Sie eine gültige JSON-Datei aus.")
            return

        if not xml_name:
            messagebox.showerror("Namensfehler", "Bitte geben Sie einen Namen für die XML-Datei ein.")
            return

        # Endung .xml ergänzen, falls sie fehlt
        if not xml_name.lower().endswith(".xml"):
            xml_name += ".xml"

        if not ordner_pfad or not os.path.isdir(ordner_pfad):
            messagebox.showerror("Ordnerfehler", "Bitte wählen Sie einen gültigen Zielordner aus.")
            return

        # Vollständigen Ausgabepfad bilden
        ausgabe_pfad = os.path.join(ordner_pfad, xml_name)

        # Generierungslogik aufrufen
        erfolg, nachricht = sfc_xml_logik_generieren(json_pfad, ausgabe_pfad)

        # Erfolgs- oder Fehlerfenster anzeigen
        if erfolg:
            messagebox.showinfo("XML erfolgreich generiert!", nachricht)
        else:
            messagebox.showerror("Fehler in der JSON-Anweisungsdatei", nachricht)


# Programmstart: Fenster erzeugen und Ereignisschleife starten
if __name__ == "__main__":
    app = SFCAnwendung()
    app.mainloop()
