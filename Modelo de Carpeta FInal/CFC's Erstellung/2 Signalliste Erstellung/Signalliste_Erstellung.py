# ==============================================================================
# Programm:      PCE-Stellen-Generator 
# Autor:         Ali Duarte Perez Luna
# Datum:         29.09.2026
# Beschreibung:  Liest eine Excel-Anweisungsdatei (Programm, Instanz, Funktionsbaustein),
#                gleicht sie mit dem Katalog der Bausteine (Typicals) ab und erzeugt
#                eine Excel-Datei mit allen Pins je Instanz. Der Katalog kann per XML
#                erweitert werden.
# ==============================================================================

import pandas as pd  # Einlesen und Verarbeiten der Excel-Tabellen
import os  # Dateipfade und Betriebssystemfunktionen
import sys  # Zugriff auf Programmargumente (Skriptname)
import pprint  # Katalog (Dictionary) als lesbaren Text formatieren
import tkinter as tk  # Grafische Oberfläche (GUI)
import xml.etree.ElementTree as ET  # Einlesen von XML-Dateien
from tkinter import ttk, filedialog, messagebox  # GUI-Elemente, Dateidialoge und Meldungsfenster
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side  # Formatierung der Excel-Zellen
from openpyxl.utils import get_column_letter  # Spaltennummer in Spaltenbuchstabe umwandeln

# Hauptfunktion: liest die Anweisungsdatei, prüft sie und erzeugt die Ausgabe-Excel-Datei.
# eingabedatei = Pfad der Anweisungsdatei, ausgabepfad = Pfad der Ausgabedatei,
# typical_daten = Katalog der Funktionsbausteine mit ihren Pins.
def daten_verarbeiten(eingabedatei, ausgabepfad, typical_daten):
    try:
        # Anweisungsdatei einlesen
        df_eingabe = pd.read_excel(eingabedatei)
    except Exception as e:
        raise Exception(f"Fehler beim Lesen der Excel-Datei: {e}")

    # Sammelbehälter für Ergebnisse und Fehler
    ergebnisse = []  # Liste aller Ergebniszeilen
    fehlende_typicals = set()  # Bausteine, die im Katalog fehlen
    unvollstaendige_zeilen = []  # Fehlermeldungen zu unvollständigen Zeilen
    instanz_vorkommen = {}  # Schlüssel: Instanzname in Kleinbuchstaben -> Liste von (Excel-Zeile, Originalschreibweise)

    # Alle Zeilen der Anweisungsdatei durchlaufen
    for index, zeile in df_eingabe.iterrows():
        # Zeilen mit weniger als 3 Spalten überspringen
        if len(zeile) < 3:
            continue

        # Prüfen, welche der drei Spalten leer sind
        ist_leer_programm = pd.isna(zeile.iloc[0]) or str(zeile.iloc[0]).strip() in ('', 'nan')  # True, wenn die Spalte 'Programm' leer ist
        ist_leer_instanz = pd.isna(zeile.iloc[1]) or str(zeile.iloc[1]).strip() in ('', 'nan')  # True, wenn die Spalte 'Instanz' leer ist
        ist_leer_typical = pd.isna(zeile.iloc[2]) or str(zeile.iloc[2]).strip() in ('', 'nan')  # True, wenn die Spalte 'Typical' leer ist

        # Komplett leere Zeile -> einfach überspringen
        if ist_leer_programm and ist_leer_instanz and ist_leer_typical:
            continue

        # Unvollständige Zeile (nur teilweise ausgefüllt) -> sammeln und Vorgang später abbrechen
        if ist_leer_programm or ist_leer_instanz or ist_leer_typical:
            fehlende_spalten = []  # Namen der fehlenden Spalten
            if ist_leer_programm:
                fehlende_spalten.append("Programm")
            if ist_leer_instanz:
                fehlende_spalten.append("Instanz")
            if ist_leer_typical:
                fehlende_spalten.append("Funktionsbaustein (Typical)")
            excel_zeilennummer = index + 2  # +2: Kopfzeile + 1-basierte Excel-Zeilennummerierung
            unvollstaendige_zeilen.append(
                f"- Zeile {excel_zeilennummer}: fehlende Angabe(n) bei {', '.join(fehlende_spalten)}"
            )
            continue

        # Werte der drei Spalten als bereinigten Text auslesen
        wert1 = str(zeile.iloc[0]).strip()  # Programmname
        wert2 = str(zeile.iloc[1]).strip()  # Instanzname
        typical = str(zeile.iloc[2]).strip()  # Name des Funktionsbausteins (Typical)

        excel_zeilennummer = index + 2  # +2: Kopfzeile + 1-basierte Excel-Zeilennummerierung
        # Instanzname merken, um doppelte Namen zu erkennen
        instanz_schluessel = wert2.lower()  # Kleinbuchstaben: Groß-/Kleinschreibung wird nicht unterschieden
        instanz_vorkommen.setdefault(instanz_schluessel, []).append((excel_zeilennummer, wert2))

        # Baustein im Katalog vorhanden? Dann pro Pin eine Ergebniszeile anlegen
        if typical in typical_daten:
            elemente = typical_daten[typical]  # Pin-Liste des Bausteins
            
            for element in elemente:
                ergebnisse.append({
                    'Programm': wert1,
                    'Instanz': wert2,
                    'Funktionsbaustein': typical,
                    'Pins': element['Pin'],
                    'Parameter': ''
                })
        else:
            # Baustein nicht im Katalog: merken für die Fehlermeldung
            fehlende_typicals.add(typical)

    # Abbruch, falls unvollständige Zeilen gefunden wurden
    if unvollstaendige_zeilen:
        zeilen_str = "\n".join(unvollstaendige_zeilen)
        raise Exception(
            f"Vorgang abgebrochen.\n\nDie Anweisungsdatei enthält unvollständige Zeilen:\n\n{zeilen_str}\n\n"
            "Bitte ergänze die fehlenden Angaben (Programm, Instanz und Funktionsbaustein) in der Excel-Datei und versuche es erneut."
        )

    # Instanznamen herausfiltern, die mehrfach vorkommen
    doppelte_instanzen = {
        schluessel: vorkommen
        for schluessel, vorkommen in instanz_vorkommen.items()
        if len(vorkommen) > 1
    }

    # Abbruch bei doppelten Instanznamen
    if doppelte_instanzen:
        zeilen_liste = []
        for vorkommen in doppelte_instanzen.values():
            zeilennummern = ", ".join(str(zeile_nr) for zeile_nr, _ in vorkommen)  # Zeilennummern der doppelten Instanz
            original_name = vorkommen[0][1]  # Originalschreibweise des Instanznamens
            zeilen_liste.append(f"- \"{original_name}\" (Zeilen: {zeilennummern})")
        doppelte_str = "\n".join(zeilen_liste)
        raise Exception(
            f"Vorgang abgebrochen.\n\nFolgende Instanznamen sind mehrfach vergeben (Groß-/Kleinschreibung wird nicht unterschieden):\n\n{doppelte_str}\n\n"
            "Jede Instanz muss eindeutig sein. Bitte korrigiere die Instanznamen in der Excel-Datei und versuche es erneut."
        )

    # Abbruch, falls Bausteine im Katalog fehlen
    if fehlende_typicals:
        fehlende_str = "\n".join([f"- {fb}" for fb in fehlende_typicals])
        raise Exception(f"Vorgang abgebrochen.\n\nFolgende Funktionsbausteine (FB) fehlen im aktuellen Katalog:\n\n{fehlende_str}\n\nFüge die fehlenden Bausteine hinzu (z. B. über XML) und versuche es erneut.")

    # Abbruch, falls kein gültiges Ergebnis vorliegt
    if not ergebnisse:
        raise Exception("Im Dokument wurden keine gültigen Typicals gefunden. Die Datei wurde nicht erstellt.")

    # Ergebnisse in eine Tabelle (DataFrame) umwandeln
    df_final = pd.DataFrame(ergebnisse)
    
    # Excel-Datei schreiben und formatieren
    try:
        writer = pd.ExcelWriter(ausgabepfad, engine='openpyxl')  # Schreibobjekt für die Excel-Datei
        df_final.to_excel(writer, index=False, sheet_name='Abgeglichene Daten')
        
        # Zugriff auf Arbeitsmappe und Tabellenblatt für die Formatierung
        workbook = writer.book
        worksheet = writer.sheets['Abgeglichene Daten']
        
        # Rahmenstil: dünne schwarze Linie
        duenner_rahmen = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        # Kopfzeile: schwarzer Hintergrund, weiße fette Schrift, zentriert
        kopf_fuellung = PatternFill(start_color="000000", end_color="000000", fill_type="solid") 
        kopf_schrift = Font(color="FFFFFF", bold=True) 
        kopf_ausrichtung = Alignment(horizontal="center", vertical="center")
        
        # Kopfzeile formatieren und Spaltenbreite festlegen
        for spalten_nr, _ in enumerate(df_final.columns, 1):
            zelle = worksheet.cell(row=1, column=spalten_nr)
            zelle.fill = kopf_fuellung
            zelle.font = kopf_schrift
            zelle.alignment = kopf_ausrichtung
            zelle.border = duenner_rahmen
            worksheet.column_dimensions[get_column_letter(spalten_nr)].width = 25
            
        # Pastellfarben zur Unterscheidung der Programme
        # aktueller_wert = aktuelles Programm, farb_idx = Index der aktuellen Farbe
        pastellfarben = ['E8F4F8', 'EAF9E6', 'FFF9E6', 'FCE5CD', 'E6E6FA']
        aktueller_wert = None
        farb_idx = -1
        
        # Datenzeilen einfärben: die Farbe wechselt bei jedem neuen Programm
        for zeilen_idx, zeilen_daten in enumerate(df_final.itertuples(index=False), start=2):
            wert1b = zeilen_daten[0]  # Programmname der aktuellen Zeile
            
            # Neues Programm: nächste Farbe wählen
            if wert1b != aktueller_wert:
                aktueller_wert = wert1b
                farb_idx = (farb_idx + 1) % len(pastellfarben)
                
            zeilen_fuellung = PatternFill(start_color=pastellfarben[farb_idx], end_color=pastellfarben[farb_idx], fill_type="solid")
            
            # Füllung und Rahmen auf alle Zellen der Zeile anwenden
            for spalten_idx in range(1, len(df_final.columns) + 1):
                zelle = worksheet.cell(row=zeilen_idx, column=spalten_idx)
                zelle.fill = zeilen_fuellung
                zelle.border = duenner_rahmen
                
        # Datei speichern und schließen
        writer.close()
        
        return "Die Anweisungen wurden erfolgreich verarbeitet und die Excel-Datei wurde erstellt."

    # Fehler, wenn die Zieldatei noch in Excel geöffnet ist
    except PermissionError:
        raise Exception("BERECHTIGUNGSFEHLER:\nDie Zieldatei ist in Excel geöffnet. Bitte schließe sie und versuche es erneut.")

# Hauptklasse der Oberfläche (Fenster mit zwei Registerkarten)
class AppDatenabgleich:
    # Konstruktor: root = Hauptfenster
    def __init__(self, root):
        self.root = root
        # Fenstertitel mit Skriptname setzen
        skriptname = os.path.basename(sys.argv[0]) if sys.argv and sys.argv[0] else os.path.basename(__file__)  # Name des Skripts für den Fenstertitel
        self.root.title(f"Signalliste-Generator ({skriptname})")
        self.root.resizable(False, False)

        # Standardwerte für Ausgabeordner und Dateiname, danach der Katalog der Funktionsbausteine:
        # Schlüssel = Bausteinname, Wert = Liste der Pins (Richtung, Pin-Name, Datentyp)
        self.standardpfad = ""
        self.standardname = "ausgabe.xlsx"
        
        self.typical_daten = {
            'MODES': [{'Richtung': 'Eingang', 'Pin': 'SEL_AUT', 'Typ': 'BOOL'},
                      {'Richtung': 'Eingang', 'Pin': 'SEL_MAN', 'Typ': 'BOOL'},
                      {'Richtung': 'Eingang', 'Pin': 'SEL_LOCAL', 'Typ': 'BOOL'},
                      {'Richtung': 'Ausgang', 'Pin': 'AUT', 'Typ': 'BOOL'},
                      {'Richtung': 'Ausgang', 'Pin': 'MAN', 'Typ': 'BOOL'},
                      {'Richtung': 'Ausgang', 'Pin': 'LOCAL', 'Typ': 'BOOL'},
                      {'Richtung': 'Ausgang', 'Pin': 'FREE', 'Typ': 'BOOL'}],
            'SIM_AIN': [{'Richtung': 'Eingang', 'Pin': 'IN', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_MAX', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_MIN', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'OUT', 'Typ': 'WORD'}],
            'SIM_AOUT': [{'Richtung': 'Eingang', 'Pin': 'IN', 'Typ': 'WORD'},
                         {'Richtung': 'Eingang', 'Pin': 'MV_HIGH', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MV_LOW', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MV_MAX', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MV_MIN', 'Typ': 'REAL'},
                         {'Richtung': 'Ausgang', 'Pin': 'MV', 'Typ': 'REAL'},
                         {'Richtung': 'Ausgang', 'Pin': 'SH', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'SL', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'ERROR', 'Typ': 'BOOL'}],
            'SIM_Axis_DC': [{'Richtung': 'Eingang', 'Pin': 'MotorOut1', 'Typ': 'BOOL'},
                            {'Richtung': 'Eingang', 'Pin': 'MotorOut2', 'Typ': 'BOOL'},
                            {'Richtung': 'Eingang', 'Pin': 'ControllerOut', 'Typ': 'REAL'},
                            {'Richtung': 'Eingang', 'Pin': 'Reset', 'Typ': 'BOOL'},
                            {'Richtung': 'Ausgang', 'Pin': 'SensorIn', 'Typ': 'UDINT'}],
            'SIM_BIN': [{'Richtung': 'Eingang', 'Pin': 'SimValue', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'SimLimit', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'SH', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'SL', 'Typ': 'BOOL'}],
            'SIM_Band': [{'Richtung': 'Eingang', 'Pin': 'Out1', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'Out2', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'StartPos', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MaxPos', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MinPos', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'Pos', 'Typ': 'REAL'},
                         {'Richtung': 'Ausgang', 'Pin': 'End1', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'End2', 'Typ': 'BOOL'}],
            'SIM_CIN': [{'Richtung': 'Eingang', 'Pin': 'IN', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_MAX', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_MIN', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'OUT', 'Typ': 'UDINT'}],
            'SIM_I': [{'Richtung': 'Eingang', 'Pin': 'y', 'Typ': 'REAL'},
                      {'Richtung': 'Eingang', 'Pin': 'K_I', 'Typ': 'REAL'},
                      {'Richtung': 'Eingang', 'Pin': 'x0', 'Typ': 'REAL'},
                      {'Richtung': 'Eingang', 'Pin': 'Reset', 'Typ': 'BOOL'},
                      {'Richtung': 'Ausgang', 'Pin': 'x', 'Typ': 'REAL'}],
            'SIM_IDF1': [{'Richtung': 'Eingang', 'Pin': 'OUT', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'MV', 'Typ': 'REAL'},
                         {'Richtung': 'Ausgang', 'Pin': 'RUN', 'Typ': 'BOOL'}],
            'SIM_IDF2': [{'Richtung': 'Eingang', 'Pin': 'OUT1', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'OUT2', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'MV', 'Typ': 'REAL'},
                         {'Richtung': 'Ausgang', 'Pin': 'RUN', 'Typ': 'BOOL'}],
            'SIM_Lift': [{'Richtung': 'Eingang', 'Pin': 'Sensor', 'Typ': 'INT'},
                         {'Richtung': 'Eingang', 'Pin': 'Out1', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'MAX_POS', 'Typ': 'INT'},
                         {'Richtung': 'Eingang', 'Pin': 'Width', 'Typ': 'INT'},
                         {'Richtung': 'Eingang', 'Pin': 'OFFSET', 'Typ': 'INT'},
                         {'Richtung': 'Ausgang', 'Pin': 'POS', 'Typ': 'INT'},
                         {'Richtung': 'Ausgang', 'Pin': 'END', 'Typ': 'BOOL'}],
            'SIM_PT1': [{'Richtung': 'Eingang', 'Pin': 'y', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'K', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'T_T0', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'STOP', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'x', 'Typ': 'REAL'}],
            'SIM_PT3': [{'Richtung': 'Eingang', 'Pin': 'y', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'T1_T0', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'T2_T0', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'T3_T0', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'x', 'Typ': 'REAL'}],
            'SIM_Pulse': [{'Richtung': 'Eingang', 'Pin': 'START', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 't0', 'Typ': 'TIME'},
                          {'Richtung': 'Ausgang', 'Pin': 'OUT', 'Typ': 'BOOL'}],
            'SIM_TANK': [{'Richtung': 'Eingang', 'Pin': 'SupplyOut', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'DrainOut', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'MaxLevel', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MinLevel', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'StartLevel', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'LSHIn', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'LSLIn', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'PV', 'Typ': 'REAL'}],
            'SIM_Vessel': [{'Richtung': 'Eingang', 'Pin': 'SupplyOut', 'Typ': 'BOOL'},
                           {'Richtung': 'Eingang', 'Pin': 'DrainOut', 'Typ': 'BOOL'},
                           {'Richtung': 'Eingang', 'Pin': 'MaxLevel', 'Typ': 'REAL'},
                           {'Richtung': 'Eingang', 'Pin': 'MinLevel', 'Typ': 'REAL'},
                           {'Richtung': 'Eingang', 'Pin': 'StartLevel', 'Typ': 'REAL'},
                           {'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                           {'Richtung': 'Ausgang', 'Pin': 'SensorIn', 'Typ': 'WORD'},
                           {'Richtung': 'Ausgang', 'Pin': 'PV', 'Typ': 'REAL'}],
            'SIM_Vessel_C': [{'Richtung': 'Eingang', 'Pin': 'SupplyOut', 'Typ': 'WORD'},
                             {'Richtung': 'Eingang', 'Pin': 'DrainOut', 'Typ': 'BOOL'},
                             {'Richtung': 'Eingang', 'Pin': 'MaxLevel', 'Typ': 'REAL'},
                             {'Richtung': 'Eingang', 'Pin': 'MinLevel', 'Typ': 'REAL'},
                             {'Richtung': 'Eingang', 'Pin': 'StartLevel', 'Typ': 'REAL'},
                             {'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                             {'Richtung': 'Ausgang', 'Pin': 'SensorIn', 'Typ': 'WORD'},
                             {'Richtung': 'Ausgang', 'Pin': 'PV', 'Typ': 'REAL'}],
            'TYP_2PT': [{'Richtung': 'Eingang', 'Pin': 'SP', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'PV', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'eps', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'AUT', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'EXT', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'MV_MAN', 'Typ': 'INT'},
                        {'Richtung': 'Ausgang', 'Pin': 'MV', 'Typ': 'INT'}],
            'TYP_AIN': [{'Richtung': 'Eingang', 'Pin': 'X_W', 'Typ': 'WORD'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_HIGH', 'Typ': 'INT'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_LOW', 'Typ': 'INT'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_MAX', 'Typ': 'INT'},
                        {'Richtung': 'Eingang', 'Pin': 'PV_MIN', 'Typ': 'INT'},
                        {'Richtung': 'Ausgang', 'Pin': 'PV', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'SH', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'SL', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'ERROR', 'Typ': 'BOOL'}],
            'TYP_AOUT': [{'Richtung': 'Eingang', 'Pin': 'MV', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MV_MAX', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'MV_MIN', 'Typ': 'REAL'},
                         {'Richtung': 'Eingang', 'Pin': 'LOCK', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'Y_W', 'Typ': 'WORD'},
                         {'Richtung': 'Ausgang', 'Pin': 'ERROR', 'Typ': 'BOOL'}],
            'TYP_BIN': [{'Richtung': 'Eingang', 'Pin': 'IN', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'SO', 'Typ': 'BOOL'}],
            'TYP_CIN': [{'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'CV', 'Typ': 'DWORD'},
                        {'Richtung': 'Eingang', 'Pin': 'Reset', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'Sist', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'Vist', 'Typ': 'REAL'}],
            'TYP_D': [{'Richtung': 'Eingang', 'Pin': 'e', 'Typ': 'REAL'},
                      {'Richtung': 'Eingang', 'Pin': 'TV_T0', 'Typ': 'REAL'},
                      {'Richtung': 'Ausgang', 'Pin': 'y_D', 'Typ': 'REAL'}],
            'TYP_I': [{'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                      {'Richtung': 'Eingang', 'Pin': 'Reset', 'Typ': 'BOOL'},
                      {'Richtung': 'Eingang', 'Pin': 'e', 'Typ': 'REAL'},
                      {'Richtung': 'Eingang', 'Pin': 'TN_T0', 'Typ': 'REAL'},
                      {'Richtung': 'Ausgang', 'Pin': 'y_I', 'Typ': 'REAL'}],
            'TYP_IDF1': [{'Richtung': 'Eingang', 'Pin': 'ON', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'OFF', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'LOCK', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'OUT', 'Typ': 'BOOL'}],
            'TYP_IDF1_Light': [{'Richtung': 'Eingang', 'Pin': 'ON', 'Typ': 'BOOL'},
                               {'Richtung': 'Eingang', 'Pin': 'OFF', 'Typ': 'BOOL'},
                               {'Richtung': 'Eingang', 'Pin': 'LOCK', 'Typ': 'BOOL'},
                               {'Richtung': 'Ausgang', 'Pin': 'OUT', 'Typ': 'BOOL'}],
            'TYP_IDF2': [{'Richtung': 'Eingang', 'Pin': 'ON1', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'ON2', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'OFF', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'LOCK1', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 'LOCK2', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'OUT1', 'Typ': 'BOOL'},
                         {'Richtung': 'Ausgang', 'Pin': 'OUT2', 'Typ': 'BOOL'}],
            'TYP_IDF2_Light': [{'Richtung': 'Eingang', 'Pin': 'ON1', 'Typ': 'BOOL'},
                               {'Richtung': 'Eingang', 'Pin': 'ON2', 'Typ': 'BOOL'},
                               {'Richtung': 'Eingang', 'Pin': 'OFF', 'Typ': 'BOOL'},
                               {'Richtung': 'Eingang', 'Pin': 'LOCK1', 'Typ': 'BOOL'},
                               {'Richtung': 'Eingang', 'Pin': 'LOCK2', 'Typ': 'BOOL'},
                               {'Richtung': 'Ausgang', 'Pin': 'OUT1', 'Typ': 'BOOL'},
                               {'Richtung': 'Ausgang', 'Pin': 'OUT2', 'Typ': 'BOOL'}],
            'TYP_IPO': [{'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'qZiel', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'qStart', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'v', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'a', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'qSoll', 'Typ': 'REAL'}],
            'TYP_PID': [{'Richtung': 'Eingang', 'Pin': 'SP', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'PV', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'KP', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'TN_T0', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'TV_T0', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'MV_HIGH', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'MV_LOW', 'Typ': 'REAL'},
                        {'Richtung': 'Eingang', 'Pin': 'AUT', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'EXT', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'MV_MAN', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'MV', 'Typ': 'REAL'},
                        {'Richtung': 'Ausgang', 'Pin': 'AH', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'AL', 'Typ': 'BOOL'}],
            'TYP_PULSE': [{'Richtung': 'Eingang', 'Pin': 'IN', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 'DIR_CU', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 'DIR_CD', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 'RESET', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 'HOME', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 'MAX_POS', 'Typ': 'INT'},
                          {'Richtung': 'Ausgang', 'Pin': 'POS', 'Typ': 'INT'},
                          {'Richtung': 'Ausgang', 'Pin': 'SH', 'Typ': 'BOOL'},
                          {'Richtung': 'Ausgang', 'Pin': 'SL', 'Typ': 'BOOL'}],
            'TYP_PWM': [{'Richtung': 'Eingang', 'Pin': 'ON1', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'ON2', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'OFF', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'LOCK1', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'LOCK2', 'Typ': 'BOOL'},
                        {'Richtung': 'Eingang', 'Pin': 'TIME1', 'Typ': 'TIME'},
                        {'Richtung': 'Eingang', 'Pin': 'TIME2', 'Typ': 'TIME'},
                        {'Richtung': 'Ausgang', 'Pin': 'OUT1', 'Typ': 'BOOL'},
                        {'Richtung': 'Ausgang', 'Pin': 'CLK', 'Typ': 'BOOL'}],
            'TYP_SPEED': [{'Richtung': 'Eingang', 'Pin': 'Start', 'Typ': 'BOOL'},
                          {'Richtung': 'Eingang', 'Pin': 'Position', 'Typ': 'REAL'},
                          {'Richtung': 'Eingang', 'Pin': 'T0', 'Typ': 'REAL'},
                          {'Richtung': 'Ausgang', 'Pin': 'SPEED', 'Typ': 'REAL'}],
            'TYP_TAKT': [{'Richtung': 'Eingang', 'Pin': 'START', 'Typ': 'BOOL'},
                         {'Richtung': 'Eingang', 'Pin': 't0', 'Typ': 'TIME'},
                         {'Richtung': 'Eingang', 'Pin': 't1', 'Typ': 'TIME'},
                         {'Richtung': 'Ausgang', 'Pin': 'OUT', 'Typ': 'BOOL'}]
        }
        
        self.gesamt_basis = len(self.typical_daten)  # Anzahl der Bausteine im Basiskatalog
        self.benutzerdefinierte_fbs = set()  # Bausteine, die per XML hinzugefügt wurden (Anzeige in Blau)
        self.ungespeicherte_aenderungen = False  # True, wenn es ungespeicherte Änderungen gibt

        # Registerkarten (Reiter) anlegen
        self.registerkarten = ttk.Notebook(root)
        self.registerkarten.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.reiter_erzeugen = ttk.Frame(self.registerkarten)
        self.reiter_bausteine = ttk.Frame(self.registerkarten)
        
        self.registerkarten.add(self.reiter_erzeugen, text="Excel erzeugen")
        self.registerkarten.add(self.reiter_bausteine, text="Verfügbare Bausteine")
        
        # Inhalte der beiden Reiter aufbauen
        self.registerkarte_erzeugen_aufbauen()
        self.registerkarte_bausteine_aufbauen()

        # Fenstergröße an den Inhalt anpassen
        self.fenstergroesse_anpassen()

        # Das Schließen des Fensters abfangen (Schaltfläche "X" oder Alt+F4),
        # um bei ungespeicherten Änderungen an den Typicals nachzufragen.
        self.root.protocol("WM_DELETE_WINDOW", self.beim_fenster_schliessen)

    def beim_fenster_schliessen(self):
        """Wird beim Versuch ausgeführt, das Fenster zu schließen. Falls es
        ungespeicherte Änderungen an den Typicals gibt (hinzugefügte oder
        entfernte Bausteine), wird der Benutzer gefragt, ob er sie vor dem
        Beenden speichern möchte, um versehentlichen Datenverlust zu vermeiden."""
        if self.ungespeicherte_aenderungen:
            antwort = messagebox.askyesnocancel(
                "Ungespeicherte Änderungen",
                "Du hast Änderungen an den verfügbaren Typicals vorgenommen "
                "(Bausteine hinzugefügt oder entfernt), die noch nicht gespeichert wurden.\n\n"
                "Möchtest du die aktuelle Konfiguration vor dem Beenden speichern?"
            )

            if antwort is None:
                # Der Benutzer hat "Abbrechen" gewählt: Fenster nicht schließen.
                return
            elif antwort:
                # Der Benutzer hat "Ja" gewählt: Versuchen, die Konfiguration zu speichern.
                self.konfiguration_speichern()
                if self.ungespeicherte_aenderungen:
                    # Falls weiterhin als ungespeichert markiert (z. B. weil der
                    # Benutzer den "Speichern unter"-Dialog abgebrochen hat), wird
                    # das Fenster nicht geschlossen, um Änderungen nicht zu verlieren.
                    return
            # Wurde "Nein" gewählt, werden die Änderungen verworfen und das Fenster geschlossen.

        self.root.destroy()

    def fenstergroesse_anpassen(self):
        """Berechnet die tatsächlich benötigte Größe, um den gesamten Inhalt
        (beider Registerkarten) anzuzeigen, und passt das Fenster an diese
        Größe an, um ungenutzten weißen/grauen Leerraum zu vermeiden. Legt
        diese Größe zudem als Mindestgröße fest, erlaubt aber weiterhin,
        das Fenster zu vergrößern."""
        self.root.update_idletasks()

        # Aktiven Reiter merken
        aktiver_reiter = self.registerkarten.select()

        # Benötigte Breite und Höhe beider Reiter messen
        breiten = []
        hoehen = []
        for reiter in (self.reiter_erzeugen, self.reiter_bausteine):
            self.registerkarten.select(reiter)
            self.root.update_idletasks()
            breiten.append(self.root.winfo_reqwidth())
            hoehen.append(self.root.winfo_reqheight())

        # Ursprünglichen Reiter wieder aktivieren
        self.registerkarten.select(aktiver_reiter)
        self.root.update_idletasks()

        # Größten Wert als Fenstergröße und Mindestgröße setzen
        breite = max(breiten)
        hoehe = max(hoehen)

        self.root.geometry(f"{breite}x{hoehe}")
        self.root.minsize(breite, hoehe)

    # Baut den Reiter 'Excel erzeugen' auf (Eingabefelder und Schaltflächen)
    def registerkarte_erzeugen_aufbauen(self):
        rahmen = ttk.Frame(self.reiter_erzeugen, padding=10)
        rahmen.pack(fill='both', expand=True)
        
        # Variablen für Eingabedatei, Ausgabename und Ausgabeordner
        self.var_eingabe = tk.StringVar()
        self.var_ausgabename = tk.StringVar(value=self.standardname)
        self.var_ausgabeordner = tk.StringVar(value=self.standardpfad)
        
        # Eingabefeld für die Anweisungsdatei
        ttk.Label(rahmen, text="Anweisungsdatei (Excel):").grid(row=0, column=0, sticky='w', pady=(0, 2))
        self.eingabefeld = ttk.Entry(rahmen, textvariable=self.var_eingabe, width=65)
        self.eingabefeld.grid(row=1, column=0, sticky='w')
        ttk.Button(rahmen, text="Durchsuchen...", command=self.eingabedatei_suchen).grid(row=1, column=1, padx=(10, 0))
        
        # Eingabefeld für den Namen der Ausgabedatei
        ttk.Label(rahmen, text="Name der Ausgabe-Excel-Datei:").grid(row=2, column=0, sticky='w', pady=(15, 2))
        self.namensfeld = ttk.Entry(rahmen, textvariable=self.var_ausgabename, width=35)
        self.namensfeld.grid(row=3, column=0, sticky='w')
        
        # Eingabefeld für den Ausgabeordner
        ttk.Label(rahmen, text="Ordner zum Speichern der Excel-Datei:").grid(row=4, column=0, sticky='w', pady=(15, 2))
        self.ordnerfeld = ttk.Entry(rahmen, textvariable=self.var_ausgabeordner, width=65)
        self.ordnerfeld.grid(row=5, column=0, sticky='w')
        ttk.Button(rahmen, text="Durchsuchen...", command=self.ausgabeordner_suchen).grid(row=5, column=1, padx=(10, 0))
        
        # Schaltfläche zum Erzeugen der Excel-Datei
        schaltflaechen_rahmen = ttk.Frame(rahmen)
        schaltflaechen_rahmen.grid(row=6, column=0, columnspan=2, pady=(25, 10))
        self.schaltflaeche_erzeugen = ttk.Button(schaltflaechen_rahmen, text="Excel erzeugen", command=self.abgleich_ausfuehren)
        self.schaltflaeche_erzeugen.pack()
        
        # Trennlinie und Schaltfläche zum Speichern der Konfiguration
        trenner = ttk.Separator(rahmen, orient='horizontal')
        trenner.grid(row=7, column=0, columnspan=2, sticky='ew', pady=(10, 15))
        
        fusszeile_schaltflaechen_rahmen = ttk.Frame(rahmen)
        fusszeile_schaltflaechen_rahmen.grid(row=8, column=0, columnspan=2)
        ttk.Button(fusszeile_schaltflaechen_rahmen, text="Aktuelle Konfiguration speichern...", command=self.konfiguration_speichern).pack()
        
        ttk.Label(
            rahmen, 
            text="Legt den Zielpfad, den Dateinamen und den aktuellen Typical-Katalog fest\nund erzeugt daraus ein neues Python-Skript als Basiskonfiguration.",
            foreground="#555"
        ).grid(row=9, column=0, columnspan=2, pady=(10, 0), sticky='w')

    # Baut den Reiter 'Verfügbare Bausteine' auf
    def registerkarte_bausteine_aufbauen(self):
        # Fußzeile mit der Anzahl der Bausteine
        fusszeile_rahmen = ttk.Frame(self.reiter_bausteine, padding=(10, 5, 10, 10))
        fusszeile_rahmen.pack(side='bottom', fill='x')
        self.label_fusszeile = ttk.Label(fusszeile_rahmen, text="", foreground="#555")
        self.label_fusszeile.pack(anchor='w')

        # Kopfbereich mit Überschrift und Hinweis
        kopf_rahmen = ttk.Frame(self.reiter_bausteine, padding=(10, 10, 5, 5))
        kopf_rahmen.pack(side='top', fill='x')
        ttk.Label(kopf_rahmen, text="Verfügbare Funktionsbausteine (FB):").pack(anchor='w')
        tk.Label(kopf_rahmen, text="In Blau: Bausteine, die nicht zum Basiskatalog gehören.", fg="#0b63c5").pack(anchor='w')

        # --- SCHALTFLÄCHEN OBEN ---
        werkzeugleiste_rahmen = ttk.Frame(self.reiter_bausteine, padding=(10, 0, 10, 5))
        werkzeugleiste_rahmen.pack(side='top', fill='x')
        
        self.schaltflaeche_hinzufuegen = ttk.Button(werkzeugleiste_rahmen, text="Neuen Baustein hinzufügen", command=self.neuen_baustein_hinzufuegen)
        self.schaltflaeche_hinzufuegen.pack(side='left', padx=(0, 10))
        
        self.schaltflaeche_entfernen = ttk.Button(werkzeugleiste_rahmen, text="Baustein entfernen", command=self.baustein_entfernen)
        self.schaltflaeche_entfernen.pack(side='left')

        inhalt_rahmen = ttk.Frame(self.reiter_bausteine, padding=(10, 5, 10, 5))
        inhalt_rahmen.pack(side='top', fill='both', expand=True)
        inhalt_rahmen.columnconfigure(1, weight=1)
        inhalt_rahmen.rowconfigure(0, weight=1)

        # --- BAUSTEINLISTE (LINKS) ---
        listen_rahmen = ttk.Frame(inhalt_rahmen)
        listen_rahmen.grid(row=0, column=0, sticky='ns', padx=(0, 10))
        
        self.listenfeld_bausteine = tk.Listbox(listen_rahmen, width=25, height=20, exportselection=False)
        self.listenfeld_bausteine.pack(side='left', fill='y')
        scrollbalken_liste = ttk.Scrollbar(listen_rahmen, orient='vertical', command=self.listenfeld_bausteine.yview)
        scrollbalken_liste.pack(side='right', fill='y')
        self.listenfeld_bausteine.config(yscrollcommand=scrollbalken_liste.set)

        # --- GETEILTE PIN-TABELLEN (RECHTS) ---
        rechter_rahmen = ttk.Frame(inhalt_rahmen)
        rechter_rahmen.grid(row=0, column=1, sticky='nsew')
        
        # Spaltenüberschriften der Pin-Tabellen
        spalten = ("Pin", "Datentyp")
        
        # 1. EINGANGSPINS
        ttk.Label(rechter_rahmen, text="Eingangspins:").pack(anchor='w', pady=(0, 2))
        baum_eingang_rahmen = ttk.Frame(rechter_rahmen)
        baum_eingang_rahmen.pack(fill='both', expand=True, pady=(0, 10))
        
        self.baum_pins_eingang = ttk.Treeview(baum_eingang_rahmen, columns=spalten, show='headings')
        self.baum_pins_eingang.heading("Pin", text="Pin")
        self.baum_pins_eingang.heading("Datentyp", text="Datentyp")
        self.baum_pins_eingang.column("Pin", width=140, anchor='w')
        self.baum_pins_eingang.column("Datentyp", width=90, anchor='w')
        
        self.baum_pins_eingang.pack(side='left', fill='both', expand=True)
        scrollbalken_baum_eingang = ttk.Scrollbar(baum_eingang_rahmen, orient='vertical', command=self.baum_pins_eingang.yview)
        scrollbalken_baum_eingang.pack(side='right', fill='y')
        self.baum_pins_eingang.config(yscrollcommand=scrollbalken_baum_eingang.set)
        
        # 2. AUSGANGSPINS
        ttk.Label(rechter_rahmen, text="Ausgangspins:").pack(anchor='w', pady=(0, 2))
        baum_ausgang_rahmen = ttk.Frame(rechter_rahmen)
        baum_ausgang_rahmen.pack(fill='both', expand=True)
        
        self.baum_pins_ausgang = ttk.Treeview(baum_ausgang_rahmen, columns=spalten, show='headings')
        self.baum_pins_ausgang.heading("Pin", text="Pin")
        self.baum_pins_ausgang.heading("Datentyp", text="Datentyp")
        self.baum_pins_ausgang.column("Pin", width=140, anchor='w')
        self.baum_pins_ausgang.column("Datentyp", width=90, anchor='w')
        
        self.baum_pins_ausgang.pack(side='left', fill='both', expand=True)
        scrollbalken_baum_ausgang = ttk.Scrollbar(baum_ausgang_rahmen, orient='vertical', command=self.baum_pins_ausgang.yview)
        scrollbalken_baum_ausgang.pack(side='right', fill='y')
        self.baum_pins_ausgang.config(yscrollcommand=scrollbalken_baum_ausgang.set)
        
        # Fußzeile und Liste füllen, Auswahl in der Liste mit der Pin-Anzeige verknüpfen
        self.fusszeile_aktualisieren()
        self.bausteinliste_befuellen()
        self.listenfeld_bausteine.bind('<<ListboxSelect>>', self.pintabelle_aktualisieren)

    # Füllt die Liste alphabetisch mit allen Bausteinen (benutzerdefinierte in Blau)
    def bausteinliste_befuellen(self):
        self.listenfeld_bausteine.delete(0, tk.END)
        for idx, fb in enumerate(sorted(self.typical_daten.keys())):
            self.listenfeld_bausteine.insert(tk.END, fb)
            if fb in self.benutzerdefinierte_fbs:
                self.listenfeld_bausteine.itemconfig(idx, {'fg': '#0b63c5'})

    # Zeigt die Pins des ausgewählten Bausteins an (event = Auswahlereignis der Liste)
    def pintabelle_aktualisieren(self, event):
        # Beide Tabellen leeren
        for item in self.baum_pins_eingang.get_children():
            self.baum_pins_eingang.delete(item)
        for item in self.baum_pins_ausgang.get_children():
            self.baum_pins_ausgang.delete(item)
            
        # Ausgewählten Baustein ermitteln
        auswahl = self.listenfeld_bausteine.curselection()
        if not auswahl:
            return
            
        ausgewaehlter_baustein = self.listenfeld_bausteine.get(auswahl[0])
        pins = self.typical_daten.get(ausgewaehlter_baustein, [])
        
        # Entsprechende Tabellen befüllen
        for pin in pins:
            if pin['Richtung'] in ['Eingang', 'Ein/Aus']:
                self.baum_pins_eingang.insert('', tk.END, values=(pin['Pin'], pin['Typ']))
            if pin['Richtung'] in ['Ausgang', 'Ein/Aus']:
                self.baum_pins_ausgang.insert('', tk.END, values=(pin['Pin'], pin['Typ']))

    # Entfernt den ausgewählten Baustein vorübergehend aus dem Katalog
    def baustein_entfernen(self):
        auswahl = self.listenfeld_bausteine.curselection()
        
        if not auswahl:
            messagebox.showwarning("Achtung", "Bitte wähle einen Baustein aus der Liste aus, um ihn zu entfernen.")
            return
            
        ausgewaehlter_baustein = self.listenfeld_bausteine.get(auswahl[0])
        antwort = messagebox.askyesno("Entfernen bestätigen", f"'{ausgewaehlter_baustein}' vorübergehend entfernen?\n\n(Du kannst diese Änderung mit der Schaltfläche 'Aktuelle Konfiguration speichern' dauerhaft machen)")
        
        # Nach Bestätigung: Baustein aus Katalog und Liste löschen
        if antwort:
            if ausgewaehlter_baustein in self.typical_daten:
                del self.typical_daten[ausgewaehlter_baustein]
            if ausgewaehlter_baustein in self.benutzerdefinierte_fbs:
                self.benutzerdefinierte_fbs.remove(ausgewaehlter_baustein)
                
            self.ungespeicherte_aenderungen = True
            self.listenfeld_bausteine.delete(auswahl[0])
            
            # Beide Tabellen leeren, wenn der Baustein entfernt wird
            for item in self.baum_pins_eingang.get_children():
                self.baum_pins_eingang.delete(item)
            for item in self.baum_pins_ausgang.get_children():
                self.baum_pins_ausgang.delete(item)
                
            self.fusszeile_aktualisieren()
            
    # Aktualisiert den Text der Fußzeile
    def fusszeile_aktualisieren(self):
        self.label_fusszeile.config(text=f"{len(self.typical_daten)} Baustein(e) verfügbar ({self.gesamt_basis} aus dem ursprünglichen Basiskatalog)")

    # Dateidialog zur Auswahl der Excel-Anweisungsdatei
    def eingabedatei_suchen(self):
        datei = filedialog.askopenfilename(title="Anweisungsdatei auswählen", filetypes=[("Excel-Dateien", "*.xlsx *.xls")])
        if datei:
            self.var_eingabe.set(datei)

    # Dialog zur Auswahl des Zielordners
    def ausgabeordner_suchen(self):
        ordner = filedialog.askdirectory(title="Zielordner auswählen")
        if ordner:
            self.var_ausgabeordner.set(ordner)

    # Prüft die Eingaben und startet die Verarbeitung (Schaltfläche 'Excel erzeugen')
    def abgleich_ausfuehren(self):
        # Eingaben aus den Feldern lesen (ohne Leerzeichen am Rand)
        eingabe = self.var_eingabe.get().strip()
        ordner = self.var_ausgabeordner.get().strip()
        name = self.var_ausgabename.get().strip()
        
        # Alle Felder müssen ausgefüllt sein
        if not eingabe or not ordner or not name:
            messagebox.showwarning("Fehlende Angaben", "Bitte stelle sicher, dass alle Felder ausgefüllt sind (Anweisungsdatei, Ordner und Name).")
            return
            
        # Dateiendung .xlsx sicherstellen
        if not name.endswith('.xlsx'):
            name += '.xlsx'
            
        # Vollständigen Ausgabepfad zusammensetzen
        vollstaendiger_pfad = os.path.join(ordner, name)

        # Schaltfläche sperren, Verarbeitung starten und Ergebnis melden
        try:
            self.schaltflaeche_erzeugen.config(state="disabled")
            nachricht = daten_verarbeiten(eingabe, vollstaendiger_pfad, self.typical_daten)
            messagebox.showinfo("Vorgang abgeschlossen", nachricht)
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
        finally:
            # Schaltfläche wieder freigeben
            self.schaltflaeche_erzeugen.config(state="normal")

    # Liest Funktionsbausteine (POUs) aus einer XML-Datei und fügt sie dem Katalog hinzu
    def neuen_baustein_hinzufuegen(self):
        xml_pfad = filedialog.askopenfilename(title="XML-Katalog auswählen", filetypes=[("XML-Dateien", "*.xml")])
        if not xml_pfad or not os.path.exists(xml_pfad):
            return 
            
        try:
            tree = ET.parse(xml_pfad)  # Struktur der XML-Datei einlesen
            root = tree.getroot()  # Wurzelelement der XML-Datei
            
            # Hilfsfunktion: entfernt den Namensraum (Namespace) aus einem XML-Tag
            def ns_entfernen(tag):
                return tag.split('}', 1)[1] if '}' in tag else tag
                
            # neue_bausteine: Bausteinname -> Liste der Pins
            neue_bausteine = {}
            # Alle POUs (Bausteine) in der XML-Datei durchsuchen
            for pou in root.iter():
                if ns_entfernen(pou.tag) == 'pou':
                    baustein_name = pou.get('name')
                    pins = []
                    
                    # Schnittstelle (interface) mit den Variablengruppen suchen
                    for interface in pou.iter():
                        if ns_entfernen(interface.tag) == 'interface':
                            for var_gruppe in interface:
                                gruppen_tag = ns_entfernen(var_gruppe.tag)
                                # Nur Eingänge, Ausgänge und Ein-/Ausgänge berücksichtigen
                                if gruppen_tag in ['inputVars', 'outputVars', 'inOutVars']:
                                    # Richtung anhand des Gruppennamens bestimmen
                                    richtung = "Eingang" if gruppen_tag == 'inputVars' else "Ausgang" if gruppen_tag == 'outputVars' else "Ein/Aus"
                                    
                                    # Jede Variable ist ein Pin (Name und Datentyp auslesen)
                                    for variable_node in var_gruppe.iter():
                                        if ns_entfernen(variable_node.tag) == 'variable':
                                            pin_name = variable_node.get('name')
                                            pin_typ = "UNKNOWN"
                                            
                                            # Datentyp bestimmen ('derived' = selbst definierter Typ)
                                            for typ_node in variable_node.iter():
                                                if ns_entfernen(typ_node.tag) == 'type':
                                                    for kind in typ_node:
                                                        k_tag = ns_entfernen(kind.tag)
                                                        if k_tag == 'derived':
                                                            pin_typ = kind.get('name', 'UNKNOWN')
                                                        else:
                                                            pin_typ = k_tag.upper()
                                                            
                                            # Pin zur Liste hinzufügen
                                            pins.append({'Pin': pin_name, 'Richtung': richtung, 'Typ': pin_typ})
                    # Baustein nur übernehmen, wenn er einen Namen und Pins hat
                    if baustein_name and pins:
                        neue_bausteine[baustein_name] = pins
                        
            # Meldung, falls nichts gefunden wurde
            if not neue_bausteine:
                messagebox.showinfo("Keine Ergebnisse", "In der XML-Datei wurden keine Funktionsbausteine (POUs mit Variablen) gefunden.")
                return
            
            # Nach Bestätigung die Bausteine in den Katalog übernehmen
            if messagebox.askyesno("Bestätigen", "Bist du sicher, dass du diesen Baustein hinzufügen möchtest?"):
                for bausteinname, pins_liste in neue_bausteine.items():
                    self.typical_daten[bausteinname] = pins_liste
                    self.benutzerdefinierte_fbs.add(bausteinname)
                
                # Änderung als ungespeichert markieren und Anzeige aktualisieren
                self.ungespeicherte_aenderungen = True
                self.bausteinliste_befuellen()
                self.fusszeile_aktualisieren()
                messagebox.showinfo("Erfolg", "Der Baustein wurde erfolgreich hinzugefügt. Er erscheint blau in der Liste.")
                
        # Fehlerbehandlung: ungültiges XML bzw. sonstige Lesefehler
        except ET.ParseError:
            messagebox.showerror("XML-Fehler", "Die Datei hat kein gültiges XML-Format.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Beim Lesen der XML-Datei ist ein Fehler aufgetreten:\n{str(e)}")

    # Speichert eine Kopie dieses Skripts mit aktuellem Katalog und aktuellen Standardpfaden
    def konfiguration_speichern(self):
        # Zieldatei über einen Dialog auswählen
        zieldatei = filedialog.asksaveasfilename(
            title="Skript mit aktueller Konfiguration speichern",
            defaultextension=".py",
            initialfile="Benutzerdefinierter_Generator.py",
            filetypes=[("Python-Dateien", "*.py")]
        )
        if not zieldatei:
            return

        try:
            # Eigenen Quelltext einlesen (inhalt = kompletter Text des Skripts)
            aktueller_pfad = os.path.abspath(__file__)
            with open(aktueller_pfad, 'r', encoding='utf-8') as f:
                inhalt = f.read()

            # Bereich der Standardpfade im Quelltext suchen und durch die aktuellen Werte ersetzen
            str_start_r = "self.standardpfad ="
            str_ende_r = "self.typical_daten ="
            idx_start_r = inhalt.find(str_start_r)
            idx_ende_r = inhalt.find(str_ende_r)
            
            if idx_start_r != -1 and idx_ende_r != -1:
                praefix_r = inhalt[:idx_start_r]
                suffix_r = inhalt[idx_ende_r:]
                aktueller_ordner = self.var_ausgabeordner.get().strip().replace('\\', '\\\\')
                aktueller_name = self.var_ausgabename.get().strip()
                neuer_pfade_block = f"self.standardpfad = '{aktueller_ordner}'\n        self.standardname = '{aktueller_name}'\n        \n        "
                inhalt = praefix_r + neuer_pfade_block + suffix_r
            else:
                raise Exception("Die Konfigurationsvariablen für die Pfade wurden im Quellcode nicht gefunden.")

            # Bereich des Katalogs im Quelltext suchen und durch den aktuellen Katalog ersetzen
            str_start_katalog = "self.typical_daten = {"
            str_ende_katalog = "self.gesamt_basis = len(self.typical_daten)"
            idx_start = inhalt.find(str_start_katalog)
            idx_ende = inhalt.find(str_ende_katalog)

            if idx_start != -1 and idx_ende != -1:
                praefix = inhalt[:idx_start]
                suffix = inhalt[idx_ende:]
                
                # Katalog als lesbaren Text formatieren und einrücken
                dict_formatiert = pprint.pformat(self.typical_daten, indent=4)
                dict_eingerueckt = "\n".join(["        " + line for line in dict_formatiert.split('\n')])
                neuer_katalog_block = f"self.typical_daten = {dict_eingerueckt.lstrip()}\n        \n        "
                
                inhalt = praefix + neuer_katalog_block + suffix
            else:
                raise Exception("Die Hauptvariable des Katalogs wurde im Quellcode nicht gefunden.")

            # Neues Skript schreiben
            with open(zieldatei, 'w', encoding='utf-8') as f:
                f.write(inhalt)

            # Änderungen gelten jetzt als gespeichert
            self.ungespeicherte_aenderungen = False

            messagebox.showinfo(
                "Konfiguration gespeichert", 
                f"Erfolg! Es wurde eine neue Version des Programms erstellt unter:\n\n{zieldatei}\n\nWenn du diese neue Datei öffnest, wird deine aktuelle Typical-Konfiguration (einschließlich der blau hinzugefügten Bausteine) zum neuen Standardkatalog in Schwarz."
            )
            
        # Fehler beim Speichern melden
        except Exception as e:
            messagebox.showerror("Fehler beim Speichern", f"Beim Erstellen der neuen Datei ist ein Fehler aufgetreten:\n\n{str(e)}")

# Programmstart: Hauptfenster erzeugen und Ereignisschleife starten
if __name__ == '__main__':
    root = tk.Tk()
    app = AppDatenabgleich(root)
    root.mainloop()