# ==============================================================================
# Programm:      Skriptgenerator für HMI (CODESYS)
# Autor:         Ali Duarte Perez Luna
# Datum:         29.09.2026
# Beschreibung:  Liest eine Excel-Anweisungsdatei ein und erzeugt daraus ein
#                CODESYS-Python-Skript, das eine HMI-Visualisierung mit Rahmen
#                und einer Start-Schaltfläche im Projekt anlegt.
#                In der Excel-Datei können nur die Koordinaten der Rahmen
#                festgelegt werden (vorletzte Spalte = X, letzte Spalte = Y).
#                Die Größe der Rahmen ist fest (Standardgröße der Bibliothek).
# ==============================================================================
import sys  # Systemfunktionen (Python-Laufzeitumgebung)
import os  # Dateipfade und Ordner
import re  # Reguläre Ausdrücke (Suchen/Ersetzen im Quelltext)
import threading  # Hintergrund-Thread für die Skripterzeugung
import queue  # Datenaustausch zwischen Thread und GUI
import pandas as pd  # Einlesen der Excel-Datei
import tkinter as tk  # Grundbibliothek der grafischen Oberfläche
from tkinter import filedialog, messagebox  # Dateidialoge und Meldungsfenster
from tkinter import ttk  # Moderne Widgets (Buttons, Tabellen, Registerkarten)

# Ursprüngliche Rahmenbibliothek mit Variablen, Größen (w, h) und Richtung (INPUT / IN_OUT)
# Aufbau: Rahmenname -> 'w' = Breite, 'h' = Höhe (in Pixel),
#         'vars' = Variablenliste mit 'name', 'type' (Datentyp) und 'dir' (Richtung)
RAHMEN_BIBLIOTHEK = {
    'FPL_2PT': {'w': 180, 'h': 190, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'Contr', 'type': 'TYP_2PT', 'dir': 'IN_OUT'}]},
    'FPL_Motor': {'w': 115, 'h': 235, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'FPL_Motor2': {'w': 120, 'h': 300, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}]},
    'FPL_Motor2_Light': {'w': 110, 'h': 250, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}]},
    'FPL_Motor_Light': {'w': 80, 'h': 200, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'FPL_PID': {'w': 185, 'h': 295, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'Contr', 'type': 'TYP_PID', 'dir': 'IN_OUT'}]},
    'FPL_SFC': {'w': 395, 'h': 475, 'vars': []},
    'FPL_Valve': {'w': 110, 'h': 240, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'FPL_Valve_Light': {'w': 80, 'h': 200, 'vars': [{'name': 'Name', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'SBL_AIN': {'w': 230, 'h': 70, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'Unit', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'AIN', 'type': 'TYP_AIN', 'dir': 'INPUT'}]},
    'SBL_BIN': {'w': 80, 'h': 30, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'BIN', 'type': 'TYP_BIN', 'dir': 'INPUT'}]},
    'SBL_BIN_L': {'w': 80, 'h': 30, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'BIN', 'type': 'TYP_BIN', 'dir': 'INPUT'}]},
    'SBL_ControlValve': {'w': 100, 'h': 80, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'AOUT', 'type': 'TYP_AOUT', 'dir': 'INPUT'}]},
    'SBL_IDF2': {'w': 100, 'h': 65, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF2', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'SBL_Level': {'w': 120, 'h': 140, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'AIN', 'type': 'TYP_AIN', 'dir': 'INPUT'}]},
    'SBL_Mixer': {'w': 55, 'h': 250, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'SBL_Pump': {'w': 75, 'h': 85, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'SBL_Pump_H': {'w': 115, 'h': 65, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'SBL_Valve': {'w': 85, 'h': 60, 'vars': [{'name': 'Text', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'IDF', 'type': 'TYP_IDF1', 'dir': 'IN_OUT'}]},
    'VIS_Arm2': {'w': 105, 'h': 255, 'vars': [{'name': 'Arm', 'type': 'AXES_GROUP_REF', 'dir': 'IN_OUT'}]},
    'VIS_Axis_Lin': {'w': 430, 'h': 250, 'vars': [{'name': 'Axis', 'type': 'TYP_Axis', 'dir': 'IN_OUT'}]},
    'VIS_Axis_Rot': {'w': 170, 'h': 170, 'vars': [{'name': 'Axis', 'type': 'TYP_Axis', 'dir': 'IN_OUT'}]},
    'VIS_Band': {'w': 280, 'h': 295, 'vars': [{'name': 'Motor', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'End1', 'type': 'TYP_BIN', 'dir': 'INPUT'}, {'name': 'End2', 'type': 'TYP_BIN', 'dir': 'INPUT'}, {'name': 'Position', 'type': 'REAL', 'dir': 'INPUT'}, {'name': 'TxtMotor', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'TxtEnd2', 'type': 'STRING', 'dir': 'INPUT'}, {'name': 'TxtEnd1', 'type': 'STRING', 'dir': 'INPUT'}]},
    'VIS_Lift_X': {'w': 670, 'h': 330, 'vars': [{'name': 'Dist_X', 'type': 'INT', 'dir': 'INPUT'}, {'name': 'Dist_Y', 'type': 'INT', 'dir': 'INPUT'}, {'name': 'Motor', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'Sensor', 'type': 'TYP_PULSE', 'dir': 'INPUT'}, {'name': 'Switch', 'type': 'TYP_BIN', 'dir': 'INPUT'}, {'name': 'Motor_Y', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'Sensor_Y', 'type': 'TYP_PULSE', 'dir': 'INPUT'}, {'name': 'Switch_Y', 'type': 'TYP_BIN', 'dir': 'INPUT'}, {'name': 'Motor_Z', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'Sensor_Z', 'type': 'TYP_PULSE', 'dir': 'INPUT'}, {'name': 'Switch_Z', 'type': 'TYP_BIN', 'dir': 'INPUT'}]},
    'VIS_Lift_Y': {'w': 465, 'h': 315, 'vars': [{'name': 'Dist_Y', 'type': 'INT', 'dir': 'INPUT'}, {'name': 'Motor', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'Sensor', 'type': 'TYP_PULSE', 'dir': 'INPUT'}, {'name': 'Switch', 'type': 'TYP_BIN', 'dir': 'INPUT'}, {'name': 'Motor_Z', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'Sensor_Z', 'type': 'TYP_PULSE', 'dir': 'INPUT'}, {'name': 'Switch_Z', 'type': 'TYP_BIN', 'dir': 'INPUT'}]},
    'VIS_Lift_Z': {'w': 280, 'h': 70, 'vars': [{'name': 'Motor', 'type': 'TYP_IDF2', 'dir': 'IN_OUT'}, {'name': 'Sensor', 'type': 'TYP_PULSE', 'dir': 'INPUT'}, {'name': 'Switch', 'type': 'TYP_BIN', 'dir': 'INPUT'}]},
    'VIS_TANK_B': {'w': 105, 'h': 140, 'vars': [{'name': 'Level', 'type': 'REAL', 'dir': 'INPUT'}]},
    'VIS_TANK_G': {'w': 105, 'h': 140, 'vars': [{'name': 'Level', 'type': 'REAL', 'dir': 'INPUT'}]},
    'VIS_TANK_R': {'w': 105, 'h': 140, 'vars': [{'name': 'Level', 'type': 'REAL', 'dir': 'INPUT'}]},
    'VIS_Vessel': {'w': 105, 'h': 140, 'vars': [{'name': 'AIN', 'type': 'TYP_AIN', 'dir': 'INPUT'}]},
}

# Temporäre Variablen für die aktuelle Sitzung (werden beim Schließen gelöscht)
BENUTZERDEFINIERTE_RAHMEN = {}  # Vom Benutzer neu deklarierte Rahmen {name: {'w','h','vars'}}
ENTFERNTE_RAHMEN = set()  # Namen der in dieser Sitzung entfernten Rahmen
GEAENDERTE_GROESSEN = {}  # {rahmen_name: (breite, hoehe)} -> manuell angepasste Standardgrößen

# Liefert die aktuell gültige Rahmenbibliothek (Original + Änderungen der Sitzung)
def alle_rahmen_abrufen():
    """Kombiniert die ursprüngliche Bibliothek mit den benutzerdefinierten Rahmen, lässt die entfernten
    Rahmen aus und wendet zuletzt eventuell manuell angepasste Standardgrößen an."""
    # Bibliothek und benutzerdefinierte Rahmen zusammenführen
    alle = {**RAHMEN_BIBLIOTHEK, **BENUTZERDEFINIERTE_RAHMEN}
    # Entfernte Rahmen aus dem Ergebnis löschen
    for r in ENTFERNTE_RAHMEN:
        if r in alle:
            del alle[r]
    # Manuell geänderte Standardgrößen übernehmen
    for rahmen_name, (breite, hoehe) in GEAENDERTE_GROESSEN.items():
        if rahmen_name in alle:
            alle[rahmen_name] = {**alle[rahmen_name], 'w': breite, 'h': hoehe}
    return alle

# Parameter: rahmenliste = Rahmen aus Excel, ausgabedatei = Zielpfad,
#            hmi_name = Name der Visualisierung, bibliotheksname = Name der Rahmenbibliothek
def codesys_skript_erzeugen(rahmenliste, ausgabedatei, hmi_name, bibliotheksname):
    """Erzeugt das CODESYS-Skript und berechnet Kollisionen, um die Rahmen und die Start-Schaltfläche zu platzieren."""
    
    # skript_inhalt: Text des CODESYS-Skripts (Kopf: Projekt öffnen, Visualisierung anlegen)
    skript_inhalt = f"""# -*- coding: utf-8 -*-
from scriptengine import*
import os

scriptfile = os.path.dirname(__file__)
projectfile = os.path.join(scriptfile, "automatische Projektgenerierung")

if projects.primary:
    prj = projects.primary
else:
    prj = projects.open(projectfile)

apl = prj.active_application
visu = apl.create_visualobject("{hmi_name}")

visu.begin_modify()
elementlist = visu.visual_element_list
"""

    # Hilfsvariablen der automatischen Platzierung:
    # platzierte_rechtecke = belegte Flächen (x, y, Breite, Höhe)
    # abstand = Mindestabstand in Pixel, maximale_bildschirmbreite = Grenze für den Zeilenumbruch
    platzierte_rechtecke = []
    abstand = 20
    maximale_bildschirmbreite = 1000

    for r in rahmenliste:
        # Manuell positionierte Rahmen zuerst als belegte Fläche eintragen
        if r.get('has_manual_pos'):
            platzierte_rechtecke.append((r['x'], r['y'], r['w'], r['h']))

    # Startposition (aktuelle_x / aktuelle_y) für die automatische Platzierung
    aktuelle_x = 20
    aktuelle_y = 20
    
    for r in rahmenliste:
        # Rahmen ohne manuelle Position: freien Platz suchen (b = Breite, h = Höhe)
        if not r.get('has_manual_pos'):
            b = r['w']
            h = r['h']
            
            while True:
                # Zeilenumbruch, wenn der Rahmen über den rechten Rand hinausragen würde
                if aktuelle_x + b > maximale_bildschirmbreite and aktuelle_x > 20:
                    aktuelle_x = 20
                    aktuelle_y += 20 
                
                # Kollisionsprüfung: überlappt der Rahmen ein bereits platziertes Rechteck (inkl. Abstand)
                kollision = False
                for px, py, pb, ph in platzierte_rechtecke:
                    if (aktuelle_x < px + pb + abstand and
                        aktuelle_x + b + abstand > px and
                        aktuelle_y < py + ph + abstand and
                        aktuelle_y + h + abstand > py):
                        kollision = True
                        aktuelle_x = px + pb + abstand
                        break
                        
                # Freie Stelle gefunden: Position speichern und Fläche als belegt markieren
                if not kollision:
                    r['x'] = aktuelle_x
                    r['y'] = aktuelle_y
                    platzierte_rechtecke.append((aktuelle_x, aktuelle_y, b, h))
                    aktuelle_x += b + abstand
                    break

    # Für jeden Rahmen den CODESYS-Code erzeugen
    for i, rahmen_daten in enumerate(rahmenliste):
        # Daten des Rahmens: Name, Variablen, Position (pos_x, pos_y) und Größe (breite, hoehe)
        rahmen_name = rahmen_daten['frame_name']
        variablen = rahmen_daten['variables']
        pos_x = rahmen_daten['x']
        pos_y = rahmen_daten['y']
        breite = rahmen_daten['w']
        hoehe = rahmen_daten['h']
        
        # Eindeutige Variablennamen im erzeugten Skript (Referenz und Rahmen-Objekt)
        ref_var = f"rahmen_ref_{i}"
        rahmen_var = f"rahmen_obj_{i}"
        
        skript_inhalt += f"""
# --- Rahmen {i+1} wird hinzugefügt: {rahmen_name} ---
{ref_var} = visu.create_frame_reference("{bibliotheksname}.{rahmen_name}")
{rahmen_var} = elementlist.add_element(VisualElementType.Frame)
{rahmen_var}.set_frame_references([{ref_var}])

{rahmen_var}.set_property("Position.X", {int(pos_x)})
{rahmen_var}.set_property("Position.Y", {int(pos_y)})
{rahmen_var}.set_property("Position.Breite", {int(breite)})
{rahmen_var}.set_property("Position.Höhe", {int(hoehe)})
"""
        # Variablenwerte des Rahmens im erzeugten Skript zuweisen
        for variable_info in variablen:
            v_name = variable_info['name']
            v_wert = variable_info['value']
            
            # STRING-Werte müssen in einfachen Anführungszeichen stehen
            if "STRING" in variable_info['type'].upper() and not v_wert.startswith("'"):
                v_wert = f"'{v_wert}'"
                
            skript_inhalt += f"""
{rahmen_var}.set_property(
    "Referenzierte Visualisierungen.{bibliotheksname}.{rahmen_name}.{v_name}",
    "{v_wert}")
"""

    # Größe der Start-Schaltfläche (Ellipse) und Suche nach einem freien Platz
    start_breite = 50
    start_hoehe = 50
    
    while True:
        if aktuelle_x + start_breite > maximale_bildschirmbreite and aktuelle_x > 20:
            aktuelle_x = 20
            aktuelle_y += 20 
        
        # Gleiche Kollisionsprüfung für die Start-Schaltfläche
        kollision = False
        for px, py, pb, ph in platzierte_rechtecke:
            if (aktuelle_x < px + pb + abstand and
                aktuelle_x + start_breite + abstand > px and
                aktuelle_y < py + ph + abstand and
                aktuelle_y + start_hoehe + abstand > py):
                kollision = True
                aktuelle_x = px + pb + abstand
                break
                
        if not kollision:
            break

    skript_inhalt += f"""
# --- Start-Schaltfläche wird hinzugefügt ---
start_button = elementlist.add_element(VisualElementType.Ellipse)
start_button.set_property("Position.X", {int(aktuelle_x)})
start_button.set_property("Position.Y", {int(aktuelle_y)})
start_button.set_property("Position.Breite", {start_breite})
start_button.set_property("Position.Höhe", {start_hoehe})
start_button.set_property("Farben.Normalzustand.Rahmenfarbe.Transparenz", 255)
start_button.set_property("Farben.Normalzustand.Füllfarbe.Transparenz", 255)
start_button.set_property("Farben.Alarmzustand.Rahmenfarbe.Transparenz", 255)
start_button.set_property("Farben.Alarmzustand.Füllfarbe.Transparenz", 255)
start_button.set_property("Aussehen.Linienstärke", 1)
start_button.set_property("Aussehen.Füllart", "Ausgefüllt")
start_button.set_property("Aussehen.Linienart", "Durchgezogen")
start_button.set_property("Texte.Text", "Start")
start_button.set_property("Farbvariablen.Farbumschlag", "Ablauf.Start")
start_button.set_property("Eingabekonfiguration.Tasten.Variable", "Ablauf.Start")

visu.end_modify()

print("HMI erfolgreich mit mehreren Rahmen und Start-Schaltfläche im Projekt erzeugt.")
print("Hinweis: das Projekt wurde NICHT automatisch gespeichert.")
"""
    
    # Erzeugtes Skript in die Ausgabedatei schreiben (UTF-8)
    with open(ausgabedatei, "w", encoding="utf-8") as datei:
        datei.write(skript_inhalt)

# Wandelt einen Zellenwert in float um; gibt None zurück bei leerem Wert, 'nan' oder ungültiger Zahl
def versuche_float_zu_parsen(wert):
    try:
        # Wert als Text bereinigen (Leerzeichen entfernen, Kleinbuchstaben)
        wert_str = str(wert).strip().lower()
        if wert_str == 'nan' or wert_str == '':
            return None
        return float(wert)
    except ValueError:
        return None

# Liest die Excel-Datei zeilenweise: Spalte 1 = Rahmenname, danach die Variablenwerte.
# Die letzten beiden Spalten der Tabelle enthalten die Koordinaten:
# vorletzte Spalte = X, letzte Spalte = Y (beide optional, aber nur gemeinsam gültig).
# Die Größe der Rahmen lässt sich in der Excel-Datei NICHT ändern: es gilt immer die Standardgröße.
# Rückgabe: Rahmenliste sowie Listen mit Fehlern/Warnungen
def excel_anweisungen_lesen(excel_pfad):
    # df: Tabelle (DataFrame) mit dem Inhalt der Excel-Datei
    df = pd.read_excel(excel_pfad)
    # Ergebnisliste und Fehlerlisten (werden später in der GUI ausgewertet)
    rahmenliste = []
    fehlende_warnungen = []
    nicht_gefundene_rahmen = []
    fehlende_rahmennamen = []
    unvollstaendige_position = []
    alle_rahmen = alle_rahmen_abrufen()
    
    # Lage der Koordinatenspalten: die letzten beiden Spalten der Tabelle (vorletzte = X, letzte = Y).
    # Mindestens 3 Spalten nötig (Rahmenname + X + Y); sonst gibt es keine Koordinatenspalten.
    anzahl_spalten = len(df.columns)
    hat_koordinatenspalten = anzahl_spalten >= 3
    x_spalte = anzahl_spalten - 2
    y_spalte = anzahl_spalten - 1
    # Variablenwerte dürfen nur VOR den Koordinatenspalten stehen
    grenze_variablen = x_spalte if hat_koordinatenspalten else anzahl_spalten
    
    # Jede Zeile verarbeiten (index + 2 = Zeilennummer in Excel wegen Kopfzeile)
    for index, zeile in df.iterrows():
        # Erste Spalte = Name des Rahmens
        rahmen_name = str(zeile.iloc[0]).strip()
        if rahmen_name == 'nan' or not rahmen_name:
            # Prüfen, ob in den restlichen Spalten dieser Zeile bereits Daten eingetragen wurden,
            # obwohl der Name des Rahmens (erste Spalte) fehlt. In diesem Fall handelt es sich
            # nicht um eine leere Zeile, sondern um einen Fehler in der Excel-Datei.
            restliche_werte = zeile.iloc[1:]
            hat_weitere_daten = any(
                str(wert).strip() != '' and str(wert).strip().lower() != 'nan'
                for wert in restliche_werte
            )
            if hat_weitere_daten:
                fehlende_rahmennamen.append(
                    f"Zeile {index + 2} -> der Name des Rahmens (erste Spalte) fehlt, "
                    "obwohl in dieser Zeile weitere Daten eingetragen wurden."
                )
            continue
            
        # Unbekannten Rahmen melden und Zeile überspringen
        if rahmen_name not in alle_rahmen:
            nicht_gefundene_rahmen.append(f"Zeile {index + 2} -> '{rahmen_name}' (Rahmen existiert nicht in der Bibliothek)")
            continue
            
        # Benötigte Variablen und Standardgröße aus der Bibliothek holen
        benoetigte_variablen = alle_rahmen[rahmen_name]['vars']
        standard_breite = alle_rahmen[rahmen_name]['w']
        standard_hoehe = alle_rahmen[rahmen_name]['h']
        eingegebene_variablen = []
        fehlende_variablen_rahmen = []
        
        # Variablenwerte aus den folgenden Spalten lesen (leer = fehlende Variable).
        # Die beiden Koordinatenspalten am Tabellenende gehören nie zu den Variablen.
        for i, var in enumerate(benoetigte_variablen):
            if (i + 1) < grenze_variablen:
                wert = str(zeile.iloc[i + 1]).strip()
                if wert == 'nan':
                    wert = "" 
            else:
                wert = ""
                
            if wert == "":
                fehlende_variablen_rahmen.append(var['name'])
                
            eingegebene_variablen.append({'name': var['name'], 'type': var['type'], 'value': wert})
            
        if len(fehlende_variablen_rahmen) > 0:
            fehlende_warnungen.append(f"Zeile {index + 2} -> {rahmen_name} (Fehlende Variable: {', '.join(fehlende_variablen_rahmen)})")
            
        # Koordinaten aus den letzten beiden Spalten der Tabelle lesen (vorletzte = X, letzte = Y)
        manuelle_x = versuche_float_zu_parsen(zeile.iloc[x_spalte]) if hat_koordinatenspalten else None
        manuelle_y = versuche_float_zu_parsen(zeile.iloc[y_spalte]) if hat_koordinatenspalten else None
        
        # Für die Position müssen BEIDE Werte (X und Y) angegeben werden. Wird nur einer der
        # beiden Werte eingetragen, ist die Angabe unvollständig und das Skript darf nicht erzeugt werden.
        position_teilweise_angegeben = (manuelle_x is not None) != (manuelle_y is not None)
        
        if position_teilweise_angegeben:
            unvollstaendige_position.append(
                f"Zeile {index + 2} -> {rahmen_name} (Für die Position müssen sowohl X als auch Y angegeben werden. "
                "Trage beide Werte ein oder lösche beide Werte, damit die automatische Platzierung verwendet wird.)"
            )
        
        # Manuelle Position gilt nur, wenn X und Y beide angegeben sind
        hat_manuelle_position = manuelle_x is not None and manuelle_y is not None
        
        # Rahmen speichern; die Größe ist immer die Standardgröße aus der Bibliothek
        rahmenliste.append({
            'frame_name': rahmen_name,
            'variables': eingegebene_variablen,
            'has_manual_pos': hat_manuelle_position,
            'x': manuelle_x if hat_manuelle_position else 0,
            'y': manuelle_y if hat_manuelle_position else 0,
            'w': standard_breite,
            'h': standard_hoehe
        })
        
    return rahmenliste, fehlende_warnungen, nicht_gefundene_rahmen, fehlende_rahmennamen, unvollstaendige_position


# --- Funktionen der grafischen Benutzeroberfläche (GUI) ---

# Öffnet einen Dialog zur Auswahl der Excel-Datei und trägt den Pfad ein
def datei_suchen():
    pfad = filedialog.askopenfilename(
        title="Anweisungsdatei auswählen",
        filetypes=[("Excel-Dateien", "*.xlsx *.xls")]
    )
    if pfad:
        eingabe_excel.delete(0, tk.END)
        eingabe_excel.insert(0, pfad)

# Öffnet einen Dialog zur Auswahl des Zielordners und trägt den Pfad ein
def ausgabeordner_suchen():
    pfad = filedialog.askdirectory(
        title="Ordner zum Speichern der erzeugten Datei auswählen"
    )
    if pfad:
        eingabe_ausgabeordner.delete(0, tk.END)
        eingabe_ausgabeordner.insert(0, pfad)

# Aktiviert/deaktiviert das Feld für den Bibliotheksnamen je nach Checkbox
def bibliothek_umschalten():
    if var_check_bibliothek.get():
        eingabe_bibliothek.config(state=tk.NORMAL)
    else:
        eingabe_bibliothek.config(state=tk.DISABLED)

# Wird vom Button 'Skript erzeugen' aufgerufen: prüft die Eingaben und startet die Erzeugung
def generierung_ausfuehren():
    # Eingaben aus den Feldern lesen (excel_pfad, hmi_name, ausgabedatei, bibliotheksname, ausgabeordner)
    excel_pfad = eingabe_excel.get().strip()
    hmi_name = eingabe_hmi.get().strip()
    ausgabedatei = eingabe_ausgabe.get().strip()
    bibliotheksname = eingabe_bibliothek.get().strip()
    ausgabeordner = eingabe_ausgabeordner.get().strip()

    # Prüfen, ob alle Felder ausgefüllt sind
    if not excel_pfad or not hmi_name or not ausgabedatei or not bibliotheksname or not ausgabeordner:
        messagebox.showwarning("Unvollständige Felder", "Bitte fülle alle Felder aus.")
        return

    # Dateiendung .py sicherstellen
    if not ausgabedatei.lower().endswith(".py"):
        ausgabedatei += ".py"

    # Prüfen, ob der Zielordner existiert
    if not os.path.isdir(ausgabeordner):
        messagebox.showerror("Fehler", "Der ausgewählte Zielordner existiert nicht oder ist ungültig.")
        return

    # Vollständigen Ausgabepfad bilden
    ausgabedatei = os.path.join(ausgabeordner, ausgabedatei)

    # --- Popup-Fenster „Wird erzeugt...“ ---
    fortschritts_fenster = tk.Toplevel(hauptfenster)
    fortschritts_fenster.title("Verarbeitung")
    fortschritts_fenster.resizable(False, False)
    fortschritts_fenster.transient(hauptfenster)
    fortschritts_fenster.protocol("WM_DELETE_WINDOW", lambda: None)  # Verhindert manuelles Schließen

    fortschritts_rahmen = ttk.Frame(fortschritts_fenster, padding=20)
    fortschritts_rahmen.pack(fill='both', expand=True)

    ttk.Label(fortschritts_rahmen, text="Datei wird erzeugt, bitte warten...").pack(pady=(0, 10))
    # Unbestimmter Fortschrittsbalken (läuft, bis die Erzeugung fertig ist)
    fortschrittsbalken = ttk.Progressbar(fortschritts_rahmen, mode="indeterminate", length=280)
    fortschrittsbalken.pack()
    fortschrittsbalken.start(10)

    fortschritts_fenster.update_idletasks()
    # Popup-Fenster über dem Hauptfenster zentrieren
    x = hauptfenster.winfo_x() + (hauptfenster.winfo_width() // 2) - (fortschritts_fenster.winfo_width() // 2)
    y = hauptfenster.winfo_y() + (hauptfenster.winfo_height() // 2) - (fortschritts_fenster.winfo_height() // 2)
    fortschritts_fenster.geometry(f"+{x}+{y}")
    fortschritts_fenster.grab_set()

    # Warteschlange: Der Thread übergibt hier Ergebnistyp und Daten an die GUI
    ergebnis_warteschlange = queue.Queue()

    def hintergrundarbeit():
        """Wird in einem separaten Thread ausgeführt, um das Fortschrittsfenster nicht einzufrieren."""
        try:
            # Excel einlesen; bei Fehlern wird nur eine Meldung übergeben und keine Datei erzeugt
            zu_erstellende_rahmen, fehlende_warnungen, nicht_gefundene_rahmen, fehlende_rahmennamen, unvollstaendige_position = excel_anweisungen_lesen(excel_pfad)

            if fehlende_rahmennamen:
                ergebnis_warteschlange.put(('rahmenname_fehlt', fehlende_rahmennamen))
                return

            if unvollstaendige_position:
                ergebnis_warteschlange.put(('unvollstaendige_position', unvollstaendige_position))
                return

            if nicht_gefundene_rahmen:
                ergebnis_warteschlange.put(('rahmen_nicht_gefunden', nicht_gefundene_rahmen))
                return

            if fehlende_warnungen:
                ergebnis_warteschlange.put(('fehlende_warnungen', fehlende_warnungen))
                return

            codesys_skript_erzeugen(zu_erstellende_rahmen, ausgabedatei, hmi_name, bibliotheksname)
            ergebnis_warteschlange.put(('erfolg', len(zu_erstellende_rahmen)))

        # Fehlerbehandlung: Datei nicht gefunden bzw. unerwarteter Fehler
        except FileNotFoundError:
            ergebnis_warteschlange.put(('datei_nicht_gefunden', None))
        except Exception as e:
            ergebnis_warteschlange.put(('fehler', str(e)))

    # Hintergrund-Thread starten (daemon=True: wird mit dem Programm beendet)
    thread = threading.Thread(target=hintergrundarbeit, daemon=True)
    thread.start()

    # Prüft alle 100 ms, ob der Thread fertig ist, und zeigt dann das Ergebnis an
    def ergebnis_pruefen():
        if thread.is_alive():
            hauptfenster.after(100, ergebnis_pruefen)
            return

        # Fortschrittsfenster schließen
        fortschrittsbalken.stop()
        fortschritts_fenster.grab_release()
        fortschritts_fenster.destroy()

        # typ = Art des Ergebnisses, daten = zugehörige Informationen
        typ, daten = ergebnis_warteschlange.get()

        # Passende Meldung je nach Ergebnistyp anzeigen
        if typ == 'rahmenname_fehlt':
            fehlermeldung = ("Das Skript kann nicht erzeugt werden. In folgenden Zeilen der Excel-Datei "
                              "fehlt der Name des Rahmens (erste Spalte), obwohl weitere Daten eingetragen wurden:\n\n")
            fehlermeldung += "\n".join(daten)
            fehlermeldung += ("\n\nTrage in der ersten Spalte den Namen des gewünschten Rahmens ein oder "
                               "lösche die Zeile vollständig, bevor du das Skript erzeugst.")
            messagebox.showerror("Fehler - Rahmenname fehlt", fehlermeldung)

        elif typ == 'unvollstaendige_position':
            fehlermeldung = ("Das Skript kann nicht erzeugt werden. In folgenden Zeilen der Excel-Datei "
                              "wurde nur einer der beiden erforderlichen Koordinatenwerte angegeben:\n\n")
            fehlermeldung += "\n".join(daten)
            fehlermeldung += ("\n\nFür die Position werden immer X (vorletzte Spalte) UND Y (letzte Spalte) benötigt. "
                               "Trage den fehlenden Wert ein oder lösche beide Werte, "
                               "damit die automatische Platzierung verwendet wird.")
            messagebox.showerror("Fehler - Unvollständige Position", fehlermeldung)

        elif typ == 'rahmen_nicht_gefunden':
            fehlermeldung = ("Das Skript kann nicht erzeugt werden. Es wird auf Rahmen verwiesen, die nicht "
                              "in der Rahmenbibliothek existieren:\n\n")
            fehlermeldung += "\n".join(daten)
            fehlermeldung += ("\n\nÜberprüfe den Namen des Rahmens in der Excel-Datei oder deklariere den Rahmen "
                               "auf der Registerkarte 'Verfügbare Rahmen', bevor du das Skript erzeugst.")
            messagebox.showerror("Fehler - Rahmen nicht gefunden", fehlermeldung)

        elif typ == 'fehlende_warnungen':
            fehlermeldung = "Das Skript kann nicht erzeugt werden. In der Excel-Datei fehlen Variablen:\n\n"
            fehlermeldung += "\n".join(daten)
            messagebox.showerror("Fehler - Unvollständige Variablen", fehlermeldung)

        elif typ == 'erfolg':
            messagebox.showinfo("Erfolg", f"'{ausgabedatei}' wurde erfolgreich erzeugt.\nEs wurden {daten} Rahmen sowie die Start-Schaltfläche konfiguriert.")

        elif typ == 'datei_nicht_gefunden':
            messagebox.showerror("Fehler", "Die angegebene Excel-Datei konnte nicht gefunden werden.")

        elif typ == 'fehler':
            messagebox.showerror("Unerwarteter Fehler", f"Es ist ein Problem aufgetreten: {daten}")

    hauptfenster.after(100, ergebnis_pruefen)

# Füllt die Rahmenliste neu (benutzerdefinierte Rahmen in Blau) und leert die Variablentabellen
def rahmenliste_aktualisieren():
    listbox_rahmen.delete(0, tk.END)
    alle_rahmen = alle_rahmen_abrufen()
    
    for index, rahmen_name in enumerate(sorted(alle_rahmen.keys())):
        listbox_rahmen.insert(tk.END, rahmen_name)
        if rahmen_name in BENUTZERDEFINIERTE_RAHMEN:
            listbox_rahmen.itemconfig(index, {'fg': 'blue'})
            
    for item in baum_variablen_input.get_children():
        baum_variablen_input.delete(item)
    for item in baum_variablen_inout.get_children():
        baum_variablen_inout.delete(item)

# Entfernt den ausgewählten Rahmen (gilt nur für die aktuelle Sitzung)
def rahmen_entfernen():
    # auswahl = Index des markierten Listeneintrags
    auswahl = listbox_rahmen.curselection()
    if not auswahl:
        messagebox.showwarning("Achtung", "Wähle zuerst einen Rahmen aus der Liste aus, um ihn zu entfernen.")
        return
        
    ausgewaehlter_rahmen = listbox_rahmen.get(auswahl[0])
    
    if messagebox.askyesno("Bestätigen", f"Bist du sicher, dass du '{ausgewaehlter_rahmen}' entfernen möchtest?\n\n(Diese Änderung wird beim Neustart des Programms rückgängig gemacht)."):
        # Benutzerdefinierte Rahmen löschen, Bibliotheksrahmen nur als 'entfernt' markieren
        if ausgewaehlter_rahmen in BENUTZERDEFINIERTE_RAHMEN:
            del BENUTZERDEFINIERTE_RAHMEN[ausgewaehlter_rahmen]
        else:
            ENTFERNTE_RAHMEN.add(ausgewaehlter_rahmen)
        
        GEAENDERTE_GROESSEN.pop(ausgewaehlter_rahmen, None)
            
        rahmenliste_aktualisieren()

# Speichert die aktuelle Rahmenliste als neue Version dieses Programms (Rückgabe: True/False)
def neue_version_speichern():
    """
    Gibt True zurück, wenn erfolgreich gespeichert wurde, sonst False.
    """
    alle = alle_rahmen_abrufen()

    if not alle:
        messagebox.showwarning("Achtung", "Es sind keine Rahmen zum Speichern verfügbar.")
        return False

    try:
        # Pfad der aktuell laufenden Programmdatei ermitteln
        aktueller_skript_pfad = os.path.abspath(__file__)
    except NameError:
        messagebox.showerror("Fehler", "Der Speicherort der aktuellen Programmdatei konnte nicht ermittelt werden.")
        return False

    # Vorgeschlagener Dateiname der neuen Version (Endung _v2)
    vorgeschlagener_name = os.path.splitext(os.path.basename(aktueller_skript_pfad))[0] + "_v2.py"

    # Speicherdialog für die neue Version
    neuer_pfad = filedialog.asksaveasfilename(
        title="Neue Version des Programms mit dieser Rahmenbibliothek speichern",
        defaultextension=".py",
        filetypes=[("Python-Datei", "*.py")],
        initialfile=vorgeschlagener_name
    )
    if not neuer_pfad:
        return False

    if not neuer_pfad.lower().endswith(".py"):
        neuer_pfad += ".py"

    if os.path.abspath(neuer_pfad) == aktueller_skript_pfad:
        messagebox.showerror(
            "Fehler",
            "Du musst die neue Version unter einem anderen Namen oder an einem anderen Ort speichern "
            "als die aktuell verwendete Datei, um das aktuelle Programm nicht zu überschreiben."
        )
        return False

    try:
        with open(aktueller_skript_pfad, "r", encoding="utf-8") as f:
            aktueller_code = f.read()

        # Neuen Quelltext-Block der Bibliothek aus den aktuellen Rahmen aufbauen
        dict_zeilen = [f"    {repr(name)}: {repr(alle[name])}," for name in sorted(alle.keys())]
        neuer_block = "RAHMEN_BIBLIOTHEK = {\n" + "\n".join(dict_zeilen) + "\n}"

        # Regulärer Ausdruck: findet den bisherigen Bibliotheks-Block und ersetzt ihn durch den neuen
        muster = re.compile(r"RAHMEN_BIBLIOTHEK = \{.*?^\}", re.DOTALL | re.MULTILINE)
        neuer_code, anzahl_ersetzungen = muster.subn(neuer_block, aktueller_code, count=1)

        if anzahl_ersetzungen == 0:
            messagebox.showerror(
                "Fehler",
                "Die Rahmenbibliothek konnte im aktuellen Quellcode nicht gefunden werden.\n"
                "Es wurde keine neue Version gespeichert."
            )
            return False

        with open(neuer_pfad, "w", encoding="utf-8") as f:
            f.write(neuer_code)

        messagebox.showinfo(
            "Erfolg",
            f"Die neue Version des Programms wurde gespeichert unter:\n{neuer_pfad}\n\n"
            f"Die {len(alle)} aktuellen Rahmen sind nun die BASIS-Bibliothek dieser Datei."
        )
        
        # Nach dem Speichern leeren wir die temporären Listen, damit beim Schließen nicht erneut gefragt wird
        BENUTZERDEFINIERTE_RAHMEN.clear()
        ENTFERNTE_RAHMEN.clear()
        GEAENDERTE_GROESSEN.clear()
        rahmenliste_aktualisieren()
        
        return True
    except Exception as e:
        messagebox.showerror("Unerwarteter Fehler", f"Die neue Version konnte nicht gespeichert werden: {str(e)}")
        return False

# Wird beim Schließen des Fensters aufgerufen (siehe protocol WM_DELETE_WINDOW)
def beim_schliessen():
    """Funktion, die das Schließereignis des Fensters abfängt, um Datenverlust zu verhindern."""
    # Bei ungespeicherten Änderungen fragen: Ja = speichern, Nein = verwerfen, Abbrechen = weiterarbeiten
    if len(BENUTZERDEFINIERTE_RAHMEN) > 0 or len(ENTFERNTE_RAHMEN) > 0 or len(GEAENDERTE_GROESSEN) > 0:
        antwort = messagebox.askyesnocancel(
            "Nicht gespeicherte Änderungen",
            "Du hast Änderungen an den verfügbaren Rahmen vorgenommen (hinzugefügt, entfernt oder deren "
            "Standardgröße geändert).\n\n"
            "Möchtest du diese Änderungen vor dem Beenden als neue Version des Programms speichern?"
        )
        
        if antwort is True:  # Der Benutzer hat Ja gewählt
            erfolg = neue_version_speichern()
            if erfolg:
                hauptfenster.destroy()
        elif antwort is False:  # Der Benutzer hat Nein gewählt
            hauptfenster.destroy()
        # Wenn die Antwort None ist (Abbrechen), endet die Funktion ohne das Fenster zu schließen
    else:
        hauptfenster.destroy()

# Wird bei Auswahl eines Rahmens aufgerufen: zeigt Variablen (INPUT / IN_OUT) und Standardgröße an
# event = Tkinter-Ereignis (wird nicht ausgewertet)
def variablentabelle_aktualisieren(event):
    auswahl = listbox_rahmen.curselection()
    if not auswahl:
        return
        
    ausgewaehlter_rahmen = listbox_rahmen.get(auswahl[0])
    alle_rahmen = alle_rahmen_abrufen()
    
    for item in baum_variablen_input.get_children():
        baum_variablen_input.delete(item)
    for item in baum_variablen_inout.get_children():
        baum_variablen_inout.delete(item)
        
    if ausgewaehlter_rahmen in alle_rahmen:
        # Variablen nach Richtung auf die beiden Tabellen verteilen
        rahmen_variablen = alle_rahmen[ausgewaehlter_rahmen]['vars']
        for var in rahmen_variablen:
            if var['dir'] == 'INPUT':
                baum_variablen_input.insert("", tk.END, values=(var['name'], var['dir'], var['type']))
            elif var['dir'] == 'IN_OUT':
                baum_variablen_inout.insert("", tk.END, values=(var['name'], var['dir'], var['type']))
        
        # Bearbeitungsmodus der Standardgröße immer zurücksetzen (deaktiviert) und die
        # aktuellen Werte des neu ausgewählten Rahmens anzeigen.
        var_groesse_bearbeiten.set(False)
        groesse_bearbeiten_umschalten()
        
        eingabe_groesse_breite.config(state=tk.NORMAL)
        eingabe_groesse_breite.delete(0, tk.END)
        eingabe_groesse_breite.insert(0, str(alle_rahmen[ausgewaehlter_rahmen]['w']))
        eingabe_groesse_breite.config(state=tk.DISABLED)
        
        eingabe_groesse_hoehe.config(state=tk.NORMAL)
        eingabe_groesse_hoehe.delete(0, tk.END)
        eingabe_groesse_hoehe.insert(0, str(alle_rahmen[ausgewaehlter_rahmen]['h']))
        eingabe_groesse_hoehe.config(state=tk.DISABLED)

# Öffnet ein Fenster zum Deklarieren eines neuen Rahmens (Name, Größe, Variablen)
def neues_rahmen_fenster_oeffnen():
    # top = Fenster 'Neuer Rahmen' (modal durch grab_set)
    top = tk.Toplevel(hauptfenster)
    top.title("Neuer Rahmen")
    top.geometry("520x450")
    top.resizable(False, False)
    top.grab_set() 

    # neue_variablen: Variablen des neuen Rahmens (Name, Typ, Richtung)
    neue_variablen = []

    info_rahmen = ttk.Frame(top, padding=10)
    info_rahmen.pack(fill='x')

    ttk.Label(info_rahmen, text="Name des Rahmens:").grid(row=0, column=0, sticky="e", pady=5)
    eingabe_name = ttk.Entry(info_rahmen, width=25)
    eingabe_name.grid(row=0, column=1, pady=5, sticky="w", padx=5)

    ttk.Label(info_rahmen, text="Standardbreite:").grid(row=1, column=0, sticky="e", pady=5)
    eingabe_breite = ttk.Entry(info_rahmen, width=10)
    eingabe_breite.grid(row=1, column=1, pady=5, sticky="w", padx=5)

    ttk.Label(info_rahmen, text="Standardhöhe:").grid(row=2, column=0, sticky="e", pady=5)
    eingabe_hoehe = ttk.Entry(info_rahmen, width=10)
    eingabe_hoehe.grid(row=2, column=1, pady=5, sticky="w", padx=5)

    variablen_rahmen = ttk.LabelFrame(top, text="Variablen zum Rahmen hinzufügen", padding=10)
    variablen_rahmen.pack(fill='both', expand=True, padx=10, pady=5)

    ttk.Label(variablen_rahmen, text="Name:").grid(row=0, column=0, padx=2)
    eingabe_var_name = ttk.Entry(variablen_rahmen, width=15)
    eingabe_var_name.grid(row=0, column=1, padx=2)

    ttk.Label(variablen_rahmen, text="Typ:").grid(row=0, column=2, padx=2)
    eingabe_var_typ = ttk.Entry(variablen_rahmen, width=12)
    eingabe_var_typ.grid(row=0, column=3, padx=2)

    ttk.Label(variablen_rahmen, text="Richtung:").grid(row=0, column=4, padx=2)
    cb_richtung = ttk.Combobox(variablen_rahmen, values=["INPUT", "IN_OUT"], state="readonly", width=8)
    cb_richtung.current(0)
    cb_richtung.grid(row=0, column=5, padx=2)

    # Fügt die eingegebene Variable der Liste und der Vorschau-Tabelle hinzu
    def variable_hinzufuegen():
        name = eingabe_var_name.get().strip()
        typ = eingabe_var_typ.get().strip()
        richtung = cb_richtung.get().strip()
        
        if name and typ and richtung:
            neue_variablen.append({'name': name, 'type': typ, 'dir': richtung})
            baum_temp.insert("", tk.END, values=(name, richtung, typ))
            eingabe_var_name.delete(0, tk.END)
            eingabe_var_typ.delete(0, tk.END)
        else:
            messagebox.showwarning("Unvollständig", "Fülle alle Felder der Variable aus.", parent=top)

    ttk.Button(variablen_rahmen, text="Variable hinzufügen", command=variable_hinzufuegen).grid(row=1, column=0, columnspan=6, pady=10)

    baum_temp = ttk.Treeview(variablen_rahmen, columns=("name", "richtung", "typ"), show="headings", height=5)
    baum_temp.heading("name", text="Name")
    baum_temp.heading("richtung", text="Richtung")
    baum_temp.heading("typ", text="Typ")
    baum_temp.column("name", width=130)
    baum_temp.column("richtung", width=100)
    baum_temp.column("typ", width=130)
    baum_temp.grid(row=2, column=0, columnspan=6, sticky="we")

    # Prüft die Eingaben und speichert den Rahmen in BENUTZERDEFINIERTE_RAHMEN (nur Sitzung)
    def neuen_rahmen_speichern():
        rahmen_name = eingabe_name.get().strip()
        
        try:
            # Breite und Höhe müssen ganze Zahlen sein
            breite = int(eingabe_breite.get().strip())
            hoehe = int(eingabe_hoehe.get().strip())
        except ValueError:
            messagebox.showerror("Fehler", "Breite und Höhe müssen ganzzahlige numerische Werte sein.", parent=top)
            return
            
        if not rahmen_name:
            messagebox.showerror("Fehler", "Du musst einen Namen für den Rahmen eingeben.", parent=top)
            return
            
        BENUTZERDEFINIERTE_RAHMEN[rahmen_name] = {
            'w': breite,
            'h': hoehe,
            'vars': neue_variablen
        }
        
        if rahmen_name in ENTFERNTE_RAHMEN:
            ENTFERNTE_RAHMEN.remove(rahmen_name)
        
        rahmenliste_aktualisieren()
        messagebox.showinfo("Erfolg", f"Rahmen '{rahmen_name}' wurde erfolgreich hinzugefügt.", parent=top)
        top.destroy()

    ttk.Button(top, text="Rahmen speichern", command=neuen_rahmen_speichern, width=20).pack(pady=10)


# --- Konfiguration des Hauptfensters (Tkinter) ---

# hauptfenster: Hauptfenster der Anwendung
hauptfenster = tk.Tk()

# --- Anpassung des Fenstertitels ---
try:
    aktueller_dateiname = os.path.basename(__file__)
except NameError:
    aktueller_dateiname = "unbekannte_datei.py"
    
hauptfenster.title(f"Skriptgenerator für HMI ({aktueller_dateiname})")
hauptfenster.resizable(False, False)

# --- Protokoll zum Abfangen des Fensterschließens ---
hauptfenster.protocol("WM_DELETE_WINDOW", beim_schliessen)

# Notizbuch mit zwei Registerkarten: 'Skript erzeugen' und 'Verfügbare Rahmen'
notizbuch = ttk.Notebook(hauptfenster)
notizbuch.pack(fill='both', expand=True, padx=5, pady=5)

reiter_generieren = ttk.Frame(notizbuch)
reiter_rahmen = ttk.Frame(notizbuch)

notizbuch.add(reiter_generieren, text="Skript erzeugen")
notizbuch.add(reiter_rahmen, text="Verfügbare Rahmen")

# Passt die Fenstergröße an die größere der beiden Registerkarten an
def fenstergroesse_anpassen():
    hauptfenster.update_idletasks()
    aktiver_reiter = notizbuch.select()

    breiten = []
    hoehen = []
    for reiter in (reiter_generieren, reiter_rahmen):
        notizbuch.select(reiter)
        hauptfenster.update_idletasks()
        breiten.append(hauptfenster.winfo_reqwidth())
        hoehen.append(hauptfenster.winfo_reqheight())

    notizbuch.select(aktiver_reiter)
    hauptfenster.update_idletasks()

    breite = max(breiten)
    hoehe = max(hoehen)
    hauptfenster.geometry(f"{breite}x{hoehe}")


# --- INHALT REGISTERKARTE 1: Skript erzeugen ---
# Layout per grid (row = Zeile, column = Spalte). Eingabefelder:
# eingabe_excel = Excel-Pfad, eingabe_hmi = HMI-Name, eingabe_ausgabe = Name der Skriptdatei,
# eingabe_ausgabeordner = Zielordner, eingabe_bibliothek = Name der Bibliothek
generierungs_rahmen = ttk.Frame(reiter_generieren, padding=10)
generierungs_rahmen.pack(fill='both', expand=True)

ttk.Label(generierungs_rahmen, text="Anweisungsdatei (Excel):").grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 2))
eingabe_excel = ttk.Entry(generierungs_rahmen, width=65)
eingabe_excel.grid(row=1, column=0, sticky='w')
ttk.Button(generierungs_rahmen, text="Durchsuchen...", command=datei_suchen).grid(row=1, column=1, padx=(10, 0))

ttk.Label(generierungs_rahmen, text="Name der HMI im Projekt:").grid(row=2, column=0, columnspan=2, sticky='w', pady=(15, 2))
eingabe_hmi = ttk.Entry(generierungs_rahmen, width=35)
eingabe_hmi.insert(0, "Visualisierung_Motor")
eingabe_hmi.grid(row=3, column=0, sticky='w')

ttk.Label(generierungs_rahmen, text="Name der Ausgabe-Python-Datei:").grid(row=4, column=0, columnspan=2, sticky='w', pady=(15, 2))
eingabe_ausgabe = ttk.Entry(generierungs_rahmen, width=35)
eingabe_ausgabe.insert(0, "erzeugtes_skript")
eingabe_ausgabe.grid(row=5, column=0, sticky='w')

ttk.Label(generierungs_rahmen, text="Ordner zum Speichern der erzeugten Datei:").grid(row=6, column=0, columnspan=2, sticky='w', pady=(15, 2))
eingabe_ausgabeordner = ttk.Entry(generierungs_rahmen, width=65)
eingabe_ausgabeordner.grid(row=7, column=0, sticky='w')
ttk.Button(generierungs_rahmen, text="Durchsuchen...", command=ausgabeordner_suchen).grid(row=7, column=1, padx=(10, 0))

bibliothek_rahmen = ttk.Frame(generierungs_rahmen)
bibliothek_rahmen.grid(row=8, column=0, columnspan=2, sticky='we', pady=(15, 0))

# Zustand der Checkbox (an = Bibliotheksname kann geändert werden)
var_check_bibliothek = tk.BooleanVar(value=False)
check_bibliothek = ttk.Checkbutton(bibliothek_rahmen, text="Bibliotheksnamen ändern:", variable=var_check_bibliothek, command=bibliothek_umschalten)
check_bibliothek.pack(side='left')

eingabe_bibliothek = ttk.Entry(bibliothek_rahmen)
eingabe_bibliothek.insert(0, "AutomationLibrary")
eingabe_bibliothek.config(state=tk.DISABLED)
eingabe_bibliothek.pack(side='left', fill='x', expand=True, padx=(10, 0))

schaltflaechen_rahmen = ttk.Frame(generierungs_rahmen)
schaltflaechen_rahmen.grid(row=9, column=0, columnspan=2, pady=(25, 10))
schaltflaeche_generieren = ttk.Button(schaltflaechen_rahmen, text="Skript erzeugen", command=generierung_ausfuehren)
schaltflaeche_generieren.pack()

trenner = ttk.Separator(generierungs_rahmen, orient='horizontal')
trenner.grid(row=10, column=0, columnspan=2, sticky='ew', pady=(10, 15))

fusszeile_schaltflaechen_rahmen = ttk.Frame(generierungs_rahmen)
fusszeile_schaltflaechen_rahmen.grid(row=11, column=0, columnspan=2)
schaltflaeche_konfig_speichern = ttk.Button(fusszeile_schaltflaechen_rahmen, text="Aktuelle Konfiguration speichern...", command=neue_version_speichern)
schaltflaeche_konfig_speichern.pack()

ttk.Label(
    generierungs_rahmen,
    text="Speichert die aktuellen Rahmen (mit den Änderungen dieser Sitzung) als\n"
         "Basisbibliothek einer neuen Kopie des Programms (einer weiteren .py-Datei).",
    foreground="#555"
).grid(row=12, column=0, columnspan=2, pady=(10, 0), sticky='w')


# --- INHALT REGISTERKARTE 2: Verfügbare Rahmen ---
kopf_rahmen = ttk.Frame(reiter_rahmen, padding=(10, 10, 5, 5))
kopf_rahmen.pack(side='top', fill='x')
ttk.Label(kopf_rahmen, text="Verfügbare Rahmen:").pack(anchor='w')
tk.Label(kopf_rahmen, text="In Blau: vom Benutzer erstellte Rahmen (temporär).", fg="#0b63c5").pack(anchor='w')

werkzeugleisten_rahmen = ttk.Frame(reiter_rahmen, padding=(10, 0, 10, 5))
werkzeugleisten_rahmen.pack(side='top', fill='x')

schaltflaeche_neuer_rahmen = ttk.Button(werkzeugleisten_rahmen, text="Neuen Rahmen deklarieren", command=neues_rahmen_fenster_oeffnen)
schaltflaeche_neuer_rahmen.pack(side='left', padx=(0, 10))

schaltflaeche_rahmen_entfernen = ttk.Button(werkzeugleisten_rahmen, text="Rahmen entfernen", command=rahmen_entfernen)
schaltflaeche_rahmen_entfernen.pack(side='left')

inhalts_rahmen = ttk.Frame(reiter_rahmen, padding=(10, 5, 10, 10))
inhalts_rahmen.pack(side='top', fill='both', expand=True)
inhalts_rahmen.columnconfigure(1, weight=1)
inhalts_rahmen.rowconfigure(0, weight=1)

# --- RAHMENLISTE (LINKS) ---
listen_rahmen = ttk.Frame(inhalts_rahmen)
listen_rahmen.grid(row=0, column=0, sticky='ns', padx=(0, 10))

# Liste aller verfügbaren Rahmen
listbox_rahmen = tk.Listbox(listen_rahmen, width=25, height=20, exportselection=False)
listbox_rahmen.pack(side='left', fill='y')
scroll_liste = ttk.Scrollbar(listen_rahmen, orient='vertical', command=listbox_rahmen.yview)
scroll_liste.pack(side='right', fill='y')
listbox_rahmen.config(yscrollcommand=scroll_liste.set)

# --- VARIABLENTABELLEN (RECHTS) ---
tabellen_rahmen = ttk.Frame(inhalts_rahmen)
tabellen_rahmen.grid(row=0, column=1, sticky='nsew')

tabellen_rahmen.rowconfigure(1, weight=1)
tabellen_rahmen.rowconfigure(3, weight=1)
tabellen_rahmen.columnconfigure(0, weight=1)

# Spaltenkennungen der Variablentabellen
spalten = ("rahmen", "richtung", "typ")

# -- TABELLE 1: INPUT (Oben) --
ttk.Label(tabellen_rahmen, text="INPUT-Variablen:").grid(row=0, column=0, sticky='w', pady=(0, 2))

baum_rahmen_eingang = ttk.Frame(tabellen_rahmen)
baum_rahmen_eingang.grid(row=1, column=0, sticky='nsew', pady=(0, 10))

# Tabelle (Treeview) für die INPUT-Variablen: Name, Variablentyp, Datentyp
baum_variablen_input = ttk.Treeview(baum_rahmen_eingang, columns=spalten, show="headings")
baum_variablen_input.heading("rahmen", text="Rahmen")
baum_variablen_input.heading("richtung", text="Variablentyp")
baum_variablen_input.heading("typ", text="Datentyp")
baum_variablen_input.column("rahmen", width=120)
baum_variablen_input.column("richtung", width=110)
baum_variablen_input.column("typ", width=110)

baum_variablen_input.pack(side='left', fill='both', expand=True)
scroll_tabelle_eingang = ttk.Scrollbar(baum_rahmen_eingang, orient='vertical', command=baum_variablen_input.yview)
baum_variablen_input.configure(yscrollcommand=scroll_tabelle_eingang.set)
scroll_tabelle_eingang.pack(side='right', fill='y')

# -- TABELLE 2: IN_OUT (Unten) --
ttk.Label(tabellen_rahmen, text="IN_OUT-Variablen:").grid(row=2, column=0, sticky='w', pady=(0, 2))

baum_rahmen_ausgang = ttk.Frame(tabellen_rahmen)
baum_rahmen_ausgang.grid(row=3, column=0, sticky='nsew')

# Tabelle (Treeview) für die IN_OUT-Variablen: Name, Variablentyp, Datentyp
baum_variablen_inout = ttk.Treeview(baum_rahmen_ausgang, columns=spalten, show="headings")
baum_variablen_inout.heading("rahmen", text="Rahmen")
baum_variablen_inout.heading("richtung", text="Variablentyp")
baum_variablen_inout.heading("typ", text="Datentyp")
baum_variablen_inout.column("rahmen", width=120)
baum_variablen_inout.column("richtung", width=110)
baum_variablen_inout.column("typ", width=110)

baum_variablen_inout.pack(side='left', fill='both', expand=True)
scroll_tabelle_ausgang = ttk.Scrollbar(baum_rahmen_ausgang, orient='vertical', command=baum_variablen_inout.yview)
baum_variablen_inout.configure(yscrollcommand=scroll_tabelle_ausgang.set)
scroll_tabelle_ausgang.pack(side='right', fill='y')

# -- SEKTION: Standardgröße des ausgewählten Rahmens bearbeiten --
groesse_rahmen = ttk.LabelFrame(tabellen_rahmen, text="Standardgröße des ausgewählten Rahmens", padding=10)
groesse_rahmen.grid(row=4, column=0, sticky='we', pady=(10, 0))

ttk.Label(groesse_rahmen, text="Breite:").grid(row=0, column=0, sticky='e', padx=(0, 5))
eingabe_groesse_breite = ttk.Entry(groesse_rahmen, width=10, state=tk.DISABLED)
eingabe_groesse_breite.grid(row=0, column=1, padx=(0, 15))

ttk.Label(groesse_rahmen, text="Höhe:").grid(row=0, column=2, sticky='e', padx=(0, 5))
eingabe_groesse_hoehe = ttk.Entry(groesse_rahmen, width=10, state=tk.DISABLED)
eingabe_groesse_hoehe.grid(row=0, column=3, padx=(0, 15))

# Gibt die Felder Breite/Höhe je nach Checkbox frei oder sperrt sie
def groesse_bearbeiten_umschalten():
    if var_groesse_bearbeiten.get():
        eingabe_groesse_breite.config(state=tk.NORMAL)
        eingabe_groesse_hoehe.config(state=tk.NORMAL)
    else:
        eingabe_groesse_breite.config(state=tk.DISABLED)
        eingabe_groesse_hoehe.config(state=tk.DISABLED)

# Zustand der Checkbox 'Größe bearbeiten'
var_groesse_bearbeiten = tk.BooleanVar(value=False)
check_groesse_bearbeiten = ttk.Checkbutton(
    groesse_rahmen, text="Größe bearbeiten", variable=var_groesse_bearbeiten, command=groesse_bearbeiten_umschalten
)
check_groesse_bearbeiten.grid(row=0, column=4, padx=(10, 15))

# Übernimmt die geänderte Standardgröße (nur Sitzung, gespeichert in GEAENDERTE_GROESSEN)
def groesse_bestaetigen():
    auswahl = listbox_rahmen.curselection()
    if not auswahl:
        messagebox.showwarning("Achtung", "Wähle zuerst einen Rahmen aus der Liste aus.")
        return

    if not var_groesse_bearbeiten.get():
        messagebox.showinfo("Hinweis", "Aktiviere zuerst 'Größe bearbeiten', um die Standardgröße zu ändern.")
        return

    ausgewaehlter_rahmen = listbox_rahmen.get(auswahl[0])

    try:
        # Eingaben in ganze Zahlen umwandeln
        neue_breite = int(eingabe_groesse_breite.get().strip())
        neue_hoehe = int(eingabe_groesse_hoehe.get().strip())
    except ValueError:
        messagebox.showerror("Fehler", "Breite und Höhe müssen ganzzahlige numerische Werte sein.")
        return

    # Breite und Höhe müssen positiv sein
    if neue_breite <= 0 or neue_hoehe <= 0:
        messagebox.showerror("Fehler", "Breite und Höhe müssen größer als 0 sein.")
        return

    GEAENDERTE_GROESSEN[ausgewaehlter_rahmen] = (neue_breite, neue_hoehe)

    var_groesse_bearbeiten.set(False)
    groesse_bearbeiten_umschalten()

    messagebox.showinfo(
        "Erfolg",
        f"Die Standardgröße von '{ausgewaehlter_rahmen}' wurde auf {neue_breite} x {neue_hoehe} geändert.\n\n"
        "(Diese Änderung wird beim Neustart des Programms rückgängig gemacht, außer sie wird als neue "
        "Version des Programms gespeichert)."
    )

    rahmenliste_aktualisieren()

    # Auswahl des bearbeiteten Rahmens wiederherstellen, damit die neue Größe sofort sichtbar ist.
    for i in range(listbox_rahmen.size()):
        if listbox_rahmen.get(i) == ausgewaehlter_rahmen:
            listbox_rahmen.selection_set(i)
            listbox_rahmen.event_generate("<<ListboxSelect>>")
            break

schaltflaeche_groesse_bestaetigen = ttk.Button(groesse_rahmen, text="Änderung bestätigen", command=groesse_bestaetigen)
schaltflaeche_groesse_bestaetigen.grid(row=0, column=5, padx=(10, 0))

# Bei Auswahl eines Listeneintrags die Variablentabellen aktualisieren
listbox_rahmen.bind("<<ListboxSelect>>", variablentabelle_aktualisieren)

# Startzustand: Liste füllen und ersten Rahmen auswählen
rahmenliste_aktualisieren()

if listbox_rahmen.size() > 0:
    listbox_rahmen.selection_set(0)
    listbox_rahmen.event_generate("<<ListboxSelect>>")

# Fenstergröße anpassen
fenstergroesse_anpassen()

# Programmstart: Hauptschleife der GUI (wartet auf Benutzeraktionen)
if __name__ == "__main__":
    hauptfenster.mainloop()
