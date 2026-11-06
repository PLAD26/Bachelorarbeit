# -*- coding: utf-8 -*-
# Autor: Ali Duarte Perez Luna
# Datum: 10.10.2026
# Zweck: Liest eine Excel-Tabelle mit Tasks ein und erzeugt daraus ein CODESYS-Skript.
"""
Generator für CODESYS-Skripte zum Erstellen von Tasks aus einer Excel-Datei.

Verwendung:
    python task_skript_generator.py
        -> öffnet ein Fenster: Excel-Datei wählen, Zielordner wählen, den
           Skriptnamen bearbeiten und auf "Skript erzeugen" klicken.
    python task_skript_generator.py Task_Konfiguration.xlsx [ausgabe.py]
        -> Verwendung über die Kommandozeile (ohne Fenster).
Das erzeugte Skript wird in CODESYS ausgeführt über:
    Tools -> Scripting -> Skriptdatei ausfuehren...

Format der Excel-Datei (Zeile 1 enthält die Überschriften und wird ignoriert):
    A: Name der Task
    B: Priorität
    C: Typ: Zyklisch | Ereignis | Externes | Freilaufend | Status
    D: Intervall (ms, z. B. 20 -> T#20ms; auch "T#20ms" ist zulässig).
       Bei Ereignis/Externes wird der Wert als Ereignisname interpretiert.
       Bei Freilaufend/Status wird er ignoriert.
    E: Watchdog: "Nein"  oder  "Ja (Zeit, Empfindlichkeit)"  z. B. "Ja (4000, 1)"
    F und folgende: Namen der Programme (POUs) der Task
"""
from __future__ import print_function  # Kompatibilität: print() als Funktion (auch für ältere Python-Versionen)

# Importe: os = Dateipfade, re = reguläre Ausdrücke (Textmuster), sys = Kommandozeilenargumente
import os
import re
import sys

from openpyxl import load_workbook  # load_workbook: öffnet eine Excel-Datei

# Typ auf Deutsch (Excel) -> Name der Aufzählung KindOfTask in CODESYS
TYPEN = {
    "zyklisch": "Cyclic",
    "ereignis": "Event",
    "externes": "External",
    "freilaufend": "Freewheeling",
    "status": "Status",
}


def als_text(wert):
    """Wandelt eine Zelle in bereinigten Text um ('' wenn sie leer ist)."""
    if wert is None:  # leere Zelle -> leerer Text
        return ""
    if isinstance(wert, float) and wert.is_integer():  # 20.0 -> 20 (Excel liefert Zahlen oft als float)
        wert = int(wert)
    return str(wert).strip()  # in Text umwandeln und Leerzeichen entfernen


class ExcelFehler(ValueError):
    """Validierungsfehler der Excel-Datei (enthält alle gefundenen Probleme)."""


def parse_intervall(wert, typ):
    """Liefert das Intervall als IEC-Zeichenkette (T#20ms) oder Ereignis; '' wenn nicht zutreffend."""
    t = als_text(wert)  # t = Zellwert als Text
    if typ in ("Freewheeling", "Status"):
        return ""  # diese Typen verwenden kein Intervall
    if not t:  # Zelle leer -> Fehlermeldung je nach Typ
        if typ == "Cyclic":
            raise ValueError("Das Intervall fehlt (bei einer Task vom Typ Zyklisch erforderlich).")
        raise ValueError("Das Ereignis in der Intervallspalte fehlt "
                         "(bei einer Task vom Typ %s erforderlich)." %
                         ("Ereignis" if typ == "Event" else "Externes"))
    if typ in ("Event", "External"):  # Ereignisname unverändert übernehmen
        return t
    # Cyclic
    if re.fullmatch(r"\d+(\.\d+)?", t):  # nur Zahl, z. B. 20 -> T#20ms
        return "T#" + t + "ms"
    if re.fullmatch(r"\d+\s*ms", t, re.I):  # Zahl mit 'ms', z. B. 20 ms -> T#20ms
        return "T#" + re.sub(r"\s*ms", "", t, flags=re.I) + "ms"
    if t.upper().startswith("T#"):  # bereits im IEC-Format, z. B. T#20ms
        return t
    # Alles andere ist kein gültiges Intervall
    raise ValueError("Das Intervall %r ist ungültig; bitte eine Zahl in ms "
                     "(z. B. 20) oder einen IEC-Wert (z. B. T#20ms) verwenden." % t)


def parse_watchdog(wert):
    """None bei 'Nein'; (Zeit, Empfindlichkeit) bei 'Ja (t, s)'."""
    t = als_text(wert)  # t = Zellwert als Text
    if not t:
        raise ValueError("Der Watchdog fehlt; bitte 'Nein' oder 'Ja (Zeit, Empfindlichkeit)' eintragen.")
    klein = t.lower()  # klein = Kleinbuchstaben für den Vergleich
    if klein.startswith("nein"):  # 'Nein' -> Watchdog deaktiviert
        return None
    if not klein.startswith("ja"):  # weder 'Nein' noch 'Ja' -> ungültig
        raise ValueError("Der Watchdog %r ist ungültig; bitte 'Nein' oder "
                         "'Ja (Zeit, Empfindlichkeit)' eintragen." % t)
    rest = t[2:].strip()  # rest = Text nach 'Ja', z. B. (4000, 1)
    if not rest:
        raise ValueError("Der Watchdog ist 'Ja', aber die Argumente fehlen; "
                         "bitte 'Ja (Zeit, Empfindlichkeit)' verwenden.")
    if not (rest.startswith("(") and rest.endswith(")")):  # Klammern müssen vorhanden sein
        raise ValueError("Der Watchdog %r hat nicht das richtige Format; "
                         "bitte 'Ja (Zeit, Empfindlichkeit)' verwenden." % t)
    argumente = [a.strip() for a in re.split(r"[,;]", rest[1:-1])]  # Argumente an Komma oder Semikolon trennen
    if len(argumente) > 2:
        raise ValueError("Der Watchdog %r hat zu viele Argumente; "
                         "erlaubt sind nur Zeit und Empfindlichkeit." % t)
    fehlend = []  # fehlend = Liste der fehlenden Angaben
    if not argumente[0]:
        fehlend.append("die Zeit (1. Argument)")
    if len(argumente) < 2 or not argumente[1]:
        fehlend.append("die Empfindlichkeit (2. Argument)")
    if fehlend:
        raise ValueError("Beim Watchdog %r fehlen folgende Angaben: %s." % (t, " und ".join(fehlend)))
    # Zeit und Empfindlichkeit müssen ganze Zahlen sein
    for bezeichnung, argument in (("Zeit", argumente[0]), ("Empfindlichkeit", argumente[1])):
        if not argument.isdigit():
            raise ValueError("Das Argument '%s' des Watchdogs muss eine ganze "
                             "Zahl sein, ist aber %r." % (bezeichnung, argument))
    return (int(argumente[0]), int(argumente[1]))  # Ergebnis: (Zeit in ms, Empfindlichkeit)


def lese_excel(pfad):
    """Liest und validiert die Excel-Datei. Löst ExcelFehler mit ALLEN gefundenen Problemen aus."""
    ws = load_workbook(pfad, data_only=True).worksheets[0]  # erstes Tabellenblatt; data_only = Werte statt Formeln
    tasks, fehlerliste = [], []  # tasks = gültige Tasks, fehlerliste = gesammelte Fehlermeldungen
    namen = {}           # Name (Kleinbuchstaben) -> [Zeilen]
    programm_aufrufe = {}  # Programm (Kleinbuchstaben) -> [(Zeile, Task)]

    # Alle Zeilen ab Zeile 2 durchlaufen (Zeile 1 = Überschriften)
    for zeile, zellen in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        zellen = list(zellen) + [None] * 6  # auf mindestens 6 Spalten auffüllen
        if not any(als_text(z) for z in zellen):
            continue  # komplett leere Zeile
        name = als_text(zellen[0])  # Spalte A: Name der Task
        praefix = "Zeile %d (Task %r): " % (zeile, name) if name else "Zeile %d: " % zeile  # praefix = Beginn jeder Fehlermeldung (Zeile und Task)
        zeile_ok = True  # Merker: Zeile bisher fehlerfrei

        def melde_fehler(meldung):  # hängt eine Fehlermeldung mit Präfix an die Liste an
            fehlerliste.append(praefix + meldung)

        if not name:
            melde_fehler("Der Task-Name fehlt.")
            zeile_ok = False

        prio = als_text(zellen[1])  # Spalte B: Priorität
        if not prio:
            melde_fehler("Die Priorität fehlt.")
            zeile_ok = False

        typ_text = als_text(zellen[2])  # Spalte C: Typ (deutsch)
        typ = TYPEN.get(typ_text.lower())  # deutscher Typ -> CODESYS-Name (None, wenn unbekannt)
        if not typ_text:
            melde_fehler("Der Task-Typ fehlt (Zyklisch, Ereignis, Externes, "
                         "Freilaufend oder Status).")
            zeile_ok = False
        elif typ is None:
            melde_fehler("Der Typ %r ist ungültig; erlaubt sind nur Zyklisch, "
                         "Ereignis, Externes, Freilaufend und Status." % typ_text)
            zeile_ok = False

        # Spalte D: Intervall bzw. Ereignis (nur bei gültigem Typ prüfbar)
        intervall = ""
        if typ is not None:
            try:
                intervall = parse_intervall(zellen[3], typ)
            except ValueError as e:
                melde_fehler(str(e))
                zeile_ok = False

        wd = None  # Spalte E: Watchdog (None = nicht aktiv)
        try:
            wd = parse_watchdog(zellen[4])
        except ValueError as e:
            melde_fehler(str(e))
            zeile_ok = False

        prog_namen = [als_text(z) for z in zellen[5:] if als_text(z)]  # Spalten ab F: Namen der Programme
        if not prog_namen:
            print("Hinweis: Zeile %d (%s) enthält keine Programme." % (zeile, name))

        # Namen und Programmaufrufe für die Prüfung auf Doppelte merken
        if name:
            namen.setdefault(name.lower(), []).append(zeile)
        for prog in prog_namen:
            programm_aufrufe.setdefault(prog.lower(), []).append((zeile, name or "(ohne Name)", prog))

        if zeile_ok:  # nur fehlerfreie Zeilen übernehmen
            tasks.append((name, prio, typ, intervall, prog_namen, wd))

    # Doppelte Task-Namen
    for zeilen in namen.values():
        if len(zeilen) > 1:
            task_name = als_text(ws.cell(row=zeilen[0], column=1).value)  # Schreibweise aus der ersten Zeile verwenden
            fehlerliste.append("Doppelter Task-Name: %r kommt in den Zeilen %s vor. Jede "
                               "Task muss einen eindeutigen Namen haben." %
                               (task_name, ", ".join(str(z) for z in zeilen)))

    # Programme, die mehr als einmal aufgerufen werden
    for verwendungen in programm_aufrufe.values():
        if len(verwendungen) > 1:
            stellen = ", ".join("Zeile %d (Task %r)" % (z, n) for z, n, _ in verwendungen)
            fehlerliste.append("Das Programm %r wird mehrfach aufgerufen: %s. Jedes Programm "
                               "darf nur von einer einzigen Task und nur einmal aufgerufen werden." %
                               (verwendungen[0][2], stellen))

    # Auswertung: bei Fehlern werden alle gemeldet
    if fehlerliste:
        raise ExcelFehler("\n".join("- " + f for f in fehlerliste))
    if not tasks:
        raise ExcelFehler("Die Excel-Datei enthält keine Task.")
    return tasks


# Vorlage des erzeugten CODESYS-Skripts. Die Platzhalter __EXCEL__, __ZUSAMMENFASSUNG__
# und __TASKS__ werden in erzeuge_skript() durch die echten Werte ersetzt.
VORLAGE = r'''# -*- coding: utf-8 -*-
# Erzeugt mit: task_skript_generator.py
"""
Erstellt in einem bestehenden CODESYS-Projekt eine Taskkonfiguration.
Automatisch generiert aus: __EXCEL__

Tasks:
__ZUSAMMENFASSUNG__

Ausfuehrung: Tools -> Scripting -> Skriptdatei ausfuehren...
"""
from __future__ import print_function  # Kompatibilität: print() als Funktion

# Konfiguration der Tasks:
# (Name, Prioritaet, Typ (KindOfTask), Intervall/Ereignis, [Programme], Watchdog)
# Watchdog = None (nicht aktiv) oder (Zeit in ms, Empfindlichkeit)
TASKS = [
__TASKS__
]


def zeige_projektbaum(projekt):
    """Gibt alle Objekte des Projektbaums zur Diagnose aus (rekursiv)."""
    print("Projekt: " + str(projekt))
    for objekt in projekt.get_children(True):  # True = alle Unterobjekte durchsuchen
        try:
            print("  - " + objekt.get_name())
        except Exception:
            print("  - (Name nicht lesbar)")


def finde_application(projekt):
    """Sucht die Application rekursiv ueber den Objekttyp (nicht ueber feste Namen)."""
    # 1) Rekursive Suche ueber die Typpruefung is_application
    for objekt in projekt.get_children(True):  # True = gesamten Baum durchsuchen
        try:
            if objekt.is_application:  # Typpruefung statt Namensvergleich
                return objekt
        except Exception:
            pass  # Objekt hat die Eigenschaft nicht -> naechstes Objekt
    # 2) Fallback: nach dem Namen "Application" suchen
    for objekt in projekt.get_children(True):
        try:
            if objekt.get_name() == "Application":
                return objekt
        except Exception:
            pass
    return None


def finde_taskkonfig(application):
    """Sucht eine vorhandene Taskkonfiguration ueber den Objekttyp (Name ist sprachabhaengig)."""
    for objekt in application.get_children(True):  # gesamten Baum unter der Application durchsuchen
        try:
            if objekt.is_task_configuration:  # Typpruefung statt Namensvergleich
                return objekt
        except Exception:
            pass  # Objekt hat die Eigenschaft nicht -> naechstes Objekt
    return None


def finde_kind(eltern, name):
    """Sucht ein direktes Kind-Objekt per Name."""
    for kind in eltern.get_children(False):  # nur direkte Kind-Objekte durchsuchen
        if kind.get_name() == name:
            return kind
    return None


def finde_kind_of_task(typ):
    """Sucht den Aufzaehlungswert KindOfTask.<typ> ueber verschiedene Zugriffswege."""
    werte = []  # werte = gefundene Kandidaten für den Aufzählungswert
    # 1) Globaler Name in der Skriptumgebung
    try:
        werte.append(getattr(KindOfTask, typ))  # noqa: F821
    except Exception:
        pass
    # 2) Ueber das Modul scriptengine
    try:
        import scriptengine
        werte.append(getattr(scriptengine.KindOfTask, typ))
    except Exception:
        pass
    # 3) Ueber .NET-Reflection in den geladenen Assemblies
    try:
        import clr
        import System
        for asm in System.AppDomain.CurrentDomain.GetAssemblies():
            for t in asm.GetTypes():
                if t.Name == "KindOfTask":
                    werte.append(System.Enum.Parse(t, typ))
                    raise StopIteration  # Suche beenden, sobald der Typ gefunden ist
    except StopIteration:  # normales Ende der Suche
        pass
    except Exception as fehler:
        print("Reflection fuer KindOfTask fehlgeschlagen: " + str(fehler))
    return werte


def setze_mit_varianten(objekt, attribut, varianten):
    """Setzt ein Attribut und probiert dabei mehrere Wertvarianten (String/Zahl/IEC-Zeit)."""
    for wert in varianten:  # Varianten nacheinander versuchen
        try:
            setattr(objekt, attribut, wert)  # Attribut mit der Variante setzen
            print("Gesetzt: " + attribut + " = " + repr(wert))
            return True  # Erfolg -> keine weitere Variante nötig
        except Exception as fehler:
            print(attribut + " mit " + repr(wert) + " fehlgeschlagen: " + str(fehler))
    return False  # keine Variante hat funktioniert


def setze_watchdog(task, zeit_ms, empfindlichkeit):
    """Aktiviert den Watchdog einer Task und setzt Zeit und Empfindlichkeit."""
    if not hasattr(task, "watchdog"):  # Eigenschaft muss vorhanden sein
        raise RuntimeError("Task hat keine Eigenschaft 'watchdog'.")
    wd = task.watchdog  # wd = Watchdog-Objekt der Task
    namen = [n for n in ("enabled", "time", "sensitivity", "unit")  # vorhandene Eigenschaften für die Diagnose ermitteln
             if hasattr(wd, n)]
    print("watchdog -> hasattr: " + ", ".join(namen))

    # Zeit, Empfindlichkeit und Aktivierung setzen (mit mehreren Wertvarianten)
    ok_zeit = setze_mit_varianten(
        wd, "time", [str(zeit_ms), "T#" + str(zeit_ms) + "ms", zeit_ms])
    ok_emp = setze_mit_varianten(
        wd, "sensitivity", [str(empfindlichkeit), empfindlichkeit])
    ok_akt = setze_mit_varianten(wd, "enabled", [True])
    if not (ok_zeit and ok_emp and ok_akt):
        raise RuntimeError("Watchdog konnte nicht vollstaendig gesetzt werden.")
    print("Watchdog: enabled=" + str(wd.enabled) + ", time=" + str(wd.time) +  # Kontrollausgabe der gesetzten Werte
          ", sensitivity=" + str(wd.sensitivity))


def erstelle_task(application, taskkonfig, name, prioritaet, typ, intervall,
                  programme, watchdog):
    """Legt eine Task an, setzt Typ/Intervall/Watchdog und fuegt die Programmaufrufe hinzu."""
    print("--- Task '" + name + "' ---")

    # Task anlegen, falls nicht vorhanden
    task = finde_kind(taskkonfig, name)
    if task is None:
        task = taskkonfig.create_task(name)  # neue Task anlegen
        if task is None:  # ggf. erneut suchen
            task = finde_kind(taskkonfig, name)
    if task is None:
        raise RuntimeError("Task '" + name + "' nicht gefunden.")
    print("Task vorhanden: " + task.get_name())

    # Typ der Task (zuerst, da er die uebrigen Eigenschaften beeinflusst)
    gesetzt = False  # Merker: Typ erfolgreich gesetzt
    for wert in finde_kind_of_task(typ):  # alle Kandidatenwerte probieren
        try:
            task.kind_of_task = wert  # Typ der Task setzen
            print("Gesetzt: kind_of_task = " + str(wert))
            gesetzt = True
            break  # ersten funktionierenden Wert verwenden
        except Exception as fehler:
            print("kind_of_task mit " + str(wert) + " fehlgeschlagen: " + str(fehler))
    if not gesetzt:
        print("Warnung: kind_of_task = " + typ + " konnte nicht gesetzt werden.")

    # Prioritaet
    task.priority = prioritaet
    print("Gesetzt: priority = " + prioritaet)

    # Intervall (zyklisch) bzw. Ereignis (Ereignis/Externes)
    if intervall:
        if typ == "Cyclic":
            task.interval = intervall  # IEC-Zeit, z. B. T#20ms
            print("Gesetzt: interval = " + intervall)
        else:
            ok = False  # Merker: Ereignis erfolgreich gesetzt
            for attr in ("event", "interval"):  # mögliche Eigenschaftsnamen für das Ereignis
                if hasattr(task, attr):
                    ok = setze_mit_varianten(task, attr, [intervall])
                    if ok:
                        break
            if not ok:
                print("Warnung: Ereignis '" + intervall + "' konnte nicht gesetzt werden.")

    # Watchdog (nur wenn gewuenscht)
    if watchdog is not None:
        setze_watchdog(task, watchdog[0], watchdog[1])

    # Programmaufrufe hinzufuegen (Eigenschaft "pous" der Task)
    if programme and not hasattr(task, "pous"):
        raise RuntimeError("Task hat keine Eigenschaft 'pous'.")
    for programm in programme:
        # Pruefen, ob das Programm im Projekt existiert (nur Warnung)
        if finde_kind(application, programm) is None:
            print("Warnung: Programm '" + programm + "' wurde unter der Application "
                  "nicht gefunden. Der Aufruf kann fehlschlagen.")
        task.pous.add(programm)  # Programmaufruf zur Task hinzufügen
        print("Programmaufruf '" + programm + "' hinzugefuegt.")


# Hauptteil: wird beim Ausführen des Skripts in CODESYS abgearbeitet
projekt = projects.primary  # aktuell geöffnetes Projekt (Objekt 'projects' stellt CODESYS bereit)
if projekt is None:
    raise RuntimeError("Es ist kein Projekt geoeffnet.")

application = finde_application(projekt)  # Application im Projektbaum suchen
if application is None:
    zeige_projektbaum(projekt)  # Diagnose: Projektbaum ausgeben, damit man die Namen sieht
    raise RuntimeError("Application nicht gefunden.")
print("Application gefunden: " + application.get_name())

# Taskkonfiguration anlegen, falls nicht vorhanden
taskkonfig = finde_taskkonfig(application)  # vorhandene Taskkonfiguration ueber den Typ suchen
if taskkonfig is None:
    print("Keine Taskkonfiguration gefunden, lege eine neue an.")
    application.create_task_configuration()  # Taskkonfiguration neu anlegen
    taskkonfig = finde_taskkonfig(application)
if taskkonfig is None:
    raise RuntimeError("Taskkonfiguration konnte nicht angelegt werden.")
print("Taskkonfiguration vorhanden: " + taskkonfig.get_name())

# Alle Tasks erstellen
for name, prio, typ, intervall, programme, watchdog in TASKS:  # Daten aus der Liste TASKS entpacken
    erstelle_task(application, taskkonfig, name, prio, typ, intervall,
                  programme, watchdog)

projekt.save()  # Änderungen speichern
print("Projekt gespeichert. Fertig.")
'''


def erzeuge_skript(tasks, excel_name):
    zeilen, zusammenfassung = [], []  # zeilen = Task-Einträge, zusammenfassung = Übersicht für den Skriptkopf
    for name, prio, typ, intervall, programme, wd in tasks:  # jede Task als Eintrag für die Vorlage aufbereiten
        zeilen.append("    (%r, %r, %r, %r, %r, %r)," %
                      (name, prio, typ, intervall, programme, wd))
        detail = typ + (", " + intervall if intervall else "")  # detail = Typ und ggf. Intervall
        wd_text = "Watchdog aktiv (Zeit %d, Empfindlichkeit %d)" % wd if wd else "kein Watchdog"  # wd_text = Beschreibung des Watchdogs
        zusammenfassung.append('  - "%s": Prioritaet %s, %s, Programme: %s, %s' %
                               (name, prio, detail, ", ".join(programme) or "-", wd_text))
    # Platzhalter in der Vorlage durch die echten Werte ersetzen
    return (VORLAGE.replace("__EXCEL__", excel_name)
            .replace("__ZUSAMMENFASSUNG__", "\n".join(zusammenfassung))
            .replace("__TASKS__", "\n".join(zeilen)))


def erzeuge_datei(excel, ausgabename, ordner=None):
    """Liest die Excel-Datei und schreibt das Skript in 'ordner' (Standard: Ordner der Excel-Datei).
    Gibt den Pfad des erzeugten Skripts und die Anzahl der Tasks zurück."""
    ausgabename = ausgabename.strip()  # Leerzeichen am Rand entfernen
    if not ausgabename:
        raise ValueError("Bitte einen Namen für das Skript angeben.")
    if not ausgabename.lower().endswith(".py"):  # Endung .py ergänzen, falls sie fehlt
        ausgabename += ".py"
    if os.path.dirname(ausgabename):  # Name enthält bereits einen Pfad -> direkt verwenden
        ausgabe = ausgabename
    else:
        ordner = (ordner or "").strip() or os.path.dirname(os.path.abspath(excel))  # sonst Zielordner bzw. Ordner der Excel-Datei verwenden
        if not os.path.isdir(ordner):
            raise ValueError("Der Zielordner existiert nicht: %s" % ordner)
        ausgabe = os.path.join(ordner, ausgabename)
    tasks = lese_excel(excel)  # Excel-Datei lesen und prüfen
    with open(ausgabe, "w", encoding="utf-8") as f:  # Skript als UTF-8-Datei schreiben
        f.write(erzeuge_skript(tasks, os.path.basename(excel)))
    return ausgabe, len(tasks)


def oeffne_fenster():
    """Fenster zum Auswählen der Excel-Datei, des Skriptnamens und zum Erzeugen des Skripts."""
    # Fenster-Bibliothek (tkinter) erst bei Bedarf importieren
    import tkinter as tk
    from tkinter import filedialog, messagebox

    fenster = tk.Tk()  # Hauptfenster
    fenster.title("Generator für CODESYS-Task-Skripte")
    fenster.resizable(False, False)

    # Variablen der Eingabefelder und der Statuszeile
    excel_pfad = tk.StringVar()
    zielordner = tk.StringVar()
    skriptname = tk.StringVar(value="erstelle_taskkonfiguration.py")
    statustext = tk.StringVar(value="Bitte eine Excel-Datei auswählen.")

    def durchsuchen():  # Schaltfläche 'Durchsuchen': Excel-Datei auswählen
        pfad = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Alle Dateien", "*.*")])
        if pfad:
            excel_pfad.set(pfad)
            if not zielordner.get():  # Standardordner: der Ordner der Excel-Datei
                zielordner.set(os.path.dirname(os.path.abspath(pfad)))
            statustext.set("Excel-Datei ausgewählt.")

    def ordner_waehlen():  # Schaltfläche 'Durchsuchen': Zielordner auswählen
        startordner = zielordner.get() or (
            os.path.dirname(excel_pfad.get()) if excel_pfad.get() else "")
        ordner = filedialog.askdirectory(
            title="Zielordner auswählen", initialdir=startordner or None)
        if ordner:
            zielordner.set(os.path.normpath(ordner))

    def erzeugen():  # Schaltfläche 'Skript erzeugen'
        if not excel_pfad.get():
            messagebox.showwarning("Excel-Datei fehlt", "Bitte zuerst die Excel-Datei auswählen.")
            return
        try:
            ausgabe, n = erzeuge_datei(
                excel_pfad.get(), skriptname.get(), zielordner.get())
        except ExcelFehler as fehler:  # Fehler in den Excel-Daten
            statustext.set("Skript nicht erzeugt: Die Excel-Datei enthält Fehler.")
            messagebox.showerror(
                "Fehler in der Excel-Datei - Skript wurde nicht erzeugt",
                "Bitte Folgendes in der Excel-Datei korrigieren und erneut versuchen:\n\n" + str(fehler))
            return
        except Exception as fehler:  # unerwarteter Fehler
            statustext.set("Fehler.")
            messagebox.showerror("Fehler", str(fehler))
            return
        statustext.set("Skript erzeugt (%d Tasks)." % n)
        messagebox.showinfo("Fertig", "Skript mit %d Tasks erzeugt:\n%s" % (n, ausgabe))

    # Aufbau der Oberfläche: Beschriftungen, Eingabefelder und Schaltflächen
    rahmen = tk.Frame(fenster, padx=12, pady=12)
    rahmen.pack()

    tk.Label(rahmen, text="Excel-Datei:").grid(row=0, column=0, sticky="w")
    tk.Entry(rahmen, textvariable=excel_pfad, width=55).grid(row=1, column=0, padx=(0, 6))
    tk.Button(rahmen, text="Durchsuchen...", command=durchsuchen).grid(row=1, column=1)

    tk.Label(rahmen, text="Ordner, in dem das Skript gespeichert wird:").grid(
        row=2, column=0, sticky="w", pady=(10, 0))
    tk.Entry(rahmen, textvariable=zielordner, width=55).grid(row=3, column=0, padx=(0, 6))
    tk.Button(rahmen, text="Durchsuchen...", command=ordner_waehlen).grid(row=3, column=1)

    tk.Label(rahmen, text="Name des erzeugten Skripts:").grid(
        row=4, column=0, sticky="w", pady=(10, 0))
    tk.Entry(rahmen, textvariable=skriptname, width=55).grid(row=5, column=0, padx=(0, 6))

    tk.Button(rahmen, text="Skript erzeugen", command=erzeugen, width=20).grid(
        row=6, column=0, columnspan=2, pady=(14, 6))
    tk.Label(rahmen, textvariable=statustext, fg="gray30").grid(row=7, column=0, columnspan=2)

    fenster.mainloop()  # Fenster anzeigen, bis es geschlossen wird


def main():
    # Mit Argumenten läuft das Programm über die Kommandozeile; ohne Argumente öffnet sich das Fenster.
    if len(sys.argv) < 2:
        oeffne_fenster()
        return
    excel = sys.argv[1]  # 1. Argument: Excel-Datei
    name = sys.argv[2] if len(sys.argv) > 2 else "erstelle_taskkonfiguration.py"  # 2. Argument (optional): Name des Skripts
    try:
        ausgabe, n = erzeuge_datei(excel, name)
    except ExcelFehler as fehler:
        print("Skript nicht erzeugt. Fehler in der Excel-Datei:\n" + str(fehler))
        sys.exit(1)  # Programm mit Fehlercode beenden
    print("Skript erzeugt: %s (%d Tasks)" % (ausgabe, n))


if __name__ == "__main__":  # nur beim direkten Start des Programms ausführen
    main()
