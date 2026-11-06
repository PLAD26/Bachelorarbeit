# -*- coding: utf-8 -*-
from scriptengine import *
import os
import glob

# Moegliche Namen der Taskkonfiguration (sprachabhaengig)
TASKKONFIG_NAMEN = ("Taskkonfiguration", "Task Configuration", "TaskConfiguration")

# True = Python-Skripte aus dem Ordner "python" werden so kompiliert, dass Texte mit Umlauten
# (ä, ö, ü, ß) als Unicode ankommen. Bei Problemen mit einem Skript auf False setzen.
UNICODE_TEXTE = True


def find_all_files(folder_path, pattern, required=True):
    """Gibt eine sortierte Liste aller Dateien in einem Ordner zurueck, die zum Muster passen (z. B. "*.xml").
    Wenn required=True und der Ordner nicht existiert oder leer ist, wird ein Fehler ausgeloest.
    Wenn required=False und der Ordner nicht existiert oder leer ist, wird eine leere Liste ohne Fehler zurueckgegeben."""
    if not os.path.isdir(folder_path):
        if required:
            raise IOError("Der Ordner existiert nicht: {0}".format(folder_path))
        else:
            print("Hinweis: der optionale Ordner existiert nicht und wird uebersprungen: {0}".format(folder_path))
            return []

    dateien = sorted(glob.glob(os.path.join(folder_path, pattern)))

    if not dateien and required:
        raise IOError("Es wurde keine Datei ({0}) gefunden in: {1}".format(pattern, folder_path))

    return dateien


def lese_quelltext(py_path):
    """Liest eine Python-Datei als Bytes (funktioniert in Python 2 und 3)
    und entfernt ein eventuelles UTF-8-BOM am Dateianfang."""
    with open(py_path, "rb") as f:
        source = f.read()
    if source.startswith(b"\xef\xbb\xbf"):
        source = source[3:]
    return source


def ist_steuerungsskript(py_path):
    """Prueft, ob ein Python-Skript selbst eine Steuerung wie Codesys_Control_Win_V3 anlegt
    (z. B. "Mischanlage Steurung"), unabhaengig vom Dateinamen.
    Erkennung: Im Quelltext kommt der Name einer CODESYS-Steuerung vor
    ("Codesys_Control_Win_V3", "CODESYS Control Win V3", "Codesys Control RTE" usw.)."""
    try:
        text = lese_quelltext(py_path).decode("utf-8", "ignore").lower()
    except Exception as e:
        print("Hinweis: Skript konnte nicht geprueft werden ({0}): {1}".format(py_path, e))
        return False

    # Unterstriche wie Leerzeichen behandeln, damit beide Schreibweisen erkannt werden
    text = text.replace("_", " ")
    return "codesys control" in text


def kompiliere_skript(source, py_path):
    """Kompiliert den Quelltext eines Skripts.
    Mit UNICODE_TEXTE=True werden alle Texte als Unicode behandelt, damit Umlaute
    (z. B. "Position.Höhe") in CODESYS richtig ankommen und nicht als "HÃ¶he"."""
    if UNICODE_TEXTE:
        try:
            import __future__
            text = source.decode("utf-8")  # Quelltext als Unicode lesen
            # Kodierungszeile (# -*- coding: ... -*-) entschaerfen, sie ist bei Unicode-Quelltext nicht erlaubt
            zeilen = text.split(u"\n")
            for i in range(min(2, len(zeilen))):
                if zeilen[i].lstrip().startswith(u"#") and u"coding" in zeilen[i]:
                    zeilen[i] = u"#"
            text = u"\n".join(zeilen)
            return compile(text, py_path, "exec", __future__.unicode_literals.compiler_flag)
        except Exception as e:
            print("Hinweis: Unicode-Kompilierung nicht moeglich ({0}), Standard wird verwendet.".format(e))
    return compile(source, py_path, "exec")


def fuehre_python_skript_aus(py_path):
    """Fuehrt ein Python-Skript mit denselben Objekten aus (projects, prj, apl ...).
    Fehler im Skript werden ausgegeben, brechen aber das Hauptprogramm nicht ab."""
    print("Fuehre Python-Skript aus: {0}".format(py_path))
    script_globals = dict(globals())        # gleiche Objekte (projects, prj, apl ...)
    script_globals["__file__"] = py_path    # eigener Pfad fuer das ausgefuehrte Skript
    script_globals["__name__"] = "__main__"
    try:
        source = lese_quelltext(py_path)
        code = kompiliere_skript(source, py_path)
        exec(code, script_globals)
    except Exception as e:
        print("Fehler im Python-Skript {0}: {1}".format(py_path, e))


def hole_applikation(projekt):
    """Gibt die aktive Applikation zurueck oder None, wenn es keine gibt."""
    try:
        return projekt.active_application
    except Exception:
        return None


def steuerung_erstellen(projekt):
    """Legt die Standard-Steuerung Codesys_Control_Win_V3 im Projekt an."""
    projekt.add("Codesys_Control_Win_V3", type=4096, id='0000 0004', version='3.5.21.20')


def entferne_schreibschutz(pfad):
    """Entfernt das Schreibschutz-Attribut einer Datei, falls gesetzt.
    CODESYS oeffnet schreibgeschuetzte Projekte im Nur-Lesen-Modus und kann sie dann nicht speichern."""
    try:
        if os.path.isfile(pfad) and not os.access(pfad, os.W_OK):
            import stat
            os.chmod(pfad, stat.S_IWRITE | stat.S_IREAD)  # Schreibrechte setzen
            print("Schreibschutz entfernt: {0}".format(pfad))
            return True
    except Exception as e:
        print("Hinweis: Schreibschutz konnte nicht entfernt werden ({0}): {1}".format(pfad, e))
    return False


def finde_taskkonfig(applikation):
    """Sucht eine vorhandene Taskkonfiguration unter der Applikation:
    zuerst ueber den Objekttyp, danach ueber den Namen (deutsch/englisch)."""
    # 1) Typpruefung (bevorzugt)
    for objekt in applikation.get_children(True):  # True = gesamten Baum unter der Applikation durchsuchen
        try:
            if objekt.is_task_configuration:
                return objekt
        except Exception:
            pass  # Objekt hat die Eigenschaft nicht -> naechstes Objekt
    # 2) Fallback: Namensvergleich
    for objekt in applikation.get_children(True):
        try:
            if objekt.get_name() in TASKKONFIG_NAMEN:
                return objekt
        except Exception:
            pass
    return None


def finde_kind(eltern, name):
    """Sucht ein direktes Kind-Objekt per Name."""
    for kind in eltern.get_children(False):  # nur direkte Kind-Objekte durchsuchen
        if kind.get_name() == name:
            return kind
    return None


def zeige_applikation_inhalt(applikation):
    """Diagnose: gibt alle Objekte unter der Applikation mit ihren Typ-Merkmalen aus."""
    print("Inhalt von '{0}':".format(applikation.get_name()))
    for objekt in applikation.get_children(True):
        try:
            print("    - {0} | taskkonfig={1}".format(
                objekt.get_name(), getattr(objekt, "is_task_configuration", False)))
        except Exception as e:
            print("    - (Fehler: {0})".format(e))


def hole_oder_erstelle_taskkonfig(applikation):
    """Gibt die vorhandene Taskkonfiguration zurueck oder legt sie an, falls keine existiert.
    Gibt None zurueck, wenn das Anlegen nicht moeglich ist."""
    taskconfig = finde_taskkonfig(applikation)
    if taskconfig is not None:
        print("Vorhandene Taskkonfiguration wird verwendet: {0}".format(taskconfig.get_name()))
        return taskconfig

    print("Keine Taskkonfiguration gefunden, lege eine neue an.")
    try:
        applikation.create_task_configuration()
    except Exception as e:
        # Typischer Fehler: "Elternobjekt akzeptiert kein Taskkonfigurationsobjekt"
        print("Anlegen der Taskkonfiguration fehlgeschlagen: {0}".format(e))
        zeige_applikation_inhalt(applikation)  # Diagnose
        return None
    return finde_taskkonfig(applikation)


# Verzeichnis des aktuellen Skripts
scriptfile = os.path.dirname(__file__)
# Verzeichnis des Projektspeicherorts
projectfile = os.path.join(scriptfile, "automatische Projektgenerierung")

# Ordner
xml_folder = os.path.join(scriptfile, "xml")        # KI-generierte PLCopenXML-Dateien (Pflicht)
fb_folder = os.path.join(scriptfile, "fb")          # Externe Funktionsbausteine als XML (optional)
python_folder = os.path.join(scriptfile, "python")  # Weitere Python-Skripte, z. B. fuer das HMI (optional)

# Liste aller gefundenen Dateien in jedem Ordner
xml_files = find_all_files(xml_folder, "*.xml", required=True)
fb_files = find_all_files(fb_folder, "*.xml", required=False)
py_files = find_all_files(python_folder, "*.py", required=False)

# Python-Skripte aufteilen: Steuerungs-Skripte (legen die Steuerung selbst an) und alle uebrigen
steuerungs_skripte = [p for p in py_files if ist_steuerungsskript(p)]
weitere_skripte = [p for p in py_files if p not in steuerungs_skripte]

# Ordner der Bibliotheken: jede .library-Datei in diesem Ordner wird installiert (Name beliebig)
lib_folder = os.path.join(scriptfile, "lib")
lib_files = find_all_files(lib_folder, "*.library", required=True)

# Pruefen, ob ein Projekt geoeffnet ist
if projects.primary:
    projects.primary.close()


# Eine Projektdatei aus einem frueheren Lauf darf nicht schreibgeschuetzt sein
entferne_schreibschutz(projectfile + ".project")

# Projekt erstellen
prj = projects.create(projectfile, True)
if projects.primary:

    # 0) Steuerung anlegen
    if steuerungs_skripte:
        # Ein Skript wie "Mischanlage Steurung" wurde gefunden: Es legt die Steuerung selbst an,
        # deshalb wird Codesys_Control_Win_V3 hier NICHT erstellt. Diese Skripte laufen als Erstes,
        # weil die Applikation fuer Bibliotheken, XML-Import und Taskkonfiguration vorhanden sein muss.
        print("Steuerungs-Skript gefunden, Codesys_Control_Win_V3 wird nicht separat erstellt.")
        for sk_path in steuerungs_skripte:
            fuehre_python_skript_aus(sk_path)

        apl = hole_applikation(prj)
        if apl is None:
            # Sicherheitsnetz: das Skript hat keine Steuerung mit Applikation hinterlassen
            print("Warnung: Das Steuerungs-Skript hat keine Applikation erzeugt. "
                  "Codesys_Control_Win_V3 wird als Ersatz erstellt.")
            steuerung_erstellen(prj)
    else:
        # Kein aehnliches Skript gefunden: Standard-Steuerung selbst erstellen
        print("Kein Steuerungs-Skript gefunden, Codesys_Control_Win_V3 wird erstellt.")
        steuerung_erstellen(prj)

    # Zugriff auf die aktive Applikation
    apl = prj.active_application

    # Installieren und Hinzufuegen aller Bibliotheken (*.library) aus dem Ordner "lib"
    managedLibs = []
    library_manager = apl.get_library_manager()
    for lib_path in lib_files:
        print("Installiere Bibliothek: {0}".format(lib_path))
        managedLibs.append(librarymanager.install_library(lib_path, librarymanager.repositories[0], True))

    for lib in managedLibs:
        library_manager.add_library(lib)
        libname = lib.displayname
        library_manager.remove_library(libname)
        library_manager.add_placeholder(libname, libname)

    lib_name_standard = "Standard"
    lib_placeholder_standard = "Standard, 3.5.18.0 (System)"
    library_manager.add_placeholder(lib_name_standard, lib_placeholder_standard)

    # 1) Import der externen Funktionsbausteine aus XML-Dateien (alle Dateien im Ordner "fb", falls vorhanden)
    #    Sie werden zuerst importiert, weil der KI-generierte Code sie verwendet.
    for fb_path in fb_files:
        print("Importiere FB: {0}".format(fb_path))
        apl.import_xml(
            dataOrPath=fb_path,
            conflictResolve=ConflictResolve.Replace,
            import_folder_structure=False
        )

    # 2) Import der KI-generierten PLCopenXML-Dateien (alle Dateien im Ordner "xml")
    #    Es wird nicht mehr nach einem bestimmten Programm (z. B. PLC_PRG) gesucht.
    for xml_path in xml_files:
        print("Importiere XML: {0}".format(xml_path))
        apl.import_xml(
            dataOrPath=xml_path,
            conflictResolve=ConflictResolve.Replace,
            import_folder_structure=False
        )

    # 3) Taskkonfiguration verwenden bzw. erstellen
    #    Die Taskkonfiguration kann schon vorhanden sein (Steuerungs-Skript, Vorlage, XML-Import).
    #    Eine Applikation akzeptiert nur EINE Taskkonfiguration, daher zuerst suchen, dann anlegen.
    taskconfig = hole_oder_erstelle_taskkonfig(apl)

    if taskconfig is None:
        print("Warnung: Es steht keine Taskkonfiguration zur Verfuegung, MainTask wird uebersprungen.")
    else:
        # MainTask verwenden bzw. erstellen
        main_task_cyclic = finde_kind(taskconfig, 'MainTask')
        if main_task_cyclic is None:
            main_task_cyclic = taskconfig.create_task('MainTask')
            print("MainTask wurde erstellt.")
        else:
            print("Vorhandene MainTask wird angepasst.")

        # Task-Eigenschaften konfigurieren
        main_task_cyclic.kind_of_task = KindOfTask.Cyclic  # Zyklische Task
        main_task_cyclic.watchdog.enabled = False          # Watchdog deaktivieren
        main_task_cyclic.interval = "20"                   # Intervall: 20 ms
        main_task_cyclic.priority = "1"                    # Prioritaet: 1
        # Hinweis: Es wird kein Programm automatisch zur Task hinzugefuegt.

    # 4) Projekt speichern, schliessen und neu oeffnen,
    #    damit die Bibliotheken (Placeholder) vor dem Ausfuehren der Python-Skripte aufgeloest sind
    print("Speichere Projekt und oeffne es neu ...")
    prj.save()
    proj_path = prj.path
    prj.close()
    # Vor dem erneuten Oeffnen pruefen, ob die Projektdatei schreibgeschuetzt ist
    entferne_schreibschutz(proj_path)
    if not os.access(os.path.dirname(proj_path), os.W_OK):
        print("Warnung: Der Projektordner ist nicht beschreibbar: {0}".format(os.path.dirname(proj_path)))
    prj = projects.open(proj_path)
    apl = prj.active_application

    # 5) Ausfuehren der uebrigen Python-Skripte im Ordner "python" (z. B. HMI), falls vorhanden
    #    (das Steuerungs-Skript wurde bereits in Schritt 0 ausgefuehrt).
    #    prj und apl zeigen jetzt auf das neu geoeffnete Projekt.
    for py_path in weitere_skripte:
        fuehre_python_skript_aus(py_path)

    print("Projekt erstellt in: {0}".format(projectfile))
    print("Hinweis: das Projekt wurde NICHT automatisch gespeichert.")
