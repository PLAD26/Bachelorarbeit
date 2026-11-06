#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aufruf:
  python3 gsdml_leser.py                                    # interaktive Oberfläche (Tkinter)
  python3 gsdml_leser.py ordner_oder_datei                  # Kommandozeile: Ordner / einzelne GSDML-Datei
  python3 gsdml_leser.py ordner_oder_datei --ohne-optionale # ohne optionale Module

Standard: optionale Module werden angezeigt (in der Oberfläche ist das Häkchen gesetzt).
"""
import argparse
import os
import sys
import xml.etree.ElementTree as ET


def lokal(tag):
    """Namensraum vom XML-Tag entfernen: '{ns}Name' -> 'Name'."""
    return tag.split('}')[-1]


def kinder(elem, name):
    """Alle direkten Kinder mit dem lokalen Namen 'name'."""
    return [k for k in elem if lokal(k.tag) == name]


def finde(elem, name):
    """Alle Nachfahren mit dem lokalen Namen 'name'."""
    return [k for k in elem.iter() if lokal(k.tag) == name]


def erster(elem, name):
    liste = kinder(elem, name)
    return liste[0] if liste else None


class Gsdml:
    def __init__(self, pfad):
        # GSDML ist oft iso-8859-1: ElementTree liest die Deklaration selbst
        self.wurzel = ET.parse(pfad).getroot()
        self.texte = self._texte_laden()

    def _texte_laden(self):
        """ExternalTextList -> {TextId: Value}; PrimaryLanguage hat Vorrang."""
        texte = {}
        for ext in finde(self.wurzel, 'ExternalTextList'):
            for sprache in ext:
                for t in kinder(sprache, 'Text'):
                    # PrimaryLanguage überschreibt, Fremdsprachen nur als Notlösung
                    if lokal(sprache.tag) == 'PrimaryLanguage' or t.get('TextId') not in texte:
                        texte[t.get('TextId')] = t.get('Value')
        return texte

    def name(self, elem):
        """Anzeigename eines Elements: ModuleInfo/Name (Value oder TextId)."""
        info = erster(elem, 'ModuleInfo')
        n = erster(info, 'Name') if info is not None else None
        if n is None:
            return elem.get('ID')
        if n.get('TextId'):
            return self.texte.get(n.get('TextId'), n.get('TextId'))
        return n.get('Value')

    def text(self, elem):
        """Text eines DataItem / Elements über TextId."""
        tid = elem.get('TextId')
        return self.texte.get(tid, tid) if tid else None

    def hersteller(self):
        v = finde(self.wurzel, 'VendorName')
        return v[0].get('Value') if v else None

    # ---------- DAPs (= die "Geräte") ----------
    def daps(self):
        return finde(self.wurzel, 'DeviceAccessPointItem')

    # ---------- Module ----------
    def module_dict(self):
        """{ModuleItem-ID: Element}"""
        return {m.get('ID'): m for m in finde(self.wurzel, 'ModuleItem')}

    def submodule_io(self, modul):
        """Untermodule mit zyklischen E/A-Daten eines ModuleItem."""
        ergebnis = []
        for sub in finde(modul, 'VirtualSubmoduleItem'):
            io = erster(sub, 'IOData')
            if io is None:
                continue
            for richtung, tag in (('Eingang', 'Input'), ('Ausgang', 'Output')):
                bloecke = kinder(io, tag)
                for b in bloecke:
                    for di in kinder(b, 'DataItem'):
                        ergebnis.append({
                            'submodul': self.name(sub),
                            'kanal': self.text(di),
                            'richtung': richtung,
                            'typ': di.get('DataType'),
                            'bits': di.get('UseAsBits'),
                        })
        return ergebnis

    def geraet(self, dap_id=None, alle=True):
        """Fasst ein Gerät zusammen: DAP-Name + Modulliste + E/A.
        alle=True (Standard): auch optionale (nicht fest gesteckte) Module."""
        daps = self.daps()
        dap = next((d for d in daps if d.get('ID') == dap_id), daps[0]) if dap_id else daps[0]
        module = self.module_dict()
        refs = []
        use = erster(dap, 'UseableModules')
        if use is not None:
            for r in kinder(use, 'ModuleItemRef'):
                refs.append(r)
        # fest gesteckte Module nach Slot sortieren
        def slotnr(r):
            s = r.get('FixedInSlots')
            return int(s.split()[0].split('..')[0]) if s else 10**6
        fest = sorted([r for r in refs if r.get('FixedInSlots')], key=slotnr)
        frei = [r for r in refs if not r.get('FixedInSlots')]
        auswahl = fest + (frei if alle else [])
        liste = []
        for r in auswahl:
            m = module.get(r.get('ModuleItemTarget'))
            if m is None:
                continue
            liste.append({
                'slot': r.get('FixedInSlots') or r.get('AllowedInSlots') or '',
                'fest': bool(r.get('FixedInSlots')),
                'technisch': self.name(m),
                'id': m.get('ID'),
                'io': self.submodule_io(m),
            })
        return {
            'dap_id': dap.get('ID'),
            'technisch': self.name(dap),
            'hersteller': self.hersteller(),
            'module': liste,
            'weitere_daps': [d.get('ID') for d in daps if d is not dap],
            'nicht_fest': len(frei) if not alle else 0,
        }


def bericht(g):
    """Bericht als Textzeilen (wird auf dem Bildschirm und in die Datei geschrieben)."""
    z = []
    z.append('Hersteller        : %s' % g['hersteller'])
    z.append('Gerät (DAP-Name)  : %s  [ID: %s]' % (g['technisch'], g['dap_id']))
    if g['weitere_daps']:
        z.append('Weitere DAPs      : %s (mit --dap wählen)' % ', '.join(g['weitere_daps']))
    z.append('')
    z.append('Module:')
    for m in g['module']:
        z.append('  Slot %-4s %s%s' % (m['slot'], m['technisch'], '' if m['fest'] else '  (optional)'))
    if g['nicht_fest']:
        z.append('  (+ %d weitere steckbare Module, ohne --ohne-optionale anzeigen)' % g['nicht_fest'])
    z.append('')
    z.append('E/A-Kanäle:')
    for m in g['module']:
        for k in m['io']:
            sub = m['technisch'] if k['submodul'] == m['technisch'] else '%s/%s' % (m['technisch'], k['submodul'])
            z.append('  %-30s | %-26s | %-8s %s' % (sub, k['kanal'], k['richtung'], k['typ']))
    return z


def an_excel_anhaengen(g, xlsx, geraetename):
    import openpyxl
    wb = openpyxl.load_workbook(xlsx)
    ws = wb.worksheets[0]
    zeile = ws.max_row + 1
    ws.cell(zeile, 1, geraetename)
    ws.cell(zeile, 2, g['technisch'])
    for m in g['module']:
        zeile += 1
        # Spalte B = gewünschter Name (Vorschlag = Modulname, bitte anpassen),
        # Spalte C = technischer Modulname aus der GSDML
        ws.cell(zeile, 2, m['technisch'])
        ws.cell(zeile, 3, m['technisch'])
    wb.save(xlsx)
    print('Gerät "%s" mit %d Modulen an %s angehängt.' % (geraetename, len(g['module']), xlsx))


def gsdml_dateien(pfad):
    """Datei -> [Datei]; Ordner -> alle .xml/.gsdml-Dateien darin (alphabetisch)."""
    if os.path.isdir(pfad):
        return sorted(os.path.join(pfad, f) for f in os.listdir(pfad)
                      if f.lower().endswith(('.xml', '.gsdml')))
    return [pfad]


def bericht_erzeugen(pfad, dap=None, alle=True):
    """Liest Datei/Ordner und liefert (Berichtstext, Anzahl Dateien)."""
    dateien = gsdml_dateien(pfad)
    ausgabe = []
    for d in dateien:
        ausgabe.append('=' * 70)
        ausgabe.append('Datei: %s' % os.path.basename(d))
        ausgabe.append('=' * 70)
        try:
            ausgabe += bericht(Gsdml(d).geraet(dap, alle))
        except Exception as e:  # eine kaputte Datei soll die anderen nicht stoppen
            ausgabe.append('FEHLER beim Lesen: %s' % e)
        ausgabe.append('')
    return '\n'.join(ausgabe), len(dateien)


def gui():
    """Interaktive Oberfläche: Ordner wählen, Dateiname und Speicherort ändern."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    fenster = tk.Tk()
    fenster.title('GSDML-Leser - Bericht für Steuerungskonfiguration')
    fenster.geometry('900x650')

    var_ordner = tk.StringVar(value=os.path.join(os.getcwd(), 'GSDML'))
    var_zielordner = tk.StringVar(value=os.getcwd())
    var_dateiname = tk.StringVar(value='Bericht.txt')
    var_alle = tk.BooleanVar(value=True)   # optionale Module standardmäßig aktiviert

    def zeile(row, beschriftung, var, befehl=None):
        tk.Label(fenster, text=beschriftung, anchor='w').grid(row=row, column=0, sticky='w', padx=8, pady=4)
        tk.Entry(fenster, textvariable=var, width=70).grid(row=row, column=1, sticky='ew', padx=4)
        if befehl:
            tk.Button(fenster, text='Durchsuchen...', command=befehl).grid(row=row, column=2, padx=8)

    def waehle_ordner():
        d = filedialog.askdirectory(title='Ordner mit den GSDML-Dateien wählen',
                                    initialdir=var_ordner.get() if os.path.isdir(var_ordner.get()) else os.getcwd())
        if d:
            var_ordner.set(d)

    def waehle_ziel():
        d = filedialog.askdirectory(title='Speicherort des Berichts wählen',
                                    initialdir=var_zielordner.get() if os.path.isdir(var_zielordner.get()) else os.getcwd())
        if d:
            var_zielordner.set(d)

    zeile(0, 'Ordner mit GSDML-Dateien:', var_ordner, waehle_ordner)
    zeile(1, 'Speicherort des Berichts:', var_zielordner, waehle_ziel)
    zeile(2, 'Name des Berichts:', var_dateiname)
    tk.Checkbutton(fenster, text='Auch optionale Module anzeigen (nicht fest gesteckt)',
                   variable=var_alle).grid(row=3, column=1, sticky='w')

    anzeige = scrolledtext.ScrolledText(fenster, wrap='none', font=('Courier', 9))
    anzeige.grid(row=5, column=0, columnspan=3, sticky='nsew', padx=8, pady=8)
    fenster.columnconfigure(1, weight=1)
    fenster.rowconfigure(5, weight=1)

    def lesen():
        pfad = var_ordner.get()
        if not os.path.exists(pfad):
            messagebox.showerror('Fehler', 'Ordner nicht gefunden:\n%s' % pfad)
            return None
        text, n = bericht_erzeugen(pfad, alle=var_alle.get())
        if n == 0:
            messagebox.showwarning('Hinweis', 'Keine .xml/.gsdml-Dateien im Ordner gefunden.')
            return None
        anzeige.delete('1.0', 'end')
        anzeige.insert('end', text)
        return text

    def speichern():
        text = lesen()
        if text is None:
            return
        name = var_dateiname.get().strip() or 'Informe_GSDML.txt'
        if not name.lower().endswith('.txt'):
            name += '.txt'
        ziel = os.path.join(var_zielordner.get(), name)
        try:
            os.makedirs(var_zielordner.get(), exist_ok=True)
            with open(ziel, 'w', encoding='utf-8') as f:
                f.write(text)
            messagebox.showinfo('Fertig', 'Bericht gespeichert in:\n%s' % ziel)
        except OSError as e:
            messagebox.showerror('Fehler beim Speichern', str(e))

    knoepfe = tk.Frame(fenster)
    knoepfe.grid(row=4, column=0, columnspan=3, pady=6)
    tk.Button(knoepfe, text='Vorschau (nur anzeigen)', command=lesen).pack(side='left', padx=6)
    tk.Button(knoepfe, text='Bericht erzeugen und speichern', command=speichern).pack(side='left', padx=6)

    fenster.mainloop()


if __name__ == '__main__':
    if len(sys.argv) == 1:
        gui()          # ohne Argumente: interaktive Oberfläche
        sys.exit()

    # Kommandozeile
    ap = argparse.ArgumentParser()
    ap.add_argument('pfad', help='GSDML-Datei oder Ordner')
    ap.add_argument('--dap')
    ap.add_argument('--ohne-optionale', dest='alle', action='store_false',
                    help='optionale Module (nicht fest gesteckt) nicht anzeigen')
    ap.set_defaults(alle=True)
    ap.add_argument('--excel', help='Excel, an das das Gerät angehängt wird (nur bei einer Datei)')
    ap.add_argument('--name', help='gewünschter Gerätename (Spalte A)')
    ap.add_argument('--ausgabe', default='Informe_GSDML.txt', help='Berichtsdatei')
    a = ap.parse_args()
    text, n = bericht_erzeugen(a.pfad, a.dap, a.alle)
    if n == 0:
        sys.exit('Keine GSDML-Dateien gefunden in: %s' % a.pfad)
    if a.excel and n == 1:
        if not a.name:
            sys.exit('--name fehlt')
        an_excel_anhaengen(Gsdml(gsdml_dateien(a.pfad)[0]).geraet(a.dap, a.alle), a.excel, a.name)
    print(text)
    with open(a.ausgabe, 'w', encoding='utf-8') as f:
        f.write(text)
    print('Bericht gespeichert in: %s' % a.ausgabe)
