# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

section = doc.sections[0]
section.page_width    = Cm(21)
section.page_height   = Cm(29.7)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)
section.top_margin    = Cm(2)
section.bottom_margin = Cm(2)

DARK_GREEN  = RGBColor(0x1B, 0x55, 0x2E)   # BW-Gruen (Bildungsplan-Farbe)
MID_GREEN   = RGBColor(0x2E, 0x8B, 0x57)
LIGHT_GREEN = "D6EDD9"
ACCENT      = RGBColor(0xC8, 0x5A, 0x00)   # BW-Orange
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def cell_borders(cell, color="CCCCCC", sz=4):
    tc    = cell._tc
    tcPr  = tc.get_or_add_tcPr()
    tcBdr = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + side)
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBdr.append(el)
    tcPr.append(tcBdr)

def bottom_border_para(para, color="2E8B57", sz="6"):
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    sz)
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color)
    pBdr.append(bot)
    pPr.append(pBdr)

def add_heading(doc, text, level=1, color=None, size=14):
    if color is None:
        color = DARK_GREEN
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold           = True
    run.font.size      = Pt(size)
    run.font.color.rgb = color
    if level == 1:
        bottom_border_para(p)
    return p

def add_bullet(doc, text, bold_start=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    if bold_start:
        r1 = p.add_run(bold_start + ": ")
        r1.bold      = True
        r1.font.size = Pt(10.5)
        r2 = p.add_run(text)
        r2.font.size = Pt(10.5)
    else:
        r = p.add_run(text)
        r.font.size = Pt(10.5)
    return p

def add_normal(doc, text, bold=False, italic=False, size=10.5, color=None, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(space_after)
    r = p.add_run(text)
    r.bold      = bold
    r.italic    = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p

def spacer(doc, pts=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(pts)

# ==============================================================
# TITLE BLOCK
# ==============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(2)
r = p.add_run("UNTERRICHTSENTWURF")
r.bold = True; r.font.size = Pt(9)
r.font.color.rgb = RGBColor(0x88, 0x88, 0x88); r.font.all_caps = True

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(6)
r = p.add_run("Dateien und Ordner verstehen & nutzen")
r.bold = True; r.font.size = Pt(22); r.font.color.rgb = DARK_GREEN

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(4)
r = p.add_run("Informatik und Medienbildung  |  Klasse 5")
r.font.size = Pt(13); r.italic = True; r.font.color.rgb = MID_GREEN

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(14)
r = p.add_run("Bildungsplan 2016 Baden-Wuerttemberg")
r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

# Info-table
info_data = [
    ("Fach",              "Informatik / Basiskurs Medienbildung (Klasse 5)"),
    ("Thema der Stunde",  "Dateien und Ordner verstehen und nutzen"),
    ("Vorwissen der SuS", "Anmelden am Rechner mit Schueleraccount; Grundlagen Word & PowerPoint"),
    ("Zeitumfang",        "45 Minuten (eine Unterrichtsstunde)"),
    ("Sozialform",        "Einzelarbeit / Partnerarbeit / Plenum"),
    ("Medien & Material", "PC/Laptop (Windows), Beamer/Whiteboard, Arbeitsblatt AB 1"),
    ("Bildungsplan-Bezug","Bildungsplan 2016 BW - Leitperspektive Medienbildung (MB);\n"
                          "Basiskurs Medienbildung Kl. 5/6; Leitperspektive Digitale Bildung"),
]

t_info = doc.add_table(rows=len(info_data), cols=2)
t_info.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(info_data):
    row = t_info.rows[i]
    set_cell_bg(row.cells[0], "E8F5EB")
    cell_borders(row.cells[0]); cell_borders(row.cells[1])
    p0 = row.cells[0].paragraphs[0]
    p0.paragraph_format.space_before = Pt(3); p0.paragraph_format.space_after = Pt(3)
    r0 = p0.add_run(k); r0.bold = True; r0.font.size = Pt(10); r0.font.color.rgb = DARK_GREEN
    p1 = row.cells[1].paragraphs[0]
    p1.paragraph_format.space_before = Pt(3); p1.paragraph_format.space_after = Pt(3)
    r1 = p1.add_run(v); r1.font.size = Pt(10)

spacer(doc, 8)

# ==============================================================
# 1. BEZUG ZUM BILDUNGSPLAN BW 2016
# ==============================================================
add_heading(doc, "1  Bezug zum Bildungsplan 2016 Baden-Wuerttemberg")

add_normal(doc,
    "Die Stunde orientiert sich am Bildungsplan 2016 Baden-Wuerttemberg. "
    "Dort ist die Medienbildung als faecheruebergreifende Leitperspektive (MB) verankert "
    "und wird durch den obligatorischen Basiskurs Medienbildung in den Klassen 5/6 konkretisiert. "
    "Ergaenzend werden Bezuege zum Bereich Informatik / Digitale Bildung hergestellt.",
    space_after=6)

# 1a Basiskurs Medienbildung
add_heading(doc, "Basiskurs Medienbildung (Kl. 5/6) - Pflichtbestandteil des Bildungsplans 2016",
            level=2, color=MID_GREEN, size=11)

add_normal(doc,
    "Der Basiskurs Medienbildung ist ein verpflichtender, faecheruebergreifender Kurs, "
    "der in den Klassen 5 und 6 von allen Schulen in BW durchgefuehrt wird. "
    "Diese Stunde deckt folgende Inhalte des Basiskurses ab:",
    space_after=4)

basiskurs = [
    ("Dateiverwaltung und Dateiorganisation",
     "Die SuS legen Ordner an, benennen Dateien sinnvoll und navigieren in Verzeichnisstrukturen. "
     "Sie speichern eigene Arbeitsergebnisse strukturiert und rufen sie gezielt wieder ab."),
    ("Grundfunktionen des Betriebssystems",
     "Die SuS bedienen den Datei-Explorer, verstehen den Aufbau des Dateisystems "
     "und unterscheiden zwischen Dateien und Ordnern."),
    ("Dateiformate und Anwendungen",
     "Die SuS ordnen gaengige Dateiendungen (.docx, .pptx, .jpg) den entsprechenden "
     "Anwendungsprogrammen zu und verstehen deren Funktion."),
]
for k, v in basiskurs:
    add_bullet(doc, v, bold_start=k)

spacer(doc, 4)

# 1b Leitperspektive Medienbildung
add_heading(doc, "Leitperspektive Medienbildung (MB) - Kompetenzbereiche",
            level=2, color=MID_GREEN, size=11)

add_normal(doc,
    "Die Leitperspektive Medienbildung des Bildungsplans 2016 BW gliedert sich in "
    "sechs Kompetenzbereiche. Folgende sind in dieser Stunde zentral:",
    space_after=4)

mb_bereiche = [
    ("MB 1 - Information und Wissen",
     "Die SuS koennen Dateien auf einem Speichermedium auffinden, benennen und strukturiert ablegen. "
     "Sie verstehen, dass Daten in einem hierarchischen Dateisystem organisiert sind."),
    ("MB 2 - Kommunikation und Kollaboration",
     "Die SuS speichern Arbeitsergebnisse so, dass sie gezielt abrufbar und weitergegeben werden koennen "
     "(z. B. Teilen ueber den Schueleraccount)."),
    ("MB 5 - Problemloesen und Modellieren",
     "Die SuS erkennen, dass eine sinnvolle Ordnerstruktur ein Modell fuer reale Ordnungssysteme ist, "
     "und uebertragen dieses Prinzip auf ihre eigene digitale Arbeitsumgebung."),
    ("MB 6 - Analysieren und Reflektieren",
     "Die SuS reflektieren, warum eine strukturierte Ablage fuer effizientes Arbeiten notwendig ist, "
     "und beurteilen verschiedene Benennungskonventionen."),
]
for k, v in mb_bereiche:
    add_bullet(doc, v, bold_start=k)

spacer(doc, 4)

# 1c Informatik / Digitale Bildung
add_heading(doc, "Bildungsplan 2016 BW - Informatik (Gymnasium, Kl. 5/6 / Propadeutik)",
            level=2, color=MID_GREEN, size=11)

add_normal(doc,
    "Sofern die Stunde im Rahmen eines Informatik-Kurses oder Wahlpflichtfachs stattfindet, "
    "gelten zusaetzlich folgende Bildungsplan-Bezuege:",
    space_after=4)

inf_bereiche = [
    ("Daten und Information",
     "Die SuS beschreiben, wie Daten als Dateien auf einem Datentraeger gespeichert und "
     "durch ein Dateisystem verwaltet werden. Sie unterscheiden Daten (Inhalt) und Metadaten "
     "(Dateiname, -groesse, -typ)."),
    ("Informatiksysteme",
     "Die SuS bedienen das Betriebssystem auf der Ebene der Dateiverwaltung und "
     "verstehen grundlegende Konzepte wie Verzeichnishierarchie und Dateipfad."),
]
for k, v in inf_bereiche:
    add_bullet(doc, v, bold_start=k)

spacer(doc, 6)

# ==============================================================
# 2. LERNZIELE
# ==============================================================
add_heading(doc, "2  Lernziele")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
r = p.add_run("Stundenziel (Grobziel)")
r.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = ACCENT

add_normal(doc,
    "Die Schuelerinnen und Schueler koennen selbststaendig eine sinnvolle Ordnerstruktur "
    "im Datei-Explorer anlegen, Dateien korrekt benennen und gezielt speichern sowie oeffnen "
    "(Basiskurs Medienbildung BW: Dateiverwaltung, MB 1 & MB 5).",
    space_after=8)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
r = p.add_run("Feinziele - Die SuS koennen ...")
r.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = ACCENT

feinziele = [
    "den Datei-Explorer oeffnen und seine Bereiche (Navigationsbereich, Inhaltsbereich, Adressleiste) benennen.  [MB 1]",
    "den Unterschied zwischen Datei und Ordner erlaeutern und je zwei Beispiele nennen.  [MB 1]",
    "Ordner erstellen, umbenennen und loeschen.  [MB 1]",
    "Dateien regelkonform benennen (aussagekraeftiger Name, passende Endung, keine Sonderzeichen).  [MB 1, MB 6]",
    "eine eigene Ordnerstruktur fuer ihr Schulfach anlegen und Dateien darin speichern.  [MB 1, MB 5]",
    "gespeicherte Dateien ueber den Datei-Explorer wieder oeffnen.  [MB 1]",
    "(Erweiterung) Dateierweiterungen (.docx, .pptx, .jpg) dem richtigen Programm zuordnen.  [MB 1]",
]
for f in feinziele:
    add_bullet(doc, f)

spacer(doc, 6)

# ==============================================================
# 3. DIDAKTISCHE UEBERLEGUNGEN
# ==============================================================
add_heading(doc, "3  Didaktische Ueberlegungen")

add_normal(doc,
    "Der Basiskurs Medienbildung BW verlangt, dass alle SuS der Klassen 5/6 grundlegende "
    "Kompetenzen im Umgang mit digitalen Werkzeugen und Dateisystemen erwerben. "
    "Die vorliegende Stunde greift auf das bereits erworbene Vorwissen (Anmelden, Word, PowerPoint) zurueck "
    "und schliesst die zentrale Luecke: die Schuelerinnen und Schueler wissen zwar, wie man Dateien erstellt, "
    "aber (noch) nicht, wo und wie sie strukturiert gespeichert werden.",
    space_after=4)

add_normal(doc,
    "Methodisch folgt die Stunde dem Prinzip des direkten Instruierens mit unmittelbarer "
    "Uebertragung in die eigene Praxis (Learning by Doing), was der handlungsorientierten "
    "Ausrichtung des Bildungsplans 2016 BW entspricht. "
    "Der Einstieg ueber eine bewusst chaotische Ordnerablage erzeugt einen kognitiven Konflikt "
    "und aktiviert die Eigenmotivation der SuS. "
    "Die abschliessende Peer-Praesentation foerdert Reflexionskompetenz (MB 6).",
    space_after=4)

spacer(doc, 6)

# ==============================================================
# 4. STUNDENVERLAUFSPLAN
# ==============================================================
add_heading(doc, "4  Stundenverlaufsplan")

headers = ["Phase / Zeit", "Inhalt / Lehreraktivitaet", "SuS-Aktivitaet",
           "Methode / Sozialform", "Medien"]

verlauf = [
    ("Einstieg\n5 Min.",
     "L zeigt per Beamer einen Desktop mit voellig unsortierten Dateien.\n"
     "Impuls: Koennt ihr die Hausaufgaben-Datei von letzter Woche finden?\n"
     "Frage: Was ist hier das Problem? Was wuerdest du aendern?\n"
     "[Bezug: MB 6 - Analysieren und Reflektieren]",
     "SuS beobachten und benennen das Problem (Unordnung, fehlende Struktur).\n"
     "SuS aeussern erste Loesungsideen.",
     "Lehrervortrag\nUnterrichtsgespraech\nPlenum",
     "Beamer\nLehrerlaptop"),
    ("Erarbeitung I\nTheorie\n10 Min.",
     "L erklaert:\n"
     "- Datei vs. Ordner (Analogie: Hefter und Blaetter)\n"
     "- Datei-Explorer oeffnen: Windows-Taste + E\n"
     "- Aufbau: Navigationsbereich, Inhaltsbereich, Adressleiste\n"
     "- Dateierweiterungen: .docx | .pptx | .jpg\n"
     "- Benennungsregeln: kein Sonderzeichen, kein Leerzeichen am Anfang\n"
     "L demonstriert live: Ordner anlegen und benennen.\n"
     "[Bezug: MB 1 - Information und Wissen]",
     "SuS hoeren zu, notieren Fachbegriffe (AB 1, Aufgabe 1).\n"
     "SuS machen Schritt fuer Schritt mit.",
     "Direktes Instruieren\nLive-Demo\nPlenum / Einzelarbeit",
     "Beamer\nAB 1"),
    ("Erarbeitung II\nStruktur erstellen\n15 Min.",
     "L gibt Aufgabe (AB 1, Aufgabe 2):\n"
     "Legt im Schueleraccount den Ordner 'Informatik_5' an.\n"
     "Erstellt darin: Aufgaben | Praesentationen | Bilder.\n"
     "L geht herum, gibt Hilfestellung.\n"
     "Schnelle SuS: Erweiterungsaufgabe (AB 1, Aufgabe 3).\n"
     "[Bezug: MB 1, MB 5 - Problemloesen und Modellieren]",
     "SuS legen die Ordnerstruktur selbststaendig an.\n"
     "SuS benennen eine Datei korrekt um und verschieben sie.\n"
     "Schnelle SuS: Dateierweiterungen-Quiz.",
     "Einzelarbeit\n(bei Bedarf Partnerarbeit)",
     "PC\nDatei-Explorer\nAB 1"),
    ("Sicherung\n10 Min.",
     "L bittet zwei SuS, ihre Struktur per Beamer zu zeigen.\n"
     "Klasse gibt Feedback nach Leitfragen:\n"
     "Was ist gut? Was wuerdest du anders benennen?\n"
     "L bespricht haeufige Fehler (Sonderzeichen, unklare Namen).\n"
     "[Bezug: MB 6 - Analysieren und Reflektieren]",
     "SuS praesentieren ihre Ordnerstruktur.\n"
     "SuS geben Peer-Feedback.\n"
     "SuS korrigieren ggf. ihre eigene Struktur.",
     "Schuelerpraesentation\nUnterrichtsgespraech\nPlenum",
     "Beamer\nPC"),
    ("Abschluss /\nReflexion\n5 Min.",
     "Reflexionsfragen:\n"
     "- Wozu brauchen wir Ordner?\n"
     "- Welche Regeln gelten beim Benennen?\n"
     "Ausblick: Ab heute speichern wir ALLE Ergebnisse strukturiert.\n"
     "[Bezug: MB 5, MB 6]",
     "SuS antworten muendlich.\n"
     "SuS tragen Merkregel auf AB 1, Aufgabe 4 ein.",
     "Unterrichtsgespraech\nPlenum",
     "AB 1"),
]

tbl2 = doc.add_table(rows=1 + len(verlauf), cols=5)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER

hrow = tbl2.rows[0]
for j, h in enumerate(headers):
    c = hrow.cells[j]
    set_cell_bg(c, "1B552E")
    p2 = c.paragraphs[0]
    p2.paragraph_format.space_before = Pt(3); p2.paragraph_format.space_after = Pt(3)
    r2 = p2.add_run(h)
    r2.bold = True; r2.font.size = Pt(9.5); r2.font.color.rgb = WHITE

for i, row_data in enumerate(verlauf):
    row = tbl2.rows[i + 1]
    fill = "EBF5ED" if i % 2 == 0 else "FFFFFF"
    for j, text in enumerate(row_data):
        c = row.cells[j]
        set_cell_bg(c, fill)
        cell_borders(c, "BBBBBB", 3)
        p2 = c.paragraphs[0]
        p2.paragraph_format.space_before = Pt(2); p2.paragraph_format.space_after = Pt(2)
        r2 = p2.add_run(text)
        r2.font.size = Pt(9)
        if j == 0:
            r2.bold = True; r2.font.color.rgb = DARK_GREEN

spacer(doc, 6)

# ==============================================================
# 5. DIFFERENZIERUNG
# ==============================================================
add_heading(doc, "5  Differenzierung")

add_normal(doc,
    "Der Bildungsplan 2016 BW betont die individuelle Foerderung aller SuS. "
    "Folgende Massnahmen werden eingesetzt:",
    space_after=4)

add_normal(doc, "Unterstuetzung (Grundsicherung):", bold=True, size=10.5, space_after=2)
add_bullet(doc, "Schrittweise Bildkarte / Screenshotanleitung auf Papier (Oeffnen des Datei-Explorers, Ordner anlegen).")
add_bullet(doc, "Partnerarbeit mit staerkeren SuS moeglich (Kooperatives Lernen, MB 2).")

add_normal(doc, "Erweiterung (schnelle SuS):", bold=True, size=10.5, space_after=2)
add_bullet(doc, "AB 1 Aufgabe 3: Dateierweiterungen-Quiz - Symbole den Programmen zuordnen.")
add_bullet(doc, "Bonusaufgabe: Weiteren Unterordner mit sinnvoller Unterteilung nach Schulfaechern anlegen (Modellierungskompetenz, MB 5).")

spacer(doc, 6)

# ==============================================================
# 6. AUSBLICK
# ==============================================================
add_heading(doc, "6  Hausaufgabe / Ausblick")
add_normal(doc,
    "Keine schriftliche Hausaufgabe. Naechste Stunde (Basiskurs Medienbildung): "
    "SuS speichern eine Word-Datei bewusst mit 'Speichern unter' im richtigen Unterordner "
    "und nutzen dabei den Dateipfad in der Adressleiste. "
    "Langfristiges Ziel: Alle SuS pflegen ihre Ordnerstruktur im Schueleraccount eigenverantwortlich "
    "und koennen Dateien gezielt fuer andere freigeben (MB 2 - Kommunikation und Kollaboration).",
    space_after=8)

# ==============================================================
# PAGE BREAK -> ARBEITSBLATT
# ==============================================================
doc.add_page_break()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(2)
r = p.add_run("ARBEITSBLATT 1  |  Basiskurs Medienbildung")
r.bold = True; r.font.size = Pt(9)
r.font.color.rgb = RGBColor(0x88, 0x88, 0x88); r.font.all_caps = True

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Dateien und Ordner verstehen & nutzen")
r.bold = True; r.font.size = Pt(17); r.font.color.rgb = DARK_GREEN

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Informatik und Medienbildung  |  Klasse 5")
r.font.size = Pt(11); r.italic = True; r.font.color.rgb = MID_GREEN

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(10)
r = p.add_run("Bildungsplan 2016 BW  |  Leitperspektive Medienbildung (MB 1, MB 5, MB 6)")
r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

t_head = doc.add_table(rows=1, cols=3)
t_head.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, label in enumerate(["Name: ______________________________",
                             "Klasse: ______________",
                             "Datum: ______________"]):
    p2 = t_head.rows[0].cells[j].paragraphs[0]
    p2.paragraph_format.space_after = Pt(10)
    r2 = p2.add_run(label); r2.font.size = Pt(10.5)

spacer(doc, 4)

# ── Aufgabe 1 ─────────────────────────────────────────────────
add_heading(doc, "Aufgabe 1 - Was ist was?  [MB 1]", level=2, color=ACCENT, size=12)
add_normal(doc,
    "Ordne die Begriffe (links) der richtigen Erklaerung (rechts) zu. "
    "Schreibe die passende Zahl in das Kaestchen.",
    space_after=6)

begriffe = [
    "1  Datei",
    "2  Ordner",
    "3  Dateiname",
    "4  Dateiendung / -erweiterung",
    "5  Datei-Explorer",
    "6  Schueleraccount",
]
erklaerungen = [
    "[ ]  Das Programm in Windows, mit dem ich Dateien und Ordner verwalte.",
    "[ ]  Ein Behaelter fuer andere Dateien oder Ordner - wie ein Hefter.",
    "[ ]  Mein persoenlicher Bereich auf dem Schulcomputer (z. B. H:\\ oder S:\\).",
    "[ ]  Die Abkuerzung hinter dem Punkt im Dateinamen, z. B. .docx oder .jpg.",
    "[ ]  Ein einzelnes gespeichertes Dokument, Bild oder Programm.",
    "[ ]  Der selbst gewaehlte Name einer Datei, z. B. 'Meine_Geschichte'.",
]

t5 = doc.add_table(rows=1 + len(begriffe), cols=2)
t5.alignment = WD_TABLE_ALIGNMENT.CENTER
hb = t5.rows[0].cells
set_cell_bg(hb[0], LIGHT_GREEN); set_cell_bg(hb[1], LIGHT_GREEN)
cell_borders(hb[0]); cell_borders(hb[1])
for c, txt in zip(hb, ["Begriffe", "Erklaerungen"]):
    p2 = c.paragraphs[0]; p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run(txt); r2.bold = True; r2.font.size = Pt(10)

for i in range(len(begriffe)):
    row = t5.rows[i + 1]
    fill = "F2FAF3" if i % 2 == 0 else "FFFFFF"
    set_cell_bg(row.cells[0], fill); cell_borders(row.cells[0])
    set_cell_bg(row.cells[1], fill); cell_borders(row.cells[1])
    for c, txt in zip(row.cells, [begriffe[i], erklaerungen[i]]):
        p2 = c.paragraphs[0]
        p2.paragraph_format.space_before = Pt(2); p2.paragraph_format.space_after = Pt(2)
        r2 = p2.add_run(txt); r2.font.size = Pt(10)

spacer(doc, 6)

# ── Aufgabe 2 ─────────────────────────────────────────────────
add_heading(doc, "Aufgabe 2 - Ordnerstruktur anlegen  [MB 1, MB 5]", level=2, color=ACCENT, size=12)
add_normal(doc, "Folge den Schritten und lege deine eigene Ordnerstruktur an.", space_after=4)

schritte = [
    "Oeffne den Datei-Explorer (Windows-Taste + E).",
    "Gehe in deinen Schueleraccount (z. B. Laufwerk H:\\ oder S:\\).",
    "Erstelle einen neuen Ordner mit dem Namen: Informatik_5",
    "Oeffne den neuen Ordner.",
    "Erstelle darin drei Unterordner:   Aufgaben   |   Praesentationen   |   Bilder",
    "Speichere eine vorhandene Word-Datei im Ordner 'Aufgaben' unter einem neuen, sinnvollen Namen.",
    "Zeige deiner Sitznachbarin / deinem Sitznachbarn die fertige Struktur. Kontrolliert gegenseitig!",
]
for s in schritte:
    p2 = doc.add_paragraph(style="List Number")
    p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(3)
    r2 = p2.add_run(s); r2.font.size = Pt(10.5)

spacer(doc, 6)

# ── Aufgabe 3 (Erweiterung) ───────────────────────────────────
add_heading(doc, "Aufgabe 3 - Erweiterung: Welches Programm oeffnet diese Datei?  [MB 1]",
            level=2, color=ACCENT, size=12)
add_normal(doc,
    "Verbinde die Dateiendung mit dem richtigen Programm. Zeichne eine Linie oder schreibe den Namen.",
    space_after=4)

endungen  = [".docx", ".pptx", ".jpg / .png", ".xlsx", ".mp3", ".pdf"]
programme = ["Microsoft Word", "Microsoft Excel", "Bildbetrachter (Windows Fotos)",
             "Microsoft PowerPoint", "Musikplayer", "Adobe Reader / Browser"]

t6 = doc.add_table(rows=len(endungen), cols=3)
t6.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (end, prog) in enumerate(zip(endungen, programme)):
    row = t6.rows[i]
    fill = "F2FAF3" if i % 2 == 0 else "FFFFFF"
    for c in row.cells:
        set_cell_bg(c, fill); cell_borders(c, "BBBBBB")
    p0 = row.cells[0].paragraphs[0]; p0.paragraph_format.space_after = Pt(2)
    r0 = p0.add_run(end); r0.bold = True; r0.font.size = Pt(10.5); r0.font.color.rgb = MID_GREEN
    p1 = row.cells[1].paragraphs[0]; p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run("->"); r1.font.size = Pt(12)
    p2 = row.cells[2].paragraphs[0]; p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run(prog); r2.font.size = Pt(10.5)

spacer(doc, 6)

# ── Aufgabe 4 – Merkregel ─────────────────────────────────────
add_heading(doc, "Aufgabe 4 - Meine Merkregel  [MB 6]", level=2, color=ACCENT, size=12)
add_normal(doc, "Ergaenze die Saetze:", space_after=4)

regeln = [
    "Ein Ordner ist wie ein __________________, weil er ____________________________________.",
    "Eine Datei erkenne ich an ihrer Dateiendung, z. B. .docx steht fuer ________________.",
    "Beim Benennen von Dateien achte ich auf: _____________________________________________.",
    "Den Datei-Explorer oeffne ich mit der Tastenkombination: _________ + _________.",
]
for reg in regeln:
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0); p2.paragraph_format.space_after = Pt(10)
    bottom_border_para(p2, color="CCCCCC", sz="4")
    r2 = p2.add_run(reg); r2.font.size = Pt(10.5)

spacer(doc, 6)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(2)
r = p.add_run(
    "Loesungen Aufgabe 1 (nur fuer Lehrkraft): "
    "Datei=5 | Ordner=2 | Dateiname=6 | Dateiendung=4 | Datei-Explorer=1 | Schueleraccount=3"
)
r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x88, 0x88, 0x88); r.italic = True

# ==============================================================
import os
import sys

# Change output path to be within the new directory
current_dir = os.path.dirname(os.path.abspath(__file__))
out = os.path.join(current_dir, "Unterrichtsentwurf_Dateien_Ordner_Kl5_BW.docx")
doc.save(out)
print("Saved:", out)
