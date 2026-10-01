# leadsheet.py: builds the styled daily SEE:ME lead sheet (xlsx, converted to a Google Sheet on upload)
# usage: python3 leadsheet.py leads.json out.xlsx DD.MM.YYYY
#        python3 leadsheet.py --csv2json export.csv leads.json   (reads a sheet exported from Drive as text/csv)
# "Gefunden am" is kept if the lead already has one (weekly check), "Geprüft am" is always set to DD.MM.YYYY
import sys, json, re, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

COLS = [  # (header, group, width)
 ("Priorität","LEAD",13),("Firmenname","LEAD",32),("Kategorie","LEAD",22),("Signal-Typ","LEAD",17),("Lead-Signal","LEAD",46),
 ("Outreach-Aufhänger","LEAD",44),("Gesuchte Jobtypen","LEAD",32),("Offene Stellen (ca.)","LEAD",14),
 ("Adresse","STANDORT",32),("Ort / Stadtteil","STANDORT",20),
 ("Ansprechpartner","KONTAKT",24),("E-Mail","KONTAKT",32),("Telefon","KONTAKT",20),("WhatsApp","KONTAKT",20),
 ("Kontaktformular","KONTAKT",36),("Website","KONTAKT",26),("Weitere Kontaktwege","KONTAKT",32),
 ("Quelle","HERKUNFT",42),("Gefunden am","HERKUNFT",15),("Geprüft am","HERKUNFT",15),
 ("Status","TEAM",18),("Notizen","TEAM",40),
]
H = [c[0] for c in COLS]
CAT = {"Promotionagentur":"E3EDFF","Gastronomie":"FDEBDD","Hotellerie":"F3E8FD","Einzelhandel":"E4F5EA","Logistik/Lieferdienst":"FFF5D6",
       "Events/Freizeit":"FCE4EC","Fitness":"E0F4F4","Reinigung/Sicherheit":"ECEFF3","Sonstiges":"F4F4F6"}
CATS = list(CAT)
SIGNALS = ["Stellenanzeige","Expansion","Branchenfit"]
STATUS = {"Neu":"EAF2FF","Kontaktiert":"FFF8E1","Im Gespräch":"FFF0E0","Pilot":"E8F7EC","Kein Interesse":"F4F4F6","Später":"F3EEFF"}
PRIO = {"A":("E6F0FF","0165FF"),"B":("F2F2F8","000000"),"C":("FFFFFF","8A8A8A")}  # (fill, text)
BLUE, INK, MUTED, ROW2, LINE = "0165FF","1A1A1A","8A8A8A","FAFAFC","E6E6EE"
FONT, SIZE, NF = "Inter", 11, "nicht gefunden"
thin = Side(style="thin", color=LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
DATES = ("Gefunden am","Geprüft am")
CENTER = ("Priorität","Offene Stellen (ca.)","Gefunden am","Geprüft am")

def fill(c): return PatternFill("solid", start_color=c, end_color=c)
def f(**k): return Font(name=FONT, size=k.pop("size", SIZE), **k)
def domain(u):
    u = re.sub(r"^https?://", "", (u or "").strip().lower()); u = re.sub(r"^www\.", "", u)
    return u.split("/")[0]

def clean(r, date):
    d = {h: str(r.get(h, "") or "").strip() for h in H}
    if d["Website"] and d["Website"] != NF: d["Website"] = domain(d["Website"])
    d["Priorität"] = d["Priorität"].upper()[:1]
    if d["Kategorie"] not in CAT: d["Kategorie"] = "Sonstiges"
    if d["Signal-Typ"] not in SIGNALS: d["Signal-Typ"] = "Stellenanzeige" if d["Offene Stellen (ca.)"].isdigit() else "Branchenfit"
    if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", d["Gefunden am"]): d["Gefunden am"] = date
    d["Geprüft am"] = date
    d["Status"] = d["Status"] or "Neu"
    for h in H:
        if not d[h] and h != "Notizen": d[h] = NF
    return d

def height(r):
    lines = 1
    for h, _, w in COLS:
        v = r.get(h, "")
        lines = max(lines, sum(max(1, -(-len(p) // max(1, int(w * 0.95)))) for p in v.split("\n")))
    return min(18 * lines + 16, 300)

def build(leads, path, date):
    rank = {"A":0,"B":1,"C":2}
    rows = sorted((clean(r, date) for r in leads), key=lambda r: (CATS.index(r["Kategorie"]) if r["Kategorie"] in CATS else len(CATS), rank.get(r["Priorität"], 3), r["Firmenname"].lower()))
    wb = Workbook(); ws = wb.active; ws.title = "Leads"; ws.sheet_properties.tabColor = BLUE
    ws.sheet_view.showGridLines = False
    last = max(3, len(rows) + 2); dv_end = last + 200

    # row 1: quiet group band
    groups = []
    for i, (_, g, _) in enumerate(COLS, 1):
        if groups and groups[-1][0] == g: groups[-1][2] = i
        else: groups.append([g, i, i])
    for g, a, b in groups:
        for col in range(a, b + 1):
            c = ws.cell(1, col); c.fill = fill("F2F2F8")
            c.border = Border(left=Side(style="thin", color="FFFFFF") if col == a else None, bottom=Side(style="thin", color="FFFFFF"))
        ws.merge_cells(start_row=1, start_column=a, end_row=1, end_column=b)
        c = ws.cell(1, a, g); c.font = f(size=9, bold=True, color=MUTED)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    # row 2: white header with blue underline
    for i, (h, _, w) in enumerate(COLS, 1):
        c = ws.cell(2, i, h); c.font = f(bold=True, color=INK)
        c.fill = fill("FFFFFF")
        c.border = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color=BLUE))
        c.alignment = Alignment(horizontal="center" if h in CENTER else "left", vertical="center", wrap_text=True, indent=0 if h in CENTER else 1)
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[2].height = 40

    for ri, r in enumerate(rows, 3):
        bg = "FFFFFF" if ri % 2 else ROW2
        for ci, h in enumerate(H, 1):
            v = r[h]
            if h == "Offene Stellen (ca.)" and v.isdigit(): v = int(v)
            if h in DATES and re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", v): v = dt.datetime.strptime(v, "%d.%m.%Y")
            c = ws.cell(ri, ci, v)
            if h in DATES: c.number_format = "DD.MM.YYYY"
            c.fill = fill(bg); c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center" if h in CENTER else "left", indent=0 if h in CENTER else 1)
            if v == NF: c.font = f(italic=True, color="ADADAD")
            elif h == "Firmenname": c.font = f(bold=True, color=INK)
            elif h in ("Website","E-Mail","Kontaktformular","Quelle") and "\n" not in str(v) and ("." in str(v)):
                c.hyperlink = ("mailto:" + v) if h == "E-Mail" else (v if v.startswith("http") else "https://" + v)
                c.font = f(color=BLUE)
            else: c.font = f(color=INK)
            if h == "Priorität" and v in PRIO:
                c.fill = fill(PRIO[v][0]); c.font = f(size=12, bold=True, color=PRIO[v][1])
                c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[ri].height = height(r)

    sc, pc, kc, gc = (L(H.index(x) + 1) for x in ("Status","Priorität","Kategorie","Signal-Typ"))
    for s, col in STATUS.items():
        ws.conditional_formatting.add(f"{sc}3:{sc}{dv_end}", FormulaRule(formula=[f'${sc}3="{s}"'], fill=fill(col), font=Font(name=FONT, bold=True, color=INK)))
    for k, col in CAT.items():
        ws.conditional_formatting.add(f"{kc}3:{kc}{dv_end}", FormulaRule(formula=[f'${kc}3="{k}"'], fill=fill(col), font=Font(name=FONT, bold=True, color=INK)))
    for colL, opts in ((sc, list(STATUS)), (pc, list(PRIO)), (kc, CATS), (gc, SIGNALS)):
        dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
        dv.add(f"{colL}3:{colL}{dv_end}"); ws.add_data_validation(dv)
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{L(len(H))}{last}"

    # Nach Firma: live view of the Leads tab, sorted A-Z by company (follows edits in Leads automatically)
    nf = wb.create_sheet("Nach Firma"); nf.sheet_properties.tabColor = BLUE; nf.sheet_view.showGridLines = False
    nf.cell(1, 1, "Alle Leads dieser Liste nach Firmenname sortiert (A bis Z). Nur Ansicht: Status und Notizen bitte im Tab \"Leads\" ändern, hier wird alles automatisch übernommen.")
    nf.cell(1, 1).font = f(size=10, italic=True, color=MUTED); nf.cell(1, 1).alignment = Alignment(vertical="center", indent=1)
    nf.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8); nf.row_dimensions[1].height = 24
    order = ["Firmenname"] + [h for h in H if h != "Firmenname"]
    for i, h in enumerate(order, 1):
        c = nf.cell(2, i, h); c.font = f(bold=True, color=INK); c.fill = fill("FFFFFF")
        c.border = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color=BLUE))
        c.alignment = Alignment(horizontal="center" if h in CENTER else "left", vertical="center", wrap_text=True, indent=0 if h in CENTER else 1)
        nf.column_dimensions[L(i)].width = dict((x[0], x[2]) for x in COLS)[h]
    nf.row_dimensions[2].height = 40
    nf_end = last + 3  # the view only needs as many rows as this list has leads
    for ri in range(3, nf_end + 1):
        for ci, h in enumerate(order, 1):
            c = nf.cell(ri, ci); c.border = BORDER; c.font = f(bold=(h == "Firmenname"), color=INK)
            c.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center" if h in CENTER else "left", indent=0 if h in CENTER else 1)
            if h in DATES: c.number_format = "DD.MM.YYYY"
    cols = ",".join(f"Leads!{L(H.index(h) + 1)}3:{L(H.index(h) + 1)}{dv_end}" for h in order)
    nf.cell(3, 1, f'=IFERROR(SORT(FILTER({{{cols}}},Leads!B3:B{dv_end}<>""),1,TRUE),"")')
    last_col = L(len(order))
    rng_all = f"A3:{last_col}{nf_end}"
    nf.conditional_formatting.add(rng_all, FormulaRule(formula=['AND($A3<>"",ISEVEN(ROW()))'], fill=fill(ROW2)))
    nf.conditional_formatting.add(rng_all, FormulaRule(formula=[f'$A3=""'], font=Font(color="FFFFFF"), border=Border()))
    for h, rules in (("Priorität", {k: (v[0], v[1]) for k, v in PRIO.items()}), ("Kategorie", {k: (v, INK) for k, v in CAT.items()}), ("Status", {k: (v, INK) for k, v in STATUS.items()})):
        cl = L(order.index(h) + 1)
        for k, (bgc, fg) in rules.items():
            nf.conditional_formatting.add(f"{cl}3:{cl}{nf_end}", FormulaRule(formula=[f'${cl}3="{k}"'], fill=fill(bgc), font=Font(name=FONT, bold=True, color=fg)))
    for h in order:
        if h in ("Firmenname", "Kategorie", "Priorität", "Status"): continue
        cl = L(order.index(h) + 1)
        nf.conditional_formatting.add(f"{cl}3:{cl}{nf_end}", FormulaRule(formula=[f'${cl}3="{NF}"'], font=Font(name=FONT, italic=True, color="ADADAD")))
    nf.freeze_panes = "B3"

    # Übersicht
    o = wb.create_sheet("Übersicht"); o.sheet_properties.tabColor = "ADADAD"; o.sheet_view.showGridLines = False
    o.column_dimensions["A"].width = 4; o.column_dimensions["B"].width = 34; o.column_dimensions["C"].width = 14
    t = o.cell(2, 2, "SEE:ME Leads Hamburg"); t.font = f(size=18, bold=True, color=INK)
    s = o.cell(3, 2, f"Neue Leads vom {date}. Jede Nacht kommt eine neue Liste dazu."); s.font = f(size=10, color=MUTED)
    o.row_dimensions[2].height = 30
    r = 5
    def line(label, val, head=False):
        a, b = o.cell(r, 2, label), o.cell(r, 3, val)
        for c in (a, b):
            c.font = f(bold=head, color=INK)
            c.border = Border(bottom=Side(style="medium", color=BLUE) if head else thin)
            c.alignment = Alignment(vertical="center", indent=1)
        b.alignment = Alignment(horizontal="center", vertical="center")
        o.row_dimensions[r].height = 26
    line("Leads in dieser Liste", f"=COUNTA(Leads!B3:B{dv_end})", True); r += 2
    for title, colL, opts in (("Nach Priorität", pc, list(PRIO)), ("Nach Signal-Typ", gc, SIGNALS), ("Nach Status", sc, list(STATUS)), ("Nach Kategorie", kc, CATS)):
        line(title, "Anzahl", True); r += 1
        for opt in opts:
            line(opt, f'=COUNTIF(Leads!{colL}3:{colL}{dv_end},"{opt}")')
            if opt in PRIO: o.cell(r, 2).fill = fill(PRIO[opt][0]); o.cell(r, 2).font = f(bold=True, color=PRIO[opt][1])
            if opt in STATUS and colL == sc: o.cell(r, 2).fill = fill(STATUS[opt])
            if opt in CAT: o.cell(r, 2).fill = fill(CAT[opt])
            r += 1
        r += 1

    # Legende
    g = wb.create_sheet("Legende"); g.sheet_properties.tabColor = "ADADAD"; g.sheet_view.showGridLines = False
    g.column_dimensions["A"].width = 4; g.column_dimensions["B"].width = 26; g.column_dimensions["C"].width = 96
    items = [
     ("So funktioniert das System", None),
     ("Tageslisten", "Jede Nacht sucht Claude neue Leads in Hamburg und Umgebung und legt im Ordner \"SEE:ME Leads\" eine neue Liste mit dem Datum im Namen an. Neueste Liste = neueste Leads. Jeder Wochentag hat einen eigenen Schwerpunkt (Branche oder Gebiet)."),
     ("Keine Dubletten", "Ein Betrieb steht nur einmal im ganzen Ordner. Vor jedem neuen Eintrag prüft Claude alle bisherigen Listen."),
     ("Bearbeiten", "Status und Notizen pflegt ihr direkt im Tab \"Leads\" der jeweiligen Tagesliste. Claude übernimmt eure Einträge bei jeder Kontrolle."),
     ("Sortierung", "Tab \"Leads\": nach Kategorie gruppiert (Promotionagenturen zuerst), innerhalb der Kategorie Priorität A zuerst, danach alphabetisch. Tab \"Nach Firma\": dieselben Leads von A bis Z nach Firmenname, aktualisiert sich automatisch. Über die Filterpfeile im Tabellenkopf könnt ihr zusätzlich jederzeit anders sortieren oder filtern."),
     ("Wöchentliche Kontrolle", "Jeden Montag früh prüft Claude alle Tageslisten. Leads mit Status \"Neu\" und Signal-Typ \"Stellenanzeige\", deren Anzeigen nicht mehr online sind, werden entfernt. Leads mit Signal-Typ \"Expansion\" oder \"Branchenfit\" bleiben stehen. Bearbeitete Leads (anderer Status oder Notizen) bleiben immer stehen und bekommen einen Hinweis im Lead-Signal. Die vorherige Version jeder geänderten Liste liegt im Unterordner \"Archiv\"."),
     ("Geprüft am", "Datum, an dem der Lead zuletzt aktualisiert wurde (beim Anlegen oder bei der wöchentlichen Kontrolle). Listen ohne Änderungen werden geprüft, aber nicht neu angelegt."),
     ("Outreach-Aufhänger", "Ein konkreter Anknüpfungspunkt für die erste Nachricht, zum Beispiel eine Neueröffnung oder eine aktuelle Suche nach Aushilfen. Stammt nur aus öffentlichen Quellen."),
     ("nicht gefunden", "Diese Angabe war öffentlich nicht auffindbar. Kontaktdaten stammen nur aus öffentlichen Angaben des Betriebs (Website, Impressum, Stellenanzeige, Verzeichnis eines Einkaufszentrums)."),
     ("Priorität", None),
     ("A", "Promotionagentur, 5 oder mehr offene Stellen oder klare Expansion. Zuerst ansprechen."),
     ("B", "2 bis 4 offene Stellen in einer Zielbranche oder Betrieb mit mehreren Standorten in Hamburg."),
     ("C", "Passende Branche ohne aktuelle Anzeige oder schwächeres Signal. Gut für breiten Outreach."),
     ("Signal-Typ", None),
     ("Stellenanzeige", "Es gibt aktuell veröffentlichte Stellen."),
     ("Expansion", "Neueröffnung, neue Filiale, Umzug oder Finanzierung."),
     ("Branchenfit", "Keine aktuelle Anzeige gefunden, aber typischer Betrieb mit viel jungem Personal und hoher Fluktuation."),
     ("Status", None),
     ("Neu", "Von Claude gefunden, noch nicht kontaktiert."),
     ("Kontaktiert", "Erste Nachricht ist raus."),
     ("Im Gespräch", "Es gibt eine Rückmeldung oder einen Termin."),
     ("Pilot", "Der Betrieb testet SEE:ME."),
     ("Kein Interesse", "Abgesagt."),
     ("Später", "Später erneut ansprechen, Grund in den Notizen."),
     ("Kategorie", None),
     *[(k, v) for k, v in {"Promotionagentur":"Promotion-, Event- und Messepersonal-Agenturen. Wichtigste Zielgruppe.","Gastronomie":"Restaurants, Cafés, Bars, Systemgastronomie.","Hotellerie":"Hotels und Hostels.","Einzelhandel":"Läden, Filialketten, Supermärkte.","Logistik/Lieferdienst":"Lager, Kurier- und Lieferdienste.","Events/Freizeit":"Eventlocations, Kinos, Freizeitparks, Clubs.","Fitness":"Fitnessstudios und Sportanlagen.","Reinigung/Sicherheit":"Reinigungs- und Sicherheitsfirmen.","Sonstiges":"Alles andere mit Personalbedarf."}.items()],
     ("Spaltengruppen", None),
     ("LEAD", "Wer ist der Betrieb, warum ist er interessant und womit steigt ihr ins Gespräch ein."),
     ("STANDORT", "Adresse und Ort bzw. Stadtteil."),
     ("KONTAKT", "Alle öffentlichen Kontaktwege. Mehrere Einträge stehen untereinander."),
     ("HERKUNFT", "Wo und wann der Lead gefunden und zuletzt geprüft wurde."),
     ("TEAM", "Eure Spalten für Status und Notizen."),
    ]
    r = 2
    for a, b in items:
        if b is None:
            if r > 2: r += 1
            c = g.cell(r, 2, a); c.font = f(size=13, bold=True, color=INK)
            for col in (2, 3): g.cell(r, col).border = Border(bottom=Side(style="medium", color=BLUE))
            g.row_dimensions[r].height = 30
        else:
            ca, cb = g.cell(r, 2, a), g.cell(r, 3, b)
            ca.font = f(bold=True, color=INK); cb.font = f(color=INK)
            for c in (ca, cb):
                c.border = Border(bottom=thin); c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
            if a in PRIO: ca.fill = fill(PRIO[a][0]); ca.font = f(bold=True, color=PRIO[a][1])
            if a in STATUS: ca.fill = fill(STATUS[a])
            if a in CAT: ca.fill = fill(CAT[a])
            g.row_dimensions[r].height = 18 * max(1, -(-len(b) // 90)) + 16
        r += 1
    wb.save(path)
    return len(rows)

def csv2json(src, dst):
    import csv
    rows = list(csv.reader(open(src, encoding="utf-8-sig", newline="")))
    hi = next(i for i, row in enumerate(rows) if "Firmenname" in row)
    idx = {name: j for j, name in enumerate(rows[hi]) if name}
    out = []
    for row in rows[hi + 1:]:
        d = {h: (row[idx[h]].strip() if h in idx and idx[h] < len(row) else "") for h in H}
        if not d["Firmenname"]: continue
        for h in H:
            if d[h] == NF: d[h] = ""
        out.append(d)
    json.dump(out, open(dst, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"ok: {len(out)} leads read -> {dst}")

if __name__ == "__main__":
    if sys.argv[1] == "--csv2json":
        csv2json(sys.argv[2], sys.argv[3]); sys.exit(0)
    leads = json.load(open(sys.argv[1], encoding="utf-8"))
    n = build(leads, sys.argv[2], sys.argv[3])
    print(f"ok: {n} leads -> {sys.argv[2]}")
