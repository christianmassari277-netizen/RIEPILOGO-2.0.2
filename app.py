import sys
import os
import re
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

ROW_RE = re.compile(r"^\s*(\d{7})\s+(\d+)\s+(\d+)\s+(-?\d+)\b", re.MULTILINE)
HEADER = ["NUMERO GARANZIA", "SUFFISSO", "JOB", "TOTALE JOB"]
DISCLAIMER = (
    "Disclaimer: I totali riportati nel presente documento sono stati calcolati automaticamente.\n"
    "A causa di possibili arrotondamenti e differenze di calcolo, potrebbero verificarsi scostamenti minimi di qualche euro rispetto ai valori ufficiali di fatturazione."
)

def parse_txt(path: str) -> pd.DataFrame:
    with open(path, "r", encoding="latin-1", errors="ignore") as f:
        text = f.read()
    matches = ROW_RE.findall(text)
    if not matches:
        raise ValueError("Nessuna riga valida trovata nel TXT")
    df = pd.DataFrame(matches, columns=HEADER)
    df["JOB"] = pd.to_numeric(df["JOB"], errors="coerce").fillna(0).astype(int)
    df["TOTALE JOB"] = pd.to_numeric(df["TOTALE JOB"], errors="coerce").fillna(0).astype(int)
    df = df.drop_duplicates(subset=["NUMERO GARANZIA", "SUFFISSO", "JOB"])
    df = df.sort_values(by=["NUMERO GARANZIA", "SUFFISSO", "JOB"]).reset_index(drop=True)
    return df

def eur_fmt(val: Decimal) -> str:
    q = val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{q:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"

def build_pdf(df, out_pdf: str):
    totale = Decimal(int(df["TOTALE JOB"].sum()))
    iva = (totale * Decimal("0.22")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    ivato = totale + iva

    doc = SimpleDocTemplate(out_pdf, pagesize=A4)
    styles = getSampleStyleSheet()
    title = Paragraph("<b>Riepilogo Garanzie</b>", styles["Heading3"])

    data = [HEADER] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("ALIGN", (2,1), (2,-1), "CENTER"),
        ("ALIGN", (3,1), (3,-1), "RIGHT"),
    ]))

    totals = [
        ["Totale", f"{int(totale)}"],
        ["IVA 22%", eur_fmt(iva)],
        ["Totale IVA inclusa", eur_fmt(ivato)],
    ]
    t_table = Table(totals)
    t_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("ALIGN", (0,0), (-1,-1), "RIGHT"),
    ]))

    disclaimer = Paragraph("<font size=7><i>" + DISCLAIMER.replace("\n", "<br/>") + "</i></font>", styles["Normal"])
    story = [title, Spacer(1,10), table, Spacer(1,10), t_table, Spacer(1,10), disclaimer]
    doc.build(story)

def output_path_for(txt_path: str) -> str:
    base, _ = os.path.splitext(txt_path)
    return base + "_Riepilogo_Garanzie.pdf"

def main():
    files = [f for f in sys.argv[1:] if os.path.isfile(f)]
    if not files:
        print("Trascina un file .txt sull'exe per generare il PDF")
        return
    for f in files:
        try:
            df = parse_txt(f)
            pdf = output_path_for(f)
            build_pdf(df, pdf)
            print("Creato:", pdf)
        except Exception as e:
            print("Errore su", f, ":", e)

if __name__ == "__main__":
    main()
