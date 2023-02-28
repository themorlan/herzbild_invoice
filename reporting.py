import pandas as pd
from main import read_excel
from datetime import datetime
from sys import argv
import plotly.express as px
from babel.dates import format_datetime
from babel.numbers import format_currency
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
from pathlib import Path

QUARTERS = {
    "1": ["01", "02", "03"],
    "2": ["04", "05", "06"],
    "3": ["07", "08", "09"],
    "4": ["10", "11", "12"]
    }

abr_to_month = {
    '01': 'Januar',
    '02': 'Februar',
    '03': 'März',
    '04': 'April',
    '05': 'Mai',
    '06': 'Juni',
    '07': 'Juli',
    '08': 'August',
    '09': 'September',
    '10': 'Oktober',
    '11': 'November',
    '12': 'Dezember',
}


def main():
    Path("Reporting").mkdir(exist_ok=True)
    df = read_excel()
    df.loc[df['Versicherung'] == "IGEL_alt", "Versicherung"] = "IGEL"
    df.loc[df['Rechnung bezahlt'] == True, "Rechnung bezahlt"] = "Bezahlt"
    df.loc[df['Rechnung bezahlt'] == False, "Rechnung bezahlt"] = "Nicht Bezahlt"
    df['Monat'] = df['Untersuchungsdatum'].apply(format_datetime, format="MMMM", locale="de_DE")
    exams_in_quarter = df.loc[
        (df['Untersuchungsdatum'].dt.quarter == quarter) & (df['Untersuchungsdatum'].dt.year == year)]
    exams_in_quarter_vorjahr = df.loc[(df['Untersuchungsdatum'].dt.quarter == quarter) & (df['Untersuchungsdatum'].dt.year == year - 1)]

    fig_rechn_bez = px.pie(exams_in_quarter, names='Rechnung bezahlt', title='Rechnung bezahlt')
    fig_rechn_bez.update_layout(legend=dict(
        yanchor="top",
        y=1,
        xanchor="left",
        x=0,
        font=dict(size=13)
    ))
    fig_rechn_bez.write_image("fig_rechn_bez.jpeg")
    fig_vers = px.pie(exams_in_quarter, names='Versicherung', title='Versicherungsart')
    fig_vers.update_layout(legend=dict(
        yanchor="top",
        y=1,
        xanchor="left",
        x=0,
        font=dict(size=13)
    ))
    fig_vers.write_image("fig_vers.jpeg")
    fig_umsatz = px.histogram(df, x="Monat", y="Bereits Bezahlt")
    fig_umsatz.update_layout(yaxis_ticksuffix=' €', yaxis_tickformat='.,', yaxis_title='Umsatz', xaxis_title=year)
    fig_umsatz.write_image("fig_umsatz.jpeg")

    tpl = DocxTemplate("Vorlagen/reporting_template.docx")
    context = {
        "fig_vers": InlineImage(tpl, image_descriptor='fig_vers.jpeg', width=Mm(84), height=Mm(60)),
        "fig_rechn_bez": InlineImage(tpl, image_descriptor='fig_rechn_bez.jpeg', width=Mm(84), height=Mm(60)),
        "fig_umsatz": InlineImage(tpl, image_descriptor='fig_umsatz.jpeg', width=Mm(165), height=Mm(117)),
        "quarter": quarter,
        "year": year,
        "month_1": abr_to_month[QUARTERS[str(quarter)][0]],
        "month_2": abr_to_month[QUARTERS[str(quarter)][1]],
        "month_3": abr_to_month[QUARTERS[str(quarter)][2]],
        "u_month_1": len(exams_in_quarter.loc[exams_in_quarter['Monat'] == abr_to_month[QUARTERS[str(quarter)][0]]].index),
        "u_month_2": len(exams_in_quarter.loc[exams_in_quarter['Monat'] == abr_to_month[QUARTERS[str(quarter)][1]]].index),
        "u_month_3": len(exams_in_quarter.loc[exams_in_quarter['Monat'] == abr_to_month[QUARTERS[str(quarter)][2]]].index),
        "u_ges": len(exams_in_quarter.index),
        "umsatz": format_currency(exams_in_quarter['Bereits Bezahlt'].sum(), 'EUR', format='#,##0\xa0¤', locale='de_DE', currency_digits=False),
        "umsatz_vorjahr": exams_in_quarter_vorjahr['Bereits Bezahlt'].sum() if not exams_in_quarter_vorjahr.empty else "Nicht verfügbar.",}

    tpl.render(context)
    tpl.save(f"Reporting/Reporting_Q{quarter}_{year}.docx")

    os.remove("fig_umsatz.jpeg")
    os.remove("fig_rechn_bez.jpeg")
    os.remove("fig_vers.jpeg")


if __name__ == '__main__':
    try:
        quarter = int(argv[1])
    except IndexError:
        print("Es wurde kein Quartal angegeben.")
        exit(1)

    try:
        year = int(argv[2])
    except IndexError:
        print("Es wurde kein Abrechnungsjahr angegeben.")
        exit(1)
    main()
