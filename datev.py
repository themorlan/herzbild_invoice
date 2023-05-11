import datetime

from main import read_excel, create_context, today
from datetime import timedelta
import pandas as pd
from babel.dates import format_date

previous_month_end = today.replace(day=1) - timedelta(days=1)
previous_month_start = previous_month_end.replace(day=1)
previous_month_name = format_date(previous_month_start, "MMM_yyyy", locale="de_DE")


def main():
    data = read_excel()
    datev_list = []
    for idx, row in data.iterrows():
        _context = create_context(row)
        if previous_month_start <= row.Untersuchungsdatum.date() <= previous_month_end:
            _context["Buchungstext"] = f"{_context['Nachname']} {str(_context['RG_Nummer'])[1:]}"
            _context["Gesamtbetrag"] = _context["Gesamtbetrag"][:-2]
            datev_list.append(_context)

    datev = pd.DataFrame.from_dict(datev_list)
    datev.rename(columns={"Gesamtbetrag": "Umsatz", "RG_Nummer": "Rechnungsnummer", "RG_Datum": "Belegdatum",
                          "Untersuchungsdatum": "Leistungsdatum"},
                 inplace=True)
    datev["S/H"] = "H"
    datev["Gegenkonto"] = "1410"
    datev["Erlöskonto"] = "8000"
    datev.to_csv(f"neue_Rechnungen/HerzBild_Datev_export_{previous_month_name}.csv",
                 index=False,
                 sep=";",
                 encoding="UTF-8-sig",
                 columns=[
                     "Umsatz",
                     "Gegenkonto",
                     "Rechnungsnummer",
                     "Belegdatum",
                     "S/H",
                     "Erlöskonto",
                     "Buchungstext",
                     "Leistungsdatum"], )


if __name__ == "__main__":
    main()
