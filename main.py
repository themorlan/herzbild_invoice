from docxtpl import DocxTemplate
from datetime import date, datetime, timedelta
import pandas as pd
from abrechnung import get_abrechnungsziffern
from typing import Dict
from babel.numbers import format_currency
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import re
import traceback
import readchar

SOURCE_FILE = "Rechnungsliste_HerzBild.xlsx"
TEMPLATE_FILE = "Vorlagen/Musterrechnung_HerzBild_Vorlage.docx"
MAHNUNG_TEMPLATE_FILE = "Vorlagen/Mahnung_HerzBild_Vorlage.docx"

pd.set_option("display.max_columns", None)
today = date.today()
today_str = today.strftime("%d.%m.%Y")

Path("neue_Rechnungen").mkdir(exist_ok=True)
Path("neue_Mahnungen").mkdir(exist_ok=True)
Path("Vorlagen").mkdir(exist_ok=True)


def create_context(df: pd.Series) -> Dict:
    _abrechnung = get_abrechnungsziffern(df.Versicherung, df.Abrechnungsziffern)
    # Calculate price of used drugs
    regex_meto = re.search(r"\d[^\dMm]*[Mm]", df.Medikamente)
    regex_beloc = re.search(r"\d[^\dBb]*[Bb]", df.Medikamente)
    if df.Medikamente != "" and regex_meto is not None or regex_beloc is not None:
        _med_dict = {
            'pos': len(_abrechnung['tabelle']) + 1,
            'gbo': "MEDHCT",
            'beschr': "Metoprolol ",
            'preis': "",
            'faktor': 1,
            'anzahl': '1',
            'betrag': "",
            'betrag_raw': 0.0,
        }
        if regex_meto is not None and regex_beloc is not None:
            _count_meto = int(regex_meto.group(0)[0]) #0.16
            _count_beloc = int(regex_beloc.group(0)[0]) #4.30
            _med_dict['beschr'] += "oral und i.v."
            _med_dict['betrag_raw'] = _count_meto * 0.16 + _count_beloc * 4.30
        elif regex_meto is not None:
            _count_meto = int(regex_meto.group(0)[0])  # 0.16
            _med_dict['beschr'] += "oral"
            _med_dict['betrag_raw'] = _count_meto * 0.16
        elif regex_beloc is not None:
            _count_beloc = int(regex_beloc.group(0)[0])  # 4.30
            _med_dict['beschr'] += "i.v."
            _med_dict['betrag_raw'] = _count_beloc * 4.30
        _med_dict['preis'] = format_currency(_med_dict['betrag_raw'], 'EUR', format='#.00', locale='de_DE', currency_digits=False)
        _med_dict['betrag'] = format_currency(_med_dict['betrag_raw'], 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False)
        _abrechnung['tabelle'].append(_med_dict)
        _abrechnung['gesamtsumme_raw'] += _med_dict['betrag_raw']

    # Calculate price of contrast agent
    if df.Kontrastmittel == "":
        raise KeyError(f"Keine Kontrastmittelangabe für {df.Vorname} {df.Nachname} {df.Rechnungsnummer}")
    else:
        _km_dict = {
            'pos': len(_abrechnung['tabelle']) + 1,
            'gbo': f"KMA{str(int(df.Kontrastmittel))}",
            'beschr': f"Kontrastmittel {str(int(df.Kontrastmittel))} ml",
            'preis': format_currency(0.2 * df.Kontrastmittel, 'EUR', format='#.00', locale='de_DE',
                                     currency_digits=False),
            'faktor': 1,
            'anzahl': '1',
            'betrag': format_currency(0.2 * df.Kontrastmittel, 'EUR', format='#.00 ¤', locale='de_DE',
                                      currency_digits=False),
            'betrag_raw': 0.2 * df.Kontrastmittel,
        }
        _abrechnung['tabelle'].append(_km_dict)
        _abrechnung['gesamtsumme_raw'] += _km_dict['betrag_raw']

    # Gesamtsumme has to be recalculated after messing with entries for contrast media and drugs
    _abrechnung['gesamtsumme'] = format_currency(_abrechnung['gesamtsumme_raw'], 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False)

    _context = {'Anrede': df.Anrede,
                'Titel': df.Titel,
                'Vorname': df.Vorname,
                'Nachname': df.Nachname,
                'Strasse': df.Strasse,
                'PLZ': df.PLZ,
                'Stadt': df.Wohnort,
                "RG_Datum": today_str,
                "RG_Nummer": df.Rechnungsnummer,
                "Druckdatum": df.Rechnungsdatum.strftime("%d.%m.%Y"),
                "Bezahlt": format_currency(df['Bereits Bezahlt'], 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False),
                "Restbetrag": format_currency(_abrechnung['gesamtsumme_raw'] - df['Bereits Bezahlt'], 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False),
                "heute": today_str,
                "Untersuchungsdatum": df.Untersuchungsdatum.strftime("%d.%m.%Y"),
                "Geburtsdatum": df.Geburtsdatum.strftime("%d.%m.%Y"),
                "Gesamtbetrag": _abrechnung['gesamtsumme'],
                "Gesamtbetrag_raw": _abrechnung['gesamtsumme_raw'],
                "tabelle": _abrechnung['tabelle']
                }
    return _context


def create_invoice(context: Dict, export_filename: str) -> None:
    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(context)
    doc.save(export_filename)


def create_mahnung(context: Dict, export_filename: str) -> None:
    doc = DocxTemplate(MAHNUNG_TEMPLATE_FILE)
    doc.render(context)
    doc.save(export_filename)


def read_excel() -> pd.DataFrame:
    dtypes = {'Lfd. Nummer': 'int64',
              'Rechnungsnummer': 'int64',
              'Rechnungsdatum': 'string',
              'Abrechnungsziffern': 'string',
              'Versicherung': 'string',
              'Untersuchungsdatum': 'string',
              'Medikamente': 'string',
              'Fallnummer': 'string',
              'PatientenID': 'string',
              'Anrede': 'string',
              'Titel': 'string',
              'Vorname': 'string',
              'Nachname': 'string',
              'Geburtsdatum': 'string',
              'Strasse': 'string',
              'PLZ': 'string',
              'Wohnort': 'string',
              'Storno': bool,
              'Rechnung erstellt': bool,
              'Rechnung bezahlt': bool,
              'Mahnung': bool,
              'Mahndatum': 'string',
              'Bereits Bezahlt': 'float64',
              'Rechnung geschlossen': bool,
              'Rechnungsbetrag': 'float64',
              'Kontrastmittel': 'float64',
              'Kommentar': 'string',

              }
    # Dates are first read as str and then converted to datetime objects
    date_cols = ['Rechnungsdatum', 'Untersuchungsdatum', 'Geburtsdatum', 'Mahndatum']
    df = pd.read_excel(SOURCE_FILE, dtype=dtypes, header=1)
    df[date_cols] = df[date_cols].apply(pd.to_datetime, errors='coerce', dayfirst=True)
    # Templates can't deal with NaN - float values. Convert them to empty strings
    df = df.fillna("")
    # If no date for invoice was given set it to today
    df.loc[df['Rechnungsdatum'].isna(), 'Rechnungsdatum'] = pd.Timestamp(today)
    # Add whitespace after title, if present
    df.loc[df['Titel'].str.len() > 0, 'Titel'] = (df['Titel'] + " ")
    # Cell needs to be float to make calculations possible
    df.loc[df['Bereits Bezahlt'] == "", 'Bereits Bezahlt'] = 0.0
    return df


def write_excel(column: str, row, value):
    book = openpyxl.load_workbook(SOURCE_FILE)
    sheet: Worksheet = book['Rechnungsliste']
    _values = list(sheet.values)
    _header_index = 0
    for index, _tuple in enumerate(_values):
        if "Rechnungsnummer" in _tuple:
            _header_index = index
            break
    _column_index = _values[_header_index].index(column) + 1
    # Row index needs + 3 cause header index is 0-based and idx argument of row is also 0-based
    _row_index = row + 3
    sheet.cell(row=_row_index, column=_column_index, value=value)
    book.save(SOURCE_FILE)


def main():
    data = read_excel()
    for idx, row in data.iterrows():
        _context = create_context(row)
        if not row['Rechnung erstellt']:
            create_invoice(context=_context, export_filename=f"neue_Rechnungen/Rechnung_{row.Rechnungsnummer}_{row.Geburtsdatum.strftime('%d.%m.%Y')}.docx")
            if row['Rechnungsdatum'].date() == today:
                write_excel(column='Rechnungsdatum', row=idx, value=today_str)
                write_excel(column='Rechnung erstellt', row=idx, value=True)
                write_excel(column='Rechnungsbetrag', row=idx, value=_context['Gesamtbetrag_raw'])
        elif row['Rechnungsdatum'].date() < today - timedelta(days=30) and not row['Rechnung bezahlt']:
            create_mahnung(context=_context, export_filename=f"neue_Mahnungen/Mahnung_{row.Rechnungsnummer}_{row.Geburtsdatum.strftime('%d.%m.%Y')}.docx")
            write_excel(column='Mahnung', row=idx, value=True)
            write_excel(column='Mahndatum', row=idx, value=today_str)


if __name__ == "__main__":
    try:
        main()
    except:
        print(traceback.format_exc())
    finally:
        print("Drücke eine beliebige Taste um das Programm zu beenden...")
        k = readchar.readchar()