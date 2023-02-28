# TODO: Archive .xlsx file on run?
# TODO: Catch all errors in msg-box?

from docxtpl import DocxTemplate
from datetime import date, datetime, timedelta
import pandas as pd
from abrechnung import get_abrechnungsziffern, find_correct_drug_prizes, medikamente
from typing import Dict, Tuple
from babel.numbers import format_currency
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
import re
import pyautogui

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
    _abrechnung = get_abrechnungsziffern(df)
    lfd_nummer_tabelle = len(_abrechnung['tabelle']) + 1
    if df.Medikamente != "":
        for drug_name, data in medikamente.items():
            regex_search = re.search(data['regex'], df.Medikamente)
            if regex_search is None:
                continue
            drug_count = int(regex_search.group(0)[0])
            price = find_correct_drug_prizes(df.Rechnungsdatum, drug_name)
            _med_dict = {
                'pos': lfd_nummer_tabelle,
                'gbo': data['gbo'],
                'beschr': data['beschr'],
                'preis': format_currency(price, 'EUR', format='#.00', locale='de_DE', currency_digits=False),
                'faktor': 1.0,
                'anzahl': str(drug_count),
                'betrag': format_currency(drug_count * price, 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False),
                'betrag_raw': drug_count * price,
            }
            lfd_nummer_tabelle += 1
            _abrechnung['tabelle'].append(_med_dict)
            _abrechnung['gesamtsumme_raw'] += _med_dict['betrag_raw']

    # Calculate price of contrast agent
    if df.Kontrastmittel == "":
        raise KeyError(f"Keine Kontrastmittelangabe für {df.Vorname} {df.Nachname} {df.Rechnungsnummer}")
    elif df.Kontrastmittel == 0:
        _abrechnung['gesamtsumme_raw'] = 0
        _list_buffer = []
        for item in _abrechnung['tabelle']:
            if item['gbo'] not in ["DR", "346", "5376"]:
                _list_buffer.append(item)
        _abrechnung['tabelle'].clear()
        _abrechnung['tabelle'].extend(_list_buffer)
        for idx, item in enumerate(_abrechnung['tabelle'], start=1):
            item['pos'] = idx
            _abrechnung['gesamtsumme_raw'] += item['betrag_raw']
    else:
        _km_dict = {
            'pos': len(_abrechnung['tabelle']) + 1,
            'gbo': f"KMA{str(int(df.Kontrastmittel))}",
            'beschr': f"Kontrastmittel {str(int(df.Kontrastmittel))} ml",
            'preis': format_currency(0.2 * df.Kontrastmittel, 'EUR', format='#.00', locale='de_DE',
                                     currency_digits=False),
            'faktor': 1.0,
            'anzahl': '1',
            'betrag': format_currency(0.2 * df.Kontrastmittel, 'EUR', format='#.00 ¤', locale='de_DE',
                                      currency_digits=False),
            'betrag_raw': 0.2 * df.Kontrastmittel,
        }
        _abrechnung['tabelle'].append(_km_dict)
        _abrechnung['gesamtsumme_raw'] += _km_dict['betrag_raw']

    # Gesamtsumme has to be recalculated after messing with entries for contrast media and drugs
    _abrechnung['gesamtsumme'] = format_currency(_abrechnung['gesamtsumme_raw'], 'EUR', format='#.00 ¤', locale='de_DE',
                                                 currency_digits=False)
    _abrechnung['gesamtsumme_raw'] = round(_abrechnung['gesamtsumme_raw'], 2)

    _context = {'Anrede': df.Anrede,
                'Titel': df.Titel,
                'Vorname': df.Vorname,
                'Nachname': df.Nachname,
                'Strasse': df.Strasse,
                'PLZ': df.PLZ,
                'Stadt': df.Wohnort,
                "RG_Datum": df.Rechnungsdatum.strftime("%d.%m.%Y"),
                "RG_Nummer": df.Rechnungsnummer,
                "Druckdatum": df.Rechnungsdatum.strftime("%d.%m.%Y"),
                "Bezahlt": format_currency(df['Bereits Bezahlt'], 'EUR', format='#.00 ¤', locale='de_DE',
                                           currency_digits=False),
                "Restbetrag": format_currency(_abrechnung['gesamtsumme_raw'] - df['Bereits Bezahlt'], 'EUR',
                                              format='#.00 ¤', locale='de_DE', currency_digits=False),
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
              'Zahlungseingang': 'string',
              'Rechnung geschlossen': bool,
              'Rechnungsbetrag': 'float64',
              'Kontrastmittel': 'float64',
              'Kommentar': 'string',

              }
    # Dates are first read as str and then converted to datetime objects
    date_cols = ['Rechnungsdatum', 'Untersuchungsdatum', 'Geburtsdatum', 'Mahndatum']
    try:
        df = pd.read_excel(SOURCE_FILE, dtype=dtypes, header=1)
    except ValueError:
        pyautogui.alert(
            text="Fehler beim Einlesen der HerzBild-Liste. Möglicherweise fehlende Werte (z.B. Lfd. Nummer oder Rechnungsnummer).",
            title="Einlesefehler!", button="OK")
        raise ValueError
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


def initialize_excel_writer():
    book = openpyxl.load_workbook(SOURCE_FILE)
    sheet: Worksheet = book['Rechnungsliste']
    _values = list(sheet.values)
    _header_index = 0
    for index, _tuple in enumerate(_values):
        if "Rechnungsnummer" in _tuple:
            _header_index = index
            break
    header_row = _values[_header_index]
    return book, header_row


def write_excel(excel_book: Tuple[Workbook, Tuple[str]], column: str, row: int, value):
    _column_index = excel_book[1].index(column) + 1
    sheet: Worksheet = excel_book[0]['Rechnungsliste']
    # Row index needs + 3 cause header index is 0-based and idx argument of row is also 0-based
    _row_index = row + 3
    sheet.cell(row=_row_index, column=_column_index, value=value)


def main():
    data = read_excel()
    excel_book = initialize_excel_writer()
    datev_list = []
    for idx, row in data.iterrows():
        _context = create_context(row)
        if not row['Rechnung erstellt']:
            create_invoice(context=_context,
                           export_filename=f"neue_Rechnungen/Rechnung_{row.Rechnungsnummer}_{row.Geburtsdatum.strftime('%d.%m.%Y')}.docx")
            _context["kopie"] = "KOPIE"
            create_invoice(context=_context,
                           export_filename=f"neue_Rechnungen/Kopie_Rechnung_{row.Rechnungsnummer}_{row.Geburtsdatum.strftime('%d.%m.%Y')}.docx")
            if row['Rechnungsdatum'].date() == today:
                write_excel(excel_book=excel_book, column='Rechnungsdatum', row=idx, value=today_str)
            write_excel(excel_book=excel_book, column='Rechnungsbetrag', row=idx, value=_context['Gesamtbetrag_raw'])
            write_excel(excel_book=excel_book, column='Rechnung erstellt', row=idx, value=True)
            _context["Buchungstext"] = f"{_context['Nachname']} {str(_context['RG_Nummer'])[1:]}"
            _context["Gesamtbetrag"] = _context["Gesamtbetrag"][:-2]
            datev_list.append(_context)
        elif row.Mahnung and pd.isnull(row.Mahndatum):
            create_mahnung(context=_context,
                           export_filename=f"neue_Mahnungen/Mahnung_{row.Rechnungsnummer}_{row.Geburtsdatum.strftime('%d.%m.%Y')}.docx")
            write_excel(excel_book=excel_book, column='Mahnung', row=idx, value=True)
            write_excel(excel_book=excel_book, column='Mahndatum', row=idx, value=today_str)
    datev = pd.DataFrame.from_dict(datev_list)
    datev.rename(columns={"Gesamtbetrag": "Umsatz", "RG_Nummer": "Rechnungsnummer", "RG_Datum": "Belegdatum",
                          "Untersuchungsdatum": "Leistungsdatum"},
                 inplace=True)
    datev["S/H"] = "H"
    datev["Gegenkonto"] = "1410"
    datev["Erlöskonto"] = "8000"
    if not datev.empty:
        datev.to_csv(f"neue_Rechnungen/HerzBild_Datev_export_{today_str}.csv",
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
    excel_book[0].save(SOURCE_FILE)


if __name__ == "__main__":
    main()
