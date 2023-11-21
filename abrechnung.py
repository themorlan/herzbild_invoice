from typing import Dict, List
from babel.numbers import format_currency
from datetime import datetime
import pandas as pd


def get_price_for_gbo(gbo_code: str, date: datetime) -> float:
    price_for_code = {
        "DR": {
            datetime(2022, 10, 1): 1.73,
            datetime(2023, 11, 20): 2.06,
            }
    }
    return price_for_code[gbo_code][find_associated_date(date, price_for_code[gbo_code].keys())]


gbo = {
    "1": {"beschr": "Beratung auch telefonisch", "preis": 4.66},
    "5": {"beschr": "Untersuchung, symptombezogen", "preis": 4.66},
    "75": {"beschr": "Befundbericht, ausführlich", "preis": 7.58},
    "DR": {"beschr": "Sachkosten, Hochdruckspritzenleitung", "preis": 1.73},
    "346": {"beschr": "Kontrastmittel, i.v. Hochdruck", "preis": 17.49},
    "650": {"beschr": "EKG/Notfall-EKG", "preis": 8.86},
    "5371": {"beschr": "CT, Hals-/Thoraxbereich", "preis": 134.06},
    "5377": {"beschr": "Computeranalyse/3D-Rekonstruktion zusätzlich zu 5370-5375", "preis": 46.63},
    "5376": {"beschr": "CT, ergänzende Serie zusätzlich zu 5370-5375", "preis": 29.14},
    "60": {"beschr": "Konsilium", "preis": 6.99},
}

medikamente = {"Metoprolol_oral": {"gbo": "METO",
                                   "beschr": "Metoprolol oral",
                                   "regex": r"\d[^\dMm]*[Mm]",
                                   "prices": {datetime(2022, 10, 1): 0.16,
                                              }
                                   },
               "Metoprolol_iv": {"gbo": "BELOC",
                                 "beschr": "Metoprolol i.v.",
                                 "regex": r"\d[^\dBb]*[Bb]",
                                 "prices": {datetime(2022, 10, 1): 4.30,
                                            datetime(2023, 1, 3): 5.12,
                                            }
                                 },
               "Atenolol": {"gbo": "ATEN",
                            "beschr": "Atenolol oral",
                            "regex": r"\d[^\dAa]*[Aa]",
                            "prices": {datetime(2022, 12, 14): 0.25,
                                       datetime(2023, 1, 3): 0.23,
                                       }
                            },
               }

versicherungen = {
    "Selbstzahler": {"1": 2.30,
                     "5": 3.50,
                     "75": 3.50,
                     "DR": 1.00,
                     "346": 3.50,
                     "650": 2.50,
                     "5371": 2.50,
                     "5377": 1.00,
                     "5376": 2.30,
                     "60": 3.50,
                     },
    "KVBI-III": {"1": 2.20,
                 "5": 2.20,
                 "75": 2.20,
                 "DR": 1.00,
                 "346": 2.20,
                 "650": 1.80,
                 "5371": 1.80,
                 "5377": 1.00,
                 "5376": 1.80,
                 "60": 2.20,
                 },
    "PostB": {"1": 1.90,
              "5": 1.90,
              "75": 1.90,
              "DR": 1.00,
              "346": 1.90,
              "650": 1.50,
              "5371": 1.50,
              "5377": 1.00,
              "5376": 1.50,
              "60": 1.90,
              },
    "Basistarif": {"1": 1.20,
                   "5": 1.20,
                   "75": 1.20,
                   "DR": 1.00,
                   "346": 1.20,
                   "650": 1.00,
                   "5371": 1.00,
                   "5377": 1.00,
                   "5376": 1.00,
                   "60": 1.20,
                   },
    "Standardtarif": {"1": 1.80,
                      "5": 1.80,
                      "75": 1.80,
                      "DR": 1.00,
                      "346": 1.80,
                      "650": 1.38,
                      "5371": 1.38,
                      "5377": 1.00,
                      "5376": 1.38,
                      "60": 1.80,
                      },
    "IGEL_alt": {"1": 1.00,
                 "5": 1.00,
                 "75": 1.00,
                 "DR": 1.00,
                 "346": 1.00,
                 "650": 1.00,
                 "5371": 1.80,
                 "5377": 1.00,
                 "5376": 1.80,
                 "60": 1.00,
                 },
    "IGEL": {"1": 1.00,
             "5": 1.00,
             "75": 1.00,
             "DR": 1.00,
             "346": 1.00,
             "650": 1.00,
             "5371": 2.20,
             "5377": 1.00,
             "5376": 1.80,
             "60": 1.00,
             },
}


def get_abrechnungsziffern(df: pd.Series) -> Dict:
    _result = []
    if len(df.Abrechnungsziffern) > 0:
        try:
            _ziffern_list = df.Abrechnungsziffern.split(";")
            _tarif = {}
            for item in _ziffern_list:
                item = item.strip(" ").split(",")
                _tarif[item[0]] = float(item[1])
        except:
            raise KeyError("Die angegebenen Abrechnungsziffern haben nicht das richtige Format.")
    else:
        _tarif = versicherungen[df.Versicherung]
    _gesamtsumme = 0
    for index, ziffer in enumerate(_tarif.keys(), start=1):
        betrag = gbo[ziffer]['preis']
        if ziffer == "DR":
            betrag = get_price_for_gbo(ziffer, df.Rechnungsdatum)
        _dict = {'pos': index,
                 'gbo': ziffer,
                 'beschr': gbo[ziffer]['beschr'],
                 'preis': format_currency(betrag, 'EUR', format='#.00', locale='de_DE',
                                          currency_digits=False),
                 'faktor': _tarif[ziffer],
                 'anzahl': '1',
                 'betrag': format_currency(betrag * _tarif[ziffer], 'EUR', format='#.00 ¤',
                                           locale='de_DE', currency_digits=False),
                 'betrag_raw': betrag * _tarif[ziffer] if df.Rechnungsdatum < datetime(2023, 1, 25) else round(betrag * _tarif[ziffer], 2),
                 }
        _gesamtsumme += _dict['betrag_raw']
        _result.append(_dict)
    return {'tabelle': _result,
            'gesamtsumme': format_currency(_gesamtsumme, 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False),
            'gesamtsumme_raw': _gesamtsumme}


def find_associated_date(input_date: datetime, dates_list: List[datetime]):
    # initialize variables for closest older date and time difference
    closest_date = None
    min_time_diff = None

    # loop through dates in list
    for date in dates_list:
        # calculate time difference between input date and list date
        time_diff = input_date - date

        # check if list date is older than input date
        if time_diff.days > 0:
            # check if this is the first older date found or if it's closer than previous closest date
            if closest_date is None or time_diff < min_time_diff:
                closest_date = date
                min_time_diff = time_diff
    if closest_date is None:
        raise ValueError(
            f"Es konnte kein Medikamentenpreis für das Datum {input_date.strftime('%d.%m.%Y')} gefunden werden.")

    # print closest older date found
    return closest_date


def find_correct_drug_prizes(untersuchungsdatum: datetime, drug_name: str) -> float:
    _dict = medikamente[drug_name]
    return _dict["prices"][find_associated_date(untersuchungsdatum, _dict["prices"].keys())]
