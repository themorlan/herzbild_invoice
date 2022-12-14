from typing import Dict
from babel.numbers import format_currency

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
    "IGEL": {"1": 1.00,
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
}


def get_abrechnungsziffern(versicherung: str, abrechnungsziffern: str) -> Dict:
    _result = []
    if len(abrechnungsziffern) > 0:
        try:
            _ziffern_list = abrechnungsziffern.split(";")
            _tarif = {}
            for item in _ziffern_list:
                item = item.strip(" ").split(",")
                _tarif[item[0]] = float(item[1])
        except:
            raise KeyError("Die angegebenen Abrechnungsziffern haben nicht das richtige Format.")
    else:
        _tarif = versicherungen[versicherung]
    _gesamtsumme = 0
    for index, ziffer in enumerate(_tarif.keys(), start=1):
        _dict = {'pos': index,
                 'gbo': ziffer,
                 'beschr': gbo[ziffer]['beschr'],
                 'preis': format_currency(gbo[ziffer]['preis'], 'EUR', format='#.00', locale='de_DE',
                                          currency_digits=False),
                 'faktor': _tarif[ziffer],
                 'anzahl': '1',
                 'betrag': format_currency(gbo[ziffer]['preis'] * _tarif[ziffer], 'EUR', format='#.00 ¤',
                                           locale='de_DE', currency_digits=False),
                 'betrag_raw': gbo[ziffer]['preis'] * _tarif[ziffer],
                 }
        _gesamtsumme += gbo[ziffer]['preis'] * _tarif[ziffer]
        _result.append(_dict)
    return {'tabelle': _result,
            'gesamtsumme': format_currency(_gesamtsumme, 'EUR', format='#.00 ¤', locale='de_DE', currency_digits=False),
            'gesamtsumme_raw': _gesamtsumme}
