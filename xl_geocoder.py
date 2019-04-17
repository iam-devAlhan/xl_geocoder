'''
 -----------------------------------------------------------------------------------------
 Nazwa:     xl_geocoder
 Opis:      Geokoduje adresy z arkuszy kalkulacyjncyh programu Excel.
            Wynik zapisuje w postaci warstwy shp.

 Autor:     Przemek Garasz
 Data utw:  2018-02-20
 Data mod:  2018-06-01
 Wersja:    1.5
 ----------------------------------------------------------------------------------------
'''

import os
import geocoder
import shapefile
import pycrs
import datetime
import re
from time import sleep
from requests import Session
from openpyxl import load_workbook, Workbook
from tools.shp_template_from_xls import get_fields_properties_from_workseet


class FakeGC:
    '''Sztuczna klasa do symulacji statusów geocodera.osm'''

    def __init__(self, ok, status, status_code=-999, timeout=-999, osm=''):
        self.ok = ok
        self.status = status
        self.status_code = status_code
        self.timeout = timeout
        self.osm = osm


def create_empty_shp(path, field_params_list, shapeType):
    '''
    Tworzy pusty plik shp ze strukturą atrubutów
    shp_writer - instancja klasy Writer modułu shapefile
    field_params_list - parametry pola tabeli w postaci listy list [nazwa, typ, rozmiar]
    '''
    with shapefile.Writer(path, shapeType) as shp:
        for field_params in field_params_list:
            shp.field(*field_params)


def add_fields_to_shp(shp_writer, field_params_list):
    '''
    Dodaje atrybuty do shp
    shp_writer - instancja klasy Writer modułu shapefile
    field_params_list - parametry atrybutu w postaci listy list
                        [nazwa, typ, rozmiar]
    '''
    for field_params in field_params_list:
        shp_writer.field(*field_params)


def create_prj_file(path, epsg, proj_name="Unknown"):
    '''Tworzy plik z definicją systemu współrzędnych w formacie ESRI'''
    crs = pycrs.parse.from_epsg_code(epsg)

    crs.name = proj_name
    if os.path.splitext(path)[-1] == '.prj':
        with open(path, "w") as writer:
            writer.write(crs.to_esri_wkt())
    else:
        raise ValueError


def sanitize_value(value, replace_none=False):
    try:
        if replace_none:
            if not value:
                return 'BRAK DANYCH'
        if value:
            return str(value).strip()
        else:
            return ''
    except TypeError:
        return 'ZŁY TYP DANYCH'


def parse_street_name(street_name, name_filter=None, expand_abbrev=None,
                      remove_abbrev=False, building_number_first=False):
    '''
    Przetwarza i filtruje nazwę ulicy.
    (Opcja) Zupełnie odrzuca nazwę jeżeli zawiera ciąg tekstu z listy "filter".
    (Opcja) Usuwa skróty zakończone kropką (ul. Gen. Św. M.).
    (Opcja) Znajduje numer budynku na końcu i przenosi na początek
            (rozwiązanie pod osm).
    '''
    # Szuka słów zakończonych "."
    regex = re.compile(r'\w+\.', re.UNICODE)
    # Szuka numeru budynku na końcu ciągu znaków
    regex2 = re.compile(r'((?<= )\d*)((?<=\d)/|(?<=\d)-|(?<=\d)\\)?(\d+)(?: |/|\\)?([a-zA-Z])?$')

    if name_filter:
        for substring in name_filter:
            if substring in street_name:
                return False
    if expand_abbrev:
        for key in expand_abbrev:
            # TODO match case
            street_name = re.sub(key, expand_abbrev[key], street_name, flags=re.IGNORECASE)
    if remove_abbrev:
        street_name = re.sub(regex, '', street_name).strip()
    if building_number_first:
        try:
            match = re.search(regex2, street_name)
            building_number = match.group()
            mod_number = match.expand(r'\1\2\3\4')
            street_name = mod_number + ', ' + street_name.replace(building_number, '').strip()
        except AttributeError:
            None
    return street_name.strip()


if __name__ == "__main__":

    # Konfiguracja ------------------------------------------------------------

    xls_path = r'demo_data\DPSiPOC.xlsx'  # dane do geokodowania
    xls_name = os.path.splitext(os.path.basename(xls_path))[0]
    xls_has_header = True
    xls_min_row = 2
    xls_max_row = None
    xls_max_column = None
    address_columns_indxs = {'ulica': 4,
                             'miejsc1': 5,
                             'kod': 6,
                             'miejsc2': 7,
                             'powiat': 8,
                             'woj': 9
                             }
    illegal_street_names = [u'dz.', u'ew.', u' działki ', u' nr ', u' obręb ']
    abbrev_dict = {'św.':'świętego'}

    now = datetime.datetime.now()
    timestamp = now.strftime('%Y-%m-%d_%H-%M-%S')
    output_dir = 'output_' + timestamp

    output_shp_name = xls_name  # moduł shapefile ignoruje rozszerzenia plików
    output_shp_path = os.path.join(output_dir, output_shp_name)

    incorrect_data_xls_name = 'NIEPOPRAWNE_ADRESY_' + xls_name + '.xlsx'
    incorrect_data_xls_path = os.path.join(output_dir, incorrect_data_xls_name)

    delay = 1.2  # opóźnienie zapytania do serwera w sekundach

    additional_shp_fields = [
        ['QUERY', 'C', 255],
        ['OSM_ANSW', 'C', 255],
        ['CONFIDENCE', 'F', 5, 2]
    ]

    # Instrukcje --------------------------------------------------------------

    # Odczyt xls z danymi
    wb = load_workbook(xls_path, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(min_row=xls_min_row, max_row=xls_max_row,
                        max_col=xls_max_column, values_only=True)

    # Konfiguracja tabeli shp
    fields_config = get_fields_properties_from_workseet(ws, xls_has_header)
    shp_fields_config = fields_config + additional_shp_fields

    # Xls na błędne adresy
    incorrect_data_wb = Workbook()
    incorrect_data_ws = incorrect_data_wb.active
    incorrect_data_ws.title = 'No result'
    if xls_has_header:  # TODO Przenieść do osobnej funkcji
        column_headers = [field_property[0] for field_property in fields_config]
    else:
        column_headers = ['' for field_property in fields_config]
    incorrect_data_ws.append(column_headers + ['zapytanie', 'gc_status', 'gc_status_code', 'gc_timeout'])

    print('\n' + 'GEOKODOWANIE - START' + '\n')

    with shapefile.Writer(output_shp_path, 1) as shp:

        add_fields_to_shp(shp, shp_fields_config)
        create_prj_file(output_shp_path + '.prj', 4326, 'GCS_WGS_1984')

        with Session() as session:

            for i, row in enumerate(rows):

                print(i + xls_min_row)

                # Odczyt danych z xls
                idx = address_columns_indxs
                ulica     = sanitize_value(row[idx['ulica']])
                miejsc1   = sanitize_value(row[idx['miejsc1']])
                kod       = sanitize_value(row[idx['kod']])
                miejsc2   = sanitize_value(row[idx['miejsc2']])
                powiat    = sanitize_value(row[idx['powiat']])
                woj       = sanitize_value(row[idx['woj']])

                
                # WARUNKI UWZGLĘDNIAJĄCE MAŁE MIEJSCOWOŚCI BEZ NAZW ULIC
                # TODO Zabezpieczyć przed None
                if miejsc1 != '' and ulica != '':           # miejscowosc bez poczty z nazwami ulicami
                    print(f'      dane: {ulica}, {miejsc1}')
                    ul_nr_mod = parse_street_name(street_name=ulica, name_filter=illegal_street_names,
                                                  expand_abbrev=abbrev_dict, remove_abbrev=True, building_number_first=True)

                    adres = ul_nr_mod.strip() + ', ' + miejsc1 + ', powiat ' + powiat
            
                elif miejsc1 != '':                         # miejscowosc bez poczty (tylko numery budynków)
                    print(f'      dane: {miejsc1}')
                    ul_nr_mod = parse_street_name(street_name=miejsc1, name_filter=illegal_street_names,
                                                  building_number_first=True)

                    adres = ul_nr_mod.strip() + ', powiat ' + powiat

                elif miejsc2 != '' and ulica != '':         # miejscowosc z pocztą i ulicami
                    print(f'      dane: {ulica}, {miejsc2}')
                    ul_nr_mod = parse_street_name(street_name=ulica, name_filter=illegal_street_names,
                                                  expand_abbrev=abbrev_dict, remove_abbrev=True, building_number_first=True)
                    if miejsc2 == powiat:
                        powiat = ''
                    else:
                        powiat = ', powiat ' + powiat

                    adres = ul_nr_mod.strip() + ', ' + miejsc2 + powiat


                has_leading_number = re.match('^\d+', adres)  # Poprawny adres musi mieć numer budynku
                print(f'   szukane: {adres}')                    

                if adres and has_leading_number:
                    gc = geocoder.osm(adres, session=session)
                else:
                    gc = FakeGC(False, u"BŁĄD - NIERPAWIDŁOWY ADRES")
                    adres = gc.status


                if gc.ok:
                    confidence = gc.current_result.confidence
                    shp.point(gc.lng, gc.lat)
                    shp.record(*row, adres, gc.osm, confidence)
                    print(f'       lat: {gc.lat}; lng: {gc.lng}; confidence: {confidence}')
                else:
                    print(f'       {gc.status} (status:{gc.status_code}, timeout:{gc.timeout})')
                    incorrect_data_ws.append(row + (adres, gc.status, gc.status_code, gc.timeout))
                    try:
                        incorrect_data_wb.save(incorrect_data_xls_path)
                    except Exception:
                        incorrect_data_wb.save(
                            incorrect_data_xls_path.replace(xls_name, xls_name + '_alt'))

                sleep(delay)  # przeciwdziała banowaniu
