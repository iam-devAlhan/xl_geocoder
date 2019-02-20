# -*- coding: utf-8 -*-
import os
import geocoder
import shapefile
from time import sleep
from requests import Session
from openpyxl import load_workbook, Workbook

### Funkcje -------------------------------------------------------------------

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


if __name__ == "__main__":

### Konfiguracja --------------------------------------------------------------
    
    xls_path = 'dane\\Instalacje_PZ_30092018.xlsx'  # zrodlo danych do geokodowania
    xls_name = os.path.splitext(os.path.basename(xls_path))[0]
    xls_min_row = 2
    xls_max_row = None
    xls_max_column = 5
    output_dir = 'output'
    output_shp_name = xls_name  # moduł shapefile ignoruje rozszerzenia plików
    output_shp_path = os.path.join(output_dir, output_shp_name)
    incorrect_data_xls_name = 'NIEPOPRAWNE_ADRESY_' + xls_name + '.xlsx'
    incorrect_data_xls_path = os.path.join(output_dir, incorrect_data_xls_name)

    delay = 1.2  # opóźnienie zapytania do serwera w sekundach 

    # Konfiguracja atrybutów shp
    fields_config = [
        ['NAZWA', 'C', 255],
        ['UL_NR', 'C', 255],
        ['KOD', 'C', 255],
        ['MIEJSC', 'C', 255],
        ['WOJ', 'C', 255]
    ]

### Instrukcje ----------------------------------------------------------------

    # Odczyt xls z danymi
    wb = load_workbook(xls_path, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(min_row=xls_min_row, max_row=xls_max_row,
                        max_col=xls_max_column, values_only=True)
    # Xls na błędne adresy
    incorrect_data_wb = Workbook()
    incorrect_data_ws = incorrect_data_wb.active
    incorrect_data_ws.title = 'No result'
    incorrect_data_ws.append(['Nazwa', 'ul_nr', 'kod', 'miejscowosc', 'woj',
                             'nr_wiersza', 'gc_status', 'gc_status_code', 'gc_timeout'])


    print '\n' + 'GEOKODOWANIE - START' + '\n'

    with shapefile.Writer(output_shp_path, 1) as shp:

        add_fields_to_shp(shp, fields_config)
        
        with Session() as session:

            for i, row in enumerate(rows):
                
                print i+1

                # Pozycje kolumn w xls
                nazwa = row[0].encode('utf8')       # A - nazwa
                ul_nr = row[1].encode('utf8')       # B - ulica + numer
                kod = row[2].encode('utf8')         # C - kod pocztowy
                miejsc = row[3].encode('utf8')      # D - miejscowosc
                woj = row[4].encode('utf8')         # E - wojewodztwo

                adres = '{0}, {1}'.format(ul_nr, miejsc)
       
                gc = geocoder.osm(adres, session=session)

                # print adres.decode('utf8').encode('cp1250')

                if gc.ok:
                    shp.point(gc.lat, gc.lng)
                    shp.record(nazwa, ul_nr, kod, miejsc, woj)
                    print '    lat: {0}; lng: {1}'.format(gc.lat, gc.lng)
                else:
                    print '    {0} (status:{1}, timeout:{2})'.format(
                        gc.status, gc.status_code, gc.timeout)
                    incorrect_data_ws.append(
                        [nazwa, ul_nr, kod, miejsc, woj,
                         i+1, gc.status, gc.status_code, gc.timeout])
                    try:
                        incorrect_data_wb.save(incorrect_data_xls_path)
                    except:
                        incorrect_data_wb.save(incorrect_data_xls_path.replace(xls_name, xls_name + '_alt'))

                sleep(delay)  # przeciwdziała banowaniu
