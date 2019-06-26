import os
import shapefile
import pycrs


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