import os
import shapefile
import pycrs


def create_empty_shp(path, field_params, shapeType):
    """ Creates an empty shp file with attributes from the list

    Args:
        path - (string) - disk location where function will save the shp file
        field_params - (list) - [name, field_type, size]
        shapeType - (int or string) - shape type name or int in accordance with shp specification
                                      e.g. 'POINT' or 1
    """
    with shapefile.Writer(path, shapeType) as shp:
        for field_params in field_params:
            shp.field(*field_params)


def add_fields_to_shp(shp_writer, field_params):
    """ Adds fields to shapefile Writer class instance
    Args:
        shp_writer - Writer class instance of shapefile module
        field_params - (list) - [name, field_type, size]
    """
    for field_params in field_params:
        shp_writer.field(*field_params)


def create_prj_file(path, epsg, proj_name="Unknown"):
    """Creates prj file with coordinates system info in ESRI format"""

    crs = pycrs.parse.from_epsg_code(epsg)

    crs.name = proj_name
    if os.path.splitext(path)[-1] == '.prj':
        with open(path, "w") as writer:
            writer.write(crs.to_esri_wkt())
    else:
        raise ValueError
