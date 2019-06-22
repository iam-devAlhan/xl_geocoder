# xl_geocoder

## About

This is a Python app that geocodes address data* stored in Excel spreadsheets (Office Open XML) and saves it as a shp file.

It uses [openpyxl](https://openpyxl.readthedocs.io) to read xlsx spreadsheets, [geocoder](https://pypi.org/project/geocoder/) package to query OSM's [Nominatim](https://nominatim.openstreetmap.org/) search engine and [pyshp](https://pypi.org/project/pyshp/#overview) to write shapefiles to disk.

This is my first venture into making my code public. I'm certain that there are a lot of changes to be made. It does it's intended purpose, but please consider this a work in progress and an exercise in learning Python.

*currently, it's tailored for Polish addresses, but this will change in the future

## Configuration

Configuration is done through a *config.yaml* file (an example is provided in the repo). 

YAML is parsed only by means of *pyyaml*, so please stick to the provided template.

Properties are as fallows:

- xls
    - path - *[string]* - path to the xls file that you want to geocode ***[mandatory]***
    - has_header - *[bool]*
        - *True* - use first row as shp field names
        - *False* - use spreadsheet's column letters as shp field names
    - min_row - *[int]* - row number to start from ***[mandatory]***
    - max_row - *[int]* - row number to finish on 
    - max_column - *[int]* - index of the last column (starting from the left) that you want to include from the spreadsheet. The rest will be skipped.

- address
    - col_indxs - *[dict]* -
    this dictionary allows you to indicate which spreadsheet column contains which part of the address string. Put the column indexes *[int]* as values of appropriate keys ***[mandatory]***
    - illegal_street_names - *[list]* - you can provide a list of illegal substrings, that when found in the address, will  automatically bypass the query and mark the row as incorrect address
    - abbrev_expansions - *[dict]* - this allows you to expand abbreviations in addresses. Many abbreviations, like personal titles, are stored in their full form in OSM's db. Expanding them in queries improves positive search ratio. Provide the abbreviation as key and its expansion as value.
    - remove_abbrev - *[bool]* - removes all substrings that end with a dot from the address, abbreviation expansion is done beforehand

- strict_search - *[bool]*
    - *True* - if geocoding service does not recognize the address it's marked as *not found*
    - *False* - if geocoding service does not recognize the address, the program removes the part before the comma and tries again until it runs out of parts to cut, which means that it reached the county level. E.g.:
        - 12, ul. Wałbrzyska, 80-985 Gdańsk, pow. Gdańsk
        - ul. Wałbrzyska, 80-985 Gdańsk, pow. Gdańsk
        - 80-985 Gdańsk, pow. Gdańsk
        - pow. Gdańsk

        Use this option if you're willing to correct the location in the output shapefile in the future.

Anything that's not marked as ***[mandatory]*** can be set to *null*. E.g.:
- *illegal_street_names: null*

## TODO
 - translation
 - code refactoring