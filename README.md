# xslx2kml
Allows you to generate a kml file from xslx with WKT format for import to google maps your geographical objects

## Getting Started
set the parameter values in the settings.ini:
- inputFilePath - full path to the input xlsx file
- outputFilePath - full path to the output kml file
- firstRowWithData - number or row with the start data (if you have a column's headers in the input file, then set the value = 2, otherwise - 1)
- columnWithPointCoords - number of column with the point coordinates in WKT format (column numbering starts from 0)
- columnWithGeometry - number of column with the geometry in WKT format
- columnWithName - number of column with the name of geographical object
- columnWithDescription - number of column with the description of geographical object

# Acknowledgments
- About KML format - https://developers.google.com/kml
- About WKT format - https://ru.wikipedia.org/wiki/WKT 
- Pykml documentation -  https://pythonhosted.org/pykml/
