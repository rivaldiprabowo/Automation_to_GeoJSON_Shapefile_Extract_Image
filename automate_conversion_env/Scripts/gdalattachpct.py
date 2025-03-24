#!C:\Users\Widia\source\repos\Automation_to_GeoJSON_Shapefile_Extract_Image\automate_conversion_env\Scripts\python.exe

import sys

from osgeo.gdal import deprecation_warn

# import osgeo_utils.gdalattachpct as a convenience to use as a script
from osgeo_utils.gdalattachpct import *  # noqa
from osgeo_utils.gdalattachpct import main

deprecation_warn("gdalattachpct")
sys.exit(main(sys.argv))
