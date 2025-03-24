#!C:\Users\Widia\source\repos\Automation_to_GeoJSON_Shapefile_Extract_Image\automate_conversion_env\Scripts\python.exe

import sys

from osgeo.gdal import deprecation_warn

# import osgeo_utils.gdal_proximity as a convenience to use as a script
from osgeo_utils.gdal_proximity import *  # noqa
from osgeo_utils.gdal_proximity import main

deprecation_warn("gdal_proximity")
sys.exit(main(sys.argv))
