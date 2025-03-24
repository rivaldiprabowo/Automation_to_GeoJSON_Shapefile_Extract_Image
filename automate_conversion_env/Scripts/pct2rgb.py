#!C:\Users\Widia\source\repos\Automation_to_GeoJSON_Shapefile_Extract_Image\automate_conversion_env\Scripts\python.exe

import sys

from osgeo.gdal import deprecation_warn

# import osgeo_utils.pct2rgb as a convenience to use as a script
from osgeo_utils.pct2rgb import *  # noqa
from osgeo_utils.pct2rgb import main

deprecation_warn("pct2rgb")
sys.exit(main(sys.argv))
