#!C:\Users\Widia\source\repos\Automation_to_GeoJSON_Shapefile_Extract_Image\automate_conversion_env\Scripts\python.exe

import sys

from osgeo.gdal import deprecation_warn

# import osgeo_utils.rgb2pct as a convenience to use as a script
from osgeo_utils.rgb2pct import *  # noqa
from osgeo_utils.rgb2pct import main

deprecation_warn("rgb2pct")
sys.exit(main(sys.argv))
