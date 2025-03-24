#!C:\Users\Widia\source\repos\Automation_to_GeoJSON_Shapefile_Extract_Image\automate_conversion_env\Scripts\python.exe

import sys

from osgeo.gdal import deprecation_warn

# import osgeo_utils.gdalmove as a convenience to use as a script
from osgeo_utils.gdalmove import *  # noqa
from osgeo_utils.gdalmove import main

deprecation_warn("gdalmove")
sys.exit(main(sys.argv))
