"""""" # start delvewheel patch
def _delvewheel_patch_1_6_0():
    import os
    libs_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir, 'pyogrio.libs'))
    if os.path.isdir(libs_dir):
        os.add_dll_directory(libs_dir)


_delvewheel_patch_1_6_0()
del _delvewheel_patch_1_6_0
# end delvewheel patch

try:
    # we try importing shapely, to ensure it is imported (and it can load its
    # own GEOS copy) before we load GDAL and its linked GEOS
    import shapely  # noqa
    import shapely.geos  # noqa
except Exception:
    pass

from pyogrio.core import (
    list_drivers,
    detect_write_driver,
    list_layers,
    read_bounds,
    read_info,
    set_gdal_config_options,
    get_gdal_config_option,
    get_gdal_data_path,
    __gdal_version__,
    __gdal_version_string__,
    __gdal_geos_version__,
)
from pyogrio.raw import read_arrow, open_arrow, write_arrow
from pyogrio.geopandas import read_dataframe, write_dataframe
from pyogrio._version import get_versions


__version__ = get_versions()["version"]
del get_versions

__all__ = [
    "list_drivers",
    "detect_write_driver",
    "list_layers",
    "read_bounds",
    "read_info",
    "set_gdal_config_options",
    "get_gdal_config_option",
    "get_gdal_data_path",
    "read_arrow",
    "open_arrow",
    "write_arrow",
    "read_dataframe",
    "write_dataframe",
    "__gdal_version__",
    "__gdal_version_string__",
    "__gdal_geos_version__",
    "__version__",
]
