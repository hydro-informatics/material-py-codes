from gdal import ogr
from gdal import osr
import urllib
import os


def create_shp(shp_file_dir, overwrite=True, *args, **kwargs):
    """
    Create a new shapefile with a defined geometry type (optional)
    :param shp_file_dir: STR of the (relative shapefile directory (ends on ".shp")
    :param overwrite: [optional] BOOL - if True, existing files are overwritten
    :kwarg layer_name: [optional] STR of the layer_name - if None: no layer will be created
    :kwarg layer_type: [optional] STR ("point, "line", or "polygon") of the layer_name - if None: no layer will be created
    :output: ogr shapefile
    """
    shp_driver = ogr.GetDriverByName("ESRI Shapefile")

    # check if output file exists if yes delete it
    if os.path.exists(shp_file_dir) and overwrite:
        shp_driver.DeleteDataSource(shp_file_dir)

    # create and return new shapefile object
    new_shp = shp_driver.CreateDataSource(shp_file_dir)

    # create layer if layer_name and layer_type are provided
    if kwargs.get("layer_name") and kwargs.get("layer_type"):
        # create dictionary of ogr.SHP-TYPES
        geometry_dict = {"point": ogr.wkbPoint,
                         "points": ogr.wkbMultiPoint,
                         "line": ogr.wkbMultiLineString,
                         "polygon": ogr.wkbMultiPolygon}
        # create layer
        try:
            new_shp.CreateLayer(str(kwargs.get("layer_name")),
                                geom_type=geometry_dict[str(kwargs.get("layer_type").lower())])
        except KeyError:
            print("Error: Invalid layer_type provided (must be 'point', 'line', or 'polygon').")
        except TypeError:
            print("Error: layer_name and layer_type must be string.")
    return new_shp


def get_esriwkt(epsg):
    """
    Get esriwkt-formatted spatial references with epsg code online
    Usage: get_esriwkt(4326)
    :param epsg: Int of epsg
    :output: str containing esriwkt (if error: default epsg=4326 is used)
    """
    try:
        with urllib.request.urlopen("http://spatialreference.org/ref/epsg/{0}/esriwkt/".format(epsg)) as response:
            return str(response.read()).strip("b").strip("'")
    except Exception:
        pass
    try:
        with urllib.request.urlopen(
                "http://spatialreference.org/ref/sr-org/epsg{0}-wgs84-web-mercator-auxiliary-sphere/esriwkt/".format(
                    epsg)) as response:
            return str(response.read()).strip("b").strip("'")
        # sr-org codes are available at "https://spatialreference.org/ref/sr-org/{0}/esriwkt/".format(epsg)
        # for example EPSG:3857 = SR-ORG:6864 -> https://spatialreference.org/ref/sr-org/6864/esriwkt/ = EPSG:3857
    except Exception:
        print("ERROR: Could not find epsg code on spatialreference.org. Returning default WKT(epsg=4326).")
        return 'GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137,298.257223563]],PRIMEM["Greenwich",0],UNIT["Degree",0.017453292519943295],UNIT["Meter",1]]'


def get_wkt(epsg, wkt_format="esriwkt"):
    """
    Get WKT-formatted projection information for an epsg code using the osr library
    :param epsg: Int of epsg
    :kwarg wkt_format: Str of wkt format (default is esriwkt for shapefile projections)
    :output: str containing WKT (if error: default epsg=4326 is used)
    """
    default = 'GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137,298.257223563]],PRIMEM["Greenwich",0],UNIT["Degree",0.017453292519943295],UNIT["Meter",1]]'
    spatial_ref = osr.SpatialReference()
    try:
        spatial_ref.ImportFromEPSG(epsg)
    except TypeError:
        print("ERROR: epsg must be integer. Returning default WKT(epsg=4326).")
        return default
    except Exception:
        print("ERROR: epsg number does not exist. Returning default WKT(epsg=4326).")
        return default
    if wkt_format == "esriwkt":
        spatial_ref.MorphToESRI()
    return spatial_ref.ExportToPrettyWkt()
