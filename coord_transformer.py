"""Coordinate transformation utilities using pyproj.

Primary use-case: convert survey point coordinates (eastings, northings) in a
local projected CRS to WGS84 geographic coordinates (longitude, latitude) for
writing to ArchiCAD's projectLocation fields.

Note: always_xy=True is set globally so coordinates are always ordered
(easting/longitude, northing/latitude) regardless of the EPSG axis convention.
This avoids surprises with CRSs that define axes in northing-first order.
"""

from typing import Optional, Tuple


class CoordTransformer:
    """One-shot transformer between two EPSG coordinate reference systems."""

    def __init__(self, src_epsg: int, dst_epsg: int):
        from pyproj import Transformer
        self.src_epsg = src_epsg
        self.dst_epsg = dst_epsg
        self._t = Transformer.from_crs(
            f"EPSG:{src_epsg}",
            f"EPSG:{dst_epsg}",
            always_xy=True,
        )

    def transform(
        self,
        x: float,
        y: float,
        z: Optional[float] = None,
    ) -> Tuple:
        """Transform a coordinate pair (or triple) from src to dst CRS.

        With always_xy=True:
          x = easting / longitude
          y = northing / latitude

        Returns a (x, y) or (x, y, z) tuple in the destination CRS.
        """
        if z is not None:
            return self._t.transform(x, y, z)
        return self._t.transform(x, y)


def survey_to_wgs84(
    eastings: float,
    northings: float,
    src_epsg: int,
) -> Tuple[float, float]:
    """Convert survey point local-CRS coordinates to WGS84 lon/lat (EPSG:4326).

    This is the primary automatic conversion: the result is used to populate
    ArchiCAD's projectLocation.longitude and projectLocation.latitude fields.

    Args:
        eastings:  Easting in the local projected CRS (metres).
        northings: Northing in the local projected CRS (metres).
        src_epsg:  EPSG code of the local projected CRS (e.g. 3067).

    Returns:
        (longitude, latitude) in decimal degrees (WGS84 / EPSG:4326).
    """
    t = CoordTransformer(src_epsg, 4326)
    lon, lat = t.transform(eastings, northings)
    return lon, lat
