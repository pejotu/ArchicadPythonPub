"""Read current georeferencing data from ArchiCAD via Tapir GetGeoLocation.

The Tapir API stores the north direction in radians; this module converts it
to degrees so the rest of the application always works in degrees.
"""

import math
from typing import Callable

from models import (
    GeorefData,
    GeoReferencingParameters,
    ProjectLocation,
    SurveyPointPosition,
)


def read_geolocation(tapir_fn: Callable) -> GeorefData:
    """Call Tapir GetGeoLocation and return a normalised GeorefData.

    Args:
        tapir_fn: Callable matching ArchicadConnection.tapir(command, params).

    Returns:
        GeorefData with north expressed in degrees (converted from radians).

    Raises:
        RuntimeError: if the Tapir call fails or returns an unexpected structure.
    """
    try:
        response = tapir_fn("GetGeoLocation")
    except Exception as exc:
        raise RuntimeError(f"Tapir GetGeoLocation failed: {exc}") from exc

    data = GeorefData()

    pl = response.get("projectLocation") or {}
    if pl:
        data.project_location = ProjectLocation(
            longitude=float(pl.get("longitude", 0.0)),
            latitude=float(pl.get("latitude", 0.0)),
            altitude=float(pl.get("altitude", 0.0)),
            north_deg=math.degrees(float(pl.get("north", 0.0))),
        )

    sp = response.get("surveyPoint") or {}
    if sp:
        pos = sp.get("position") or {}
        data.survey_point = SurveyPointPosition(
            eastings=float(pos.get("eastings", 0.0)),
            northings=float(pos.get("northings", 0.0)),
            elevation=float(pos.get("elevation", 0.0)),
        )
        gp = sp.get("geoReferencingParameters") or {}
        data.geo_ref_params = GeoReferencingParameters(
            crs_name=gp.get("crsName", ""),
            description=gp.get("description", ""),
            geodetic_datum=gp.get("geodeticDatum", ""),
            vertical_datum=gp.get("verticalDatum", ""),
            map_projection=gp.get("mapProjection", ""),
            map_zone=gp.get("mapZone", ""),
        )

    return data
