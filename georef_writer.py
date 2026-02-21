"""Build and send a Tapir SetGeoLocation payload from a GeorefData object.

The Tapir API expects north in radians; this module converts from degrees
(the internal representation) back to radians before sending.

SetGeoLocation is wrapped in ACAPI_CallUndoableCommand on the ArchiCAD side,
so any write can be undone normally from within ArchiCAD.
"""

import math
from typing import Callable

from models import GeorefData


def build_payload(data: GeorefData) -> dict:
    """Convert a GeorefData into the dict expected by Tapir SetGeoLocation."""
    return {
        "projectLocation": {
            "longitude": data.project_location.longitude,
            "latitude": data.project_location.latitude,
            "altitude": data.project_location.altitude,
            "north": math.radians(data.project_location.north_deg),  # deg → rad
        },
        "surveyPoint": {
            "position": {
                "eastings": data.survey_point.eastings,
                "northings": data.survey_point.northings,
                "elevation": data.survey_point.elevation,
            },
            "geoReferencingParameters": {
                "crsName": data.geo_ref_params.crs_name,
                "description": data.geo_ref_params.description,
                "geodeticDatum": data.geo_ref_params.geodetic_datum,
                "verticalDatum": data.geo_ref_params.vertical_datum,
                "mapProjection": data.geo_ref_params.map_projection,
                "mapZone": data.geo_ref_params.map_zone,
            },
        },
    }


def write_geolocation(tapir_fn: Callable, data: GeorefData) -> dict:
    """Send the georeferencing data to ArchiCAD via Tapir SetGeoLocation.

    Args:
        tapir_fn: Callable matching ArchicadConnection.tapir(command, params).
        data:     Complete GeorefData to write.

    Returns:
        The raw Tapir ExecutionResult dict.

    Raises:
        RuntimeError: if the Tapir call fails.
    """
    payload = build_payload(data)
    try:
        result = tapir_fn("SetGeoLocation", payload)
    except Exception as exc:
        raise RuntimeError(f"Tapir SetGeoLocation failed: {exc}") from exc
    return result
