"""Resolve an EPSG integer code to the CRS string fields expected by ArchiCAD.

Strategy (in priority order):
  1. pyproj.CRS  — local, no network, covers all well-known EPSG codes.
  2. epsg.io REST API — fallback for fields pyproj cannot supply (mainly
     vertical_datum and map_zone for non-standard CRSs).
  3. Manual override — any field left empty after steps 1-2 can be filled
     directly in the UI before writing to ArchiCAD.

Field mapping to IFC IfcProjectedCRS:
  crs_name        → Name
  description     → Description
  geodetic_datum  → GeodeticDatum
  vertical_datum  → VerticalDatum
  map_projection  → MapProjection
  map_zone        → MapZone
"""

import re
from dataclasses import dataclass


@dataclass
class CRSMetadata:
    """CRS identification strings, ready to be written to ArchiCAD."""
    crs_name: str = ""
    description: str = ""
    geodetic_datum: str = ""
    vertical_datum: str = ""    # rarely in projected CRS; may need manual input
    map_projection: str = ""
    map_zone: str = ""


def from_epsg(code: int) -> CRSMetadata:
    """Return CRS metadata for *code*, trying pyproj then epsg.io.

    Raises:
        ValueError: if the EPSG code is not recognised by either source.
    """
    meta = _from_pyproj(code)

    # If pyproj left any key fields empty, try the network fallback.
    if not meta.crs_name or not meta.map_zone:
        _fill_from_epsg_io(code, meta)

    if not meta.crs_name:
        raise ValueError(f"EPSG:{code} could not be resolved by pyproj or epsg.io")

    return meta


# ---------------------------------------------------------------------------
# pyproj source
# ---------------------------------------------------------------------------

def _from_pyproj(code: int) -> CRSMetadata:
    """Extract CRS metadata using pyproj (no network required)."""
    try:
        from pyproj import CRS
        crs = CRS.from_epsg(code)
    except Exception:
        return CRSMetadata()

    meta = CRSMetadata()
    meta.crs_name = crs.name or ""
    meta.description = (getattr(crs, "remarks", None) or "").strip() or crs.name or ""

    # Geodetic datum — works for projected and geographic CRSs
    try:
        geodetic = getattr(crs, "geodetic_crs", None) or crs
        datum = getattr(geodetic, "datum", None)
        if datum:
            meta.geodetic_datum = datum.name or ""
    except Exception:
        pass

    # Map projection + zone — only meaningful for projected CRSs
    try:
        op = getattr(crs, "coordinate_operation", None)
        if op:
            meta.map_projection = getattr(op, "method_name", None) or ""
            op_name = getattr(op, "name", None) or ""
            meta.map_zone = _extract_zone(op_name, crs.name)
    except Exception:
        pass

    return meta


def _extract_zone(op_name: str, crs_name: str = "") -> str:
    """Heuristically extract a zone string from an operation or CRS name.

    Examples handled:
      "UTM zone 33N"    → "33N"
      "TM35FIN(E,N)"    → "35"
      "GK25FIN"         → "25"
      "zone 6"          → "6"
    """
    # "zone 33N" or "zone 6" patterns
    m = re.search(r"\bzone\s*([\d]+[A-Z]?)", op_name, re.IGNORECASE)
    if m:
        return m.group(1)

    # "TM35FIN" → 35 or "GK25FIN" → 25
    m = re.search(r"[A-Za-z]+(\d{1,3})[A-Za-z]", op_name)
    if m:
        return m.group(1)

    # Fall back to CRS name
    m = re.search(r"\bzone\s*([\d]+[A-Z]?)", crs_name, re.IGNORECASE)
    if m:
        return m.group(1)

    m = re.search(r"[A-Za-z]+(\d{1,3})[A-Za-z]", crs_name)
    if m:
        return m.group(1)

    return ""


# ---------------------------------------------------------------------------
# epsg.io fallback
# ---------------------------------------------------------------------------

_EPSG_IO_BASE = "https://epsg.io/{code}.json"


def _fill_from_epsg_io(code: int, meta: CRSMetadata) -> None:
    """Attempt to fill empty metadata fields from the epsg.io REST API.

    Silently ignores network or parsing errors so the caller always gets
    whatever pyproj managed to provide.
    """
    try:
        import requests
        resp = requests.get(_EPSG_IO_BASE.format(code=code), timeout=6)
        if resp.status_code != 200:
            return
        data = resp.json()
        results = data.get("results", [])
        if not results:
            return
        r = results[0]

        if not meta.crs_name:
            meta.crs_name = r.get("name", "")

        if not meta.description:
            meta.description = r.get("name", "")

        # epsg.io exposes "area" and "kind"; vertical datum info isn't directly
        # in the projected-CRS record, but the "bbox" / "unit" fields exist.
        # At minimum fill map_zone from the name if still empty.
        if not meta.map_zone:
            meta.map_zone = _extract_zone(r.get("name", ""))

    except Exception:
        pass
