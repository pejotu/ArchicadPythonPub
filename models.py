"""Data models for the georeferencing tool."""

from dataclasses import dataclass, field


@dataclass
class ProjectLocation:
    """WGS84 geographic location of the project origin."""
    longitude: float = 0.0   # decimal degrees
    latitude: float = 0.0    # decimal degrees
    altitude: float = 0.0    # meters
    north_deg: float = 0.0   # degrees (Tapir stores as radians; reader/writer convert)


@dataclass
class SurveyPointPosition:
    """Position of the survey point in the project's local projected CRS."""
    eastings: float = 0.0
    northings: float = 0.0
    elevation: float = 0.0   # meters above vertical datum


@dataclass
class GeoReferencingParameters:
    """CRS identification strings, matching IFC IfcProjectedCRS fields."""
    crs_name: str = ""          # CRS identifier (e.g. "ETRS89 / TM35FIN(E,N)")
    description: str = ""       # Informal description
    geodetic_datum: str = ""    # e.g. "European Terrestrial Reference System 1989"
    vertical_datum: str = ""    # e.g. "N2000"
    map_projection: str = ""    # e.g. "Transverse Mercator"
    map_zone: str = ""          # e.g. "35"


@dataclass
class GeorefData:
    """Complete georeferencing state of an ArchiCAD project."""
    project_location: ProjectLocation = field(default_factory=ProjectLocation)
    survey_point: SurveyPointPosition = field(default_factory=SurveyPointPosition)
    geo_ref_params: GeoReferencingParameters = field(default_factory=GeoReferencingParameters)
