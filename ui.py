"""PyQt5 GUI for the ArchiCAD Georeferencing Tool.

Layout (left-to-right, top-to-bottom):
  ┌ Header: title + connection indicator ──────────────────────────────────┐
  │ ┌ Current Values (left) ─┐  ┌ Set New Values (right, scrollable) ────┐ │
  │ │ [Read from ArchiCAD]   │  │ EPSG: [____] [Lookup CRS & Lon/Lat]   │ │
  │ │                        │  │ ── Survey Point ──────────────────────  │ │
  │ │ read-only display      │  │  Eastings / Northings / Elevation      │ │
  │ │ (monospace, dark)      │  │ ── Project Location (WGS84) ──────────  │ │
  │ │                        │  │  Longitude / Latitude / Altitude / N   │ │
  │ │                        │  │ ── CRS Metadata ──────────────────────  │ │
  │ │                        │  │  Name / Desc / Datum / Zone / …       │ │
  │ └────────────────────────┘  └────────────────────────────────────────┘ │
  │ ┌ Preview ───────────────────────────────────────────────────────────┐ │
  │ │ (JSON payload shown after Preview is clicked)                      │ │
  │ └────────────────────────────────────────────────────────────────────┘ │
  │ [Populate from Current]    [Preview Changes]   [Write to ArchiCAD]     │
  └────────────────────────────────────────────────────────────────────────┘
"""

import json

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from connection import ArchicadConnection
from coord_transformer import survey_to_wgs84
from crs_metadata import from_epsg as resolve_epsg
from georef_reader import read_geolocation
from georef_writer import build_payload, write_geolocation
from models import GeorefData, GeoReferencingParameters, ProjectLocation, SurveyPointPosition


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

_BTN_PRIMARY = """
QPushButton {
    background-color: #0078D4; color: white;
    border: none; border-radius: 4px;
    font-weight: bold; font-size: 11px; padding: 6px 14px;
}
QPushButton:hover   { background-color: #106EBE; }
QPushButton:pressed { background-color: #005A9E; }
QPushButton:disabled { background-color: #444; color: #888; }
"""

_BTN_SECONDARY = """
QPushButton {
    background-color: #4A4A5A; color: #D4D4D4;
    border: none; border-radius: 4px;
    font-size: 11px; padding: 6px 14px;
}
QPushButton:hover   { background-color: #5A5A6A; }
QPushButton:pressed { background-color: #3A3A4A; }
QPushButton:disabled { background-color: #333; color: #777; }
"""

_BTN_WRITE = """
QPushButton {
    background-color: #107C10; color: white;
    border: none; border-radius: 4px;
    font-weight: bold; font-size: 11px; padding: 6px 14px;
}
QPushButton:hover   { background-color: #1A8C1A; }
QPushButton:pressed { background-color: #0A6A0A; }
QPushButton:disabled { background-color: #333; color: #777; }
"""

_DARK_TEXT = """
QTextEdit {
    background-color: #1E1E1E; color: #D4D4D4;
    border: 1px solid #3E3E42; border-radius: 4px; padding: 8px;
}
"""

_INPUT = """
QLineEdit {
    background-color: #2D2D30; color: #D4D4D4;
    border: 1px solid #3E3E42; border-radius: 3px; padding: 4px 8px;
}
QLineEdit:focus { border: 1px solid #0078D4; }
"""

_GROUP = """
QGroupBox {
    color: #9CDCFE; font-weight: bold;
    border: 1px solid #3E3E42; border-radius: 4px;
    margin-top: 10px; padding-top: 6px;
}
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
"""

_MONO = QFont("Courier", 9)


# ---------------------------------------------------------------------------
# Worker threads — keep ArchiCAD calls off the UI thread
# ---------------------------------------------------------------------------

class _ConnectWorker(QThread):
    connected = pyqtSignal(object)
    failed = pyqtSignal(str)

    def run(self):
        try:
            conn = ArchicadConnection().connect()
            self.connected.emit(conn)
        except Exception as exc:
            self.failed.emit(str(exc))


class _ReadWorker(QThread):
    result = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, tapir_fn):
        super().__init__()
        self._tapir = tapir_fn

    def run(self):
        try:
            self.result.emit(read_geolocation(self._tapir))
        except Exception as exc:
            self.failed.emit(str(exc))


class _LookupWorker(QThread):
    result = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, epsg: int, eastings=None, northings=None):
        super().__init__()
        self._epsg = epsg
        self._eastings = eastings
        self._northings = northings

    def run(self):
        try:
            meta = resolve_epsg(self._epsg)
            lon = lat = None
            if self._eastings is not None and self._northings is not None:
                lon, lat = survey_to_wgs84(self._eastings, self._northings, self._epsg)
            self.result.emit({"meta": meta, "lon": lon, "lat": lat})
        except Exception as exc:
            self.failed.emit(str(exc))


class _WriteWorker(QThread):
    success = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, tapir_fn, data: GeorefData):
        super().__init__()
        self._tapir = tapir_fn
        self._data = data

    def run(self):
        try:
            self.success.emit(write_geolocation(self._tapir, self._data))
        except Exception as exc:
            self.failed.emit(str(exc))


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class GeorefUI(QMainWindow):
    """Main window of the ArchiCAD Georeferencing Tool."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ArchiCAD Georeferencing Tool")
        self.setGeometry(100, 100, 1150, 820)
        self.setStyleSheet("background-color: #252526; color: #D4D4D4;")

        self._connection = None
        self._current_data: GeorefData | None = None
        self._workers: list[QThread] = []   # prevent GC of running threads

        self._build_ui()
        self._start_connection()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        root_widget = QWidget()
        self.setCentralWidget(root_widget)
        root = QVBoxLayout(root_widget)
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 10)

        root.addLayout(self._build_header())

        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(6)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setSizes([420, 680])
        root.addWidget(splitter, stretch=3)

        root.addWidget(self._build_preview_section(), stretch=1)
        root.addLayout(self._build_button_row())

        self.statusBar().showMessage("Connecting to ArchiCAD…")
        self.statusBar().setStyleSheet("color: #9CDCFE;")

    def _build_header(self) -> QHBoxLayout:
        row = QHBoxLayout()
        title = QLabel("ArchiCAD Georeferencing Tool")
        f = QFont()
        f.setPointSize(12)
        f.setBold(True)
        title.setFont(f)
        row.addWidget(title)
        row.addStretch()
        self._conn_label = QLabel("● Connecting…")
        self._conn_label.setStyleSheet("color: #CCCC00;")
        row.addWidget(self._conn_label)
        return row

    # ---- Left panel: current values display ----

    def _build_left_panel(self) -> QGroupBox:
        panel = QGroupBox("Current Values (from ArchiCAD)")
        panel.setStyleSheet(_GROUP)
        layout = QVBoxLayout(panel)
        layout.setSpacing(6)

        self._read_btn = QPushButton("Read from ArchiCAD")
        self._read_btn.setStyleSheet(_BTN_PRIMARY)
        self._read_btn.setEnabled(False)
        self._read_btn.clicked.connect(self._read_from_archicad)
        layout.addWidget(self._read_btn)

        self._current_text = QTextEdit()
        self._current_text.setReadOnly(True)
        self._current_text.setFont(_MONO)
        self._current_text.setStyleSheet(_DARK_TEXT)
        self._current_text.setPlaceholderText(
            "Connect to ArchiCAD and click 'Read from ArchiCAD'…"
        )
        layout.addWidget(self._current_text)
        return panel

    # ---- Right panel: editable form ----

    def _build_right_panel(self) -> QGroupBox:
        panel = QGroupBox("Set New Values")
        panel.setStyleSheet(_GROUP)
        outer = QVBoxLayout(panel)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background: transparent; border: none;")

        content = QWidget()
        content.setStyleSheet("background: transparent;")
        vbox = QVBoxLayout(content)
        vbox.setSpacing(10)

        vbox.addWidget(self._build_epsg_section())
        vbox.addWidget(self._build_survey_section())
        vbox.addWidget(self._build_location_section())
        vbox.addWidget(self._build_crs_section())
        vbox.addStretch()

        scroll.setWidget(content)
        outer.addWidget(scroll)
        return panel

    def _build_epsg_section(self) -> QGroupBox:
        grp = QGroupBox("EPSG Code")
        grp.setStyleSheet(_GROUP)
        row = QHBoxLayout(grp)

        row.addWidget(QLabel("EPSG:"))
        self._epsg_input = QLineEdit()
        self._epsg_input.setPlaceholderText("e.g. 3067")
        self._epsg_input.setMaximumWidth(110)
        self._epsg_input.setStyleSheet(_INPUT)
        self._epsg_input.returnPressed.connect(self._lookup_epsg)
        row.addWidget(self._epsg_input)

        self._lookup_btn = QPushButton("Lookup CRS && Compute Lon/Lat")
        self._lookup_btn.setStyleSheet(_BTN_SECONDARY)
        self._lookup_btn.setToolTip(
            "Fetch CRS metadata for the entered EPSG code and compute\n"
            "WGS84 longitude/latitude from the survey point coordinates."
        )
        self._lookup_btn.clicked.connect(self._lookup_epsg)
        row.addWidget(self._lookup_btn)
        row.addStretch()
        return grp

    def _build_survey_section(self) -> QGroupBox:
        grp = QGroupBox("Survey Point (Local Projected CRS)")
        grp.setStyleSheet(_GROUP)
        form = QFormLayout(grp)
        self._eastings_input = self._lineedit("Easting coordinate in the local projected CRS")
        self._northings_input = self._lineedit("Northing coordinate in the local projected CRS")
        self._sp_elev_input = self._lineedit("Elevation above the vertical datum (metres)")
        form.addRow("Eastings:", self._eastings_input)
        form.addRow("Northings:", self._northings_input)
        form.addRow("Elevation (m):", self._sp_elev_input)
        return grp

    def _build_location_section(self) -> QGroupBox:
        grp = QGroupBox("Project Location  —  WGS84 / EPSG:4326  (auto-computed from survey point)")
        grp.setStyleSheet(_GROUP)
        form = QFormLayout(grp)

        lon_lbl = QLabel("Longitude (°):")
        lon_lbl.setToolTip("Auto-populated when 'Lookup CRS & Compute Lon/Lat' is run")
        lat_lbl = QLabel("Latitude (°):")
        lat_lbl.setToolTip("Auto-populated when 'Lookup CRS & Compute Lon/Lat' is run")

        self._lon_input = self._lineedit("Decimal degrees — auto-filled from EPSG lookup")
        self._lat_input = self._lineedit("Decimal degrees — auto-filled from EPSG lookup")
        self._alt_input = self._lineedit("Altitude in metres")
        self._north_input = self._lineedit("North direction in degrees (stored as radians in ArchiCAD)")

        form.addRow(lon_lbl, self._lon_input)
        form.addRow(lat_lbl, self._lat_input)
        form.addRow("Altitude (m):", self._alt_input)
        form.addRow("North (°):", self._north_input)
        return grp

    def _build_crs_section(self) -> QGroupBox:
        grp = QGroupBox("CRS Metadata  (auto-filled from EPSG lookup; all fields editable)")
        grp.setStyleSheet(_GROUP)
        form = QFormLayout(grp)

        self._crs_name_input = self._lineedit(
            "CRS identifier — maps to IFC IfcProjectedCRS.Name"
        )
        self._desc_input = self._lineedit(
            "Informal description — maps to IfcProjectedCRS.Description"
        )
        self._geodetic_input = self._lineedit(
            "Geodetic datum name — maps to IfcProjectedCRS.GeodeticDatum"
        )
        self._vertical_input = self._lineedit(
            "Vertical datum name — maps to IfcProjectedCRS.VerticalDatum\n"
            "pyproj cannot resolve this automatically; enter manually if needed (e.g. 'N2000')"
        )
        self._projection_input = self._lineedit(
            "Map projection method — maps to IfcProjectedCRS.MapProjection"
        )
        self._zone_input = self._lineedit(
            "Map zone identifier — maps to IfcProjectedCRS.MapZone"
        )

        form.addRow("CRS Name:", self._crs_name_input)
        form.addRow("Description:", self._desc_input)
        form.addRow("Geodetic Datum:", self._geodetic_input)
        form.addRow("Vertical Datum:", self._vertical_input)
        form.addRow("Map Projection:", self._projection_input)
        form.addRow("Map Zone:", self._zone_input)
        return grp

    def _build_preview_section(self) -> QGroupBox:
        grp = QGroupBox("Preview  —  payload that will be sent to ArchiCAD")
        grp.setStyleSheet(_GROUP)
        layout = QVBoxLayout(grp)
        self._preview_text = QTextEdit()
        self._preview_text.setReadOnly(True)
        self._preview_text.setFont(_MONO)
        self._preview_text.setStyleSheet(_DARK_TEXT)
        self._preview_text.setMaximumHeight(190)
        self._preview_text.setPlaceholderText(
            "Click 'Preview Changes' to see the JSON payload before writing…"
        )
        layout.addWidget(self._preview_text)
        return grp

    def _build_button_row(self) -> QHBoxLayout:
        row = QHBoxLayout()

        self._populate_btn = QPushButton("Populate from Current")
        self._populate_btn.setStyleSheet(_BTN_SECONDARY)
        self._populate_btn.setToolTip(
            "Fill all form fields with the values last read from ArchiCAD"
        )
        self._populate_btn.setEnabled(False)
        self._populate_btn.clicked.connect(self._populate_from_current)
        row.addWidget(self._populate_btn)

        row.addStretch()

        self._preview_btn = QPushButton("Preview Changes")
        self._preview_btn.setStyleSheet(_BTN_SECONDARY)
        self._preview_btn.clicked.connect(self._preview_changes)
        row.addWidget(self._preview_btn)

        self._write_btn = QPushButton("Write to ArchiCAD")
        self._write_btn.setStyleSheet(_BTN_WRITE)
        self._write_btn.setEnabled(False)
        self._write_btn.setToolTip(
            "Write the values shown in the Preview to ArchiCAD.\n"
            "The action is undoable via ArchiCAD's Undo."
        )
        self._write_btn.clicked.connect(self._write_to_archicad)
        row.addWidget(self._write_btn)

        return row

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _lineedit(self, tooltip: str = "") -> QLineEdit:
        w = QLineEdit()
        w.setStyleSheet(_INPUT)
        if tooltip:
            w.setToolTip(tooltip)
        return w

    @staticmethod
    def _try_float(text: str):
        try:
            return float(text.strip())
        except (ValueError, AttributeError):
            return None

    def _track(self, worker: QThread):
        """Keep a reference so the worker is not garbage-collected."""
        self._workers.append(worker)

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def _start_connection(self):
        w = _ConnectWorker()
        w.connected.connect(self._on_connected)
        w.failed.connect(self._on_connect_failed)
        self._track(w)
        w.start()

    def _on_connected(self, conn):
        self._connection = conn
        self._conn_label.setText("● Connected")
        self._conn_label.setStyleSheet("color: #4CAF50;")
        self._read_btn.setEnabled(True)
        self.statusBar().showMessage("Connected — reading current values…")
        self._read_from_archicad()   # auto-read on connect

    def _on_connect_failed(self, msg: str):
        self._conn_label.setText("● Not Connected")
        self._conn_label.setStyleSheet("color: #E81B23;")
        self.statusBar().showMessage("Connection failed")
        QMessageBox.critical(self, "Connection Error", msg)

    # ------------------------------------------------------------------
    # Read
    # ------------------------------------------------------------------

    def _read_from_archicad(self):
        if not self._connection:
            return
        self._read_btn.setEnabled(False)
        self.statusBar().showMessage("Reading from ArchiCAD…")

        w = _ReadWorker(self._connection.tapir)
        w.result.connect(self._on_read_success)
        w.failed.connect(self._on_read_failed)
        w.finished.connect(lambda: self._read_btn.setEnabled(bool(self._connection)))
        self._track(w)
        w.start()

    def _on_read_success(self, data: GeorefData):
        self._current_data = data
        self._current_text.setPlainText(self._format_data(data))
        self._populate_btn.setEnabled(True)
        self.statusBar().showMessage("Read successful")

    def _on_read_failed(self, msg: str):
        self.statusBar().showMessage("Read failed")
        QMessageBox.warning(self, "Read Error", f"Failed to read from ArchiCAD:\n\n{msg}")

    @staticmethod
    def _format_data(data: GeorefData) -> str:
        pl = data.project_location
        sp = data.survey_point
        gp = data.geo_ref_params
        return "\n".join([
            "── Project Location ─────────────────────",
            f"  Longitude:       {pl.longitude:.6f} °",
            f"  Latitude:        {pl.latitude:.6f} °",
            f"  Altitude:        {pl.altitude:.3f} m",
            f"  North:           {pl.north_deg:.4f} °",
            "",
            "── Survey Point ─────────────────────────",
            f"  Eastings:        {sp.eastings:.3f}",
            f"  Northings:       {sp.northings:.3f}",
            f"  Elevation:       {sp.elevation:.3f} m",
            "",
            "── CRS Metadata ─────────────────────────",
            f"  CRS Name:        {gp.crs_name}",
            f"  Description:     {gp.description}",
            f"  Geodetic Datum:  {gp.geodetic_datum}",
            f"  Vertical Datum:  {gp.vertical_datum}",
            f"  Map Projection:  {gp.map_projection}",
            f"  Map Zone:        {gp.map_zone}",
        ])

    # ------------------------------------------------------------------
    # Populate form from current ArchiCAD values
    # ------------------------------------------------------------------

    def _populate_from_current(self):
        if not self._current_data:
            return
        d = self._current_data
        pl, sp, gp = d.project_location, d.survey_point, d.geo_ref_params

        self._lon_input.setText(str(pl.longitude))
        self._lat_input.setText(str(pl.latitude))
        self._alt_input.setText(str(pl.altitude))
        self._north_input.setText(str(pl.north_deg))

        self._eastings_input.setText(str(sp.eastings))
        self._northings_input.setText(str(sp.northings))
        self._sp_elev_input.setText(str(sp.elevation))

        self._crs_name_input.setText(gp.crs_name)
        self._desc_input.setText(gp.description)
        self._geodetic_input.setText(gp.geodetic_datum)
        self._vertical_input.setText(gp.vertical_datum)
        self._projection_input.setText(gp.map_projection)
        self._zone_input.setText(gp.map_zone)

        self.statusBar().showMessage("Form populated from current ArchiCAD values")

    # ------------------------------------------------------------------
    # EPSG lookup
    # ------------------------------------------------------------------

    def _lookup_epsg(self):
        code_text = self._epsg_input.text().strip()
        if not code_text:
            QMessageBox.warning(self, "EPSG Code", "Please enter an EPSG code.")
            return
        try:
            code = int(code_text)
        except ValueError:
            QMessageBox.warning(self, "EPSG Code", "EPSG code must be an integer.")
            return

        eastings = self._try_float(self._eastings_input.text())
        northings = self._try_float(self._northings_input.text())

        self._lookup_btn.setEnabled(False)
        self.statusBar().showMessage(f"Looking up EPSG:{code}…")

        w = _LookupWorker(code, eastings, northings)
        w.result.connect(self._on_lookup_success)
        w.failed.connect(self._on_lookup_failed)
        w.finished.connect(lambda: self._lookup_btn.setEnabled(True))
        self._track(w)
        w.start()

    def _on_lookup_success(self, result: dict):
        meta = result["meta"]
        lon = result.get("lon")
        lat = result.get("lat")

        self._crs_name_input.setText(meta.crs_name)
        self._desc_input.setText(meta.description)
        self._geodetic_input.setText(meta.geodetic_datum)
        if meta.vertical_datum:
            self._vertical_input.setText(meta.vertical_datum)
        self._projection_input.setText(meta.map_projection)
        self._zone_input.setText(meta.map_zone)

        if lon is not None and lat is not None:
            self._lon_input.setText(f"{lon:.6f}")
            self._lat_input.setText(f"{lat:.6f}")
            self.statusBar().showMessage(
                f"CRS loaded  |  Lon/Lat computed: {lon:.6f} °, {lat:.6f} °"
            )
        else:
            self.statusBar().showMessage(
                "CRS metadata loaded — enter survey point coordinates to compute Lon/Lat"
            )

    def _on_lookup_failed(self, msg: str):
        self.statusBar().showMessage("EPSG lookup failed")
        QMessageBox.warning(self, "EPSG Lookup Error", f"Could not resolve EPSG code:\n\n{msg}")

    # ------------------------------------------------------------------
    # Build GeorefData from form fields
    # ------------------------------------------------------------------

    def _form_to_data(self) -> GeorefData:
        """Collect form values into a GeorefData.  Raises ValueError on bad input."""

        def need_float(widget: QLineEdit, label: str) -> float:
            text = widget.text().strip()
            if not text:
                raise ValueError(f"'{label}' is empty — please enter a value.")
            try:
                return float(text)
            except ValueError:
                raise ValueError(f"'{label}' is not a valid number: {text!r}")

        pl = ProjectLocation(
            longitude=need_float(self._lon_input, "Longitude"),
            latitude=need_float(self._lat_input, "Latitude"),
            altitude=need_float(self._alt_input, "Altitude"),
            north_deg=need_float(self._north_input, "North"),
        )
        sp = SurveyPointPosition(
            eastings=need_float(self._eastings_input, "Eastings"),
            northings=need_float(self._northings_input, "Northings"),
            elevation=need_float(self._sp_elev_input, "Survey Elevation"),
        )
        gp = GeoReferencingParameters(
            crs_name=self._crs_name_input.text().strip(),
            description=self._desc_input.text().strip(),
            geodetic_datum=self._geodetic_input.text().strip(),
            vertical_datum=self._vertical_input.text().strip(),
            map_projection=self._projection_input.text().strip(),
            map_zone=self._zone_input.text().strip(),
        )
        return GeorefData(project_location=pl, survey_point=sp, geo_ref_params=gp)

    # ------------------------------------------------------------------
    # Preview
    # ------------------------------------------------------------------

    def _preview_changes(self):
        try:
            data = self._form_to_data()
        except ValueError as exc:
            QMessageBox.warning(self, "Validation Error", str(exc))
            return

        payload = build_payload(data)
        self._preview_text.setPlainText(json.dumps(payload, indent=2))
        self._write_btn.setEnabled(bool(self._connection))
        self.statusBar().showMessage(
            "Preview ready — review the payload, then click 'Write to ArchiCAD'"
        )

    # ------------------------------------------------------------------
    # Write
    # ------------------------------------------------------------------

    def _write_to_archicad(self):
        if not self._connection:
            QMessageBox.warning(self, "Not Connected", "No active connection to ArchiCAD.")
            return

        try:
            data = self._form_to_data()
        except ValueError as exc:
            QMessageBox.warning(self, "Validation Error", str(exc))
            return

        reply = QMessageBox.question(
            self,
            "Confirm Write",
            "Write the georeferencing values to ArchiCAD?\n\n"
            "The operation is undoable via ArchiCAD's Undo.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self._write_btn.setEnabled(False)
        self.statusBar().showMessage("Writing to ArchiCAD…")

        w = _WriteWorker(self._connection.tapir, data)
        w.success.connect(self._on_write_success)
        w.failed.connect(self._on_write_failed)
        w.finished.connect(lambda: self._write_btn.setEnabled(bool(self._connection)))
        self._track(w)
        w.start()

    def _on_write_success(self, _result):
        self.statusBar().showMessage("Write successful — refreshing current values…")
        QMessageBox.information(
            self,
            "Success",
            "Georeferencing data written to ArchiCAD successfully.\n\n"
            "Use ArchiCAD's Undo if you need to revert.",
        )
        self._read_from_archicad()   # refresh the current-values panel

    def _on_write_failed(self, msg: str):
        self.statusBar().showMessage("Write failed")
        QMessageBox.critical(self, "Write Error", f"Failed to write to ArchiCAD:\n\n{msg}")
