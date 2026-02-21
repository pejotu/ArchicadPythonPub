"""Entry point for the ArchiCAD Georeferencing Tool.

Run this script from within ArchiCAD's Python environment:
    python apps/georef_tool/main.py

The script requires:
  - ArchiCAD running with the Python API enabled
  - The Tapir addon installed (https://github.com/ENZYME-APD/tapir-archicad-automation)
  - PyQt5, pyproj, and requests installed in the Python environment
"""

import os
import sys

# Make all sibling modules importable by their bare name (same pattern as classreader)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from ui import GeorefUI


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = GeorefUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
