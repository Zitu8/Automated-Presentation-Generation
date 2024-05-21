# Imports
import collections
import collections.abc
from collections import defaultdict
import logging
import json
import sys
import random
import os
import math
import re
import time
import datetime
from pptx import Presentation
from PIL import Image
from rectpack import newPacker
from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt, Emu
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData

# Add the Inkscape directory to the PATH, so that Cairosvg can find it. 
# This might need to be changed depending on the installation directory of Inkscape
os.environ["PATH"] += os.pathsep + 'C:\\Program Files\\Inkscape'
# Now import Cairosvg
import cairosvg


""" Constants """""
# Paths to the files
USER_NAME = os.getlogin()
CURRENT_TIME = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
MEDIA_PATH = "logos" # Change this to the path where the logos are stored

""" Change these values for a different map configuration """""
# Paths to the files, change these to the correct paths on your system
TECH_DATA_PATH = "technology_data.json"
SUPPLIERS_DATA_PATH = "suppliers_data.json"
FOLDERS = {"Logs", "Map Output"}
TEMPLATE_PATH = "Map Template.pptx"

# Base Colors in hex
AM_GREEN = "73BA5A"
AM_BLUE = "1E3347"
AM_GREY = "878787"
AM_SKY = "6AA7EF"
COLOR_THEME = {
    "Liquid": AM_GREY,
    "Thermoset": AM_GREY,
    "Elastomer": AM_GREEN,
    "Clay / Concrete": AM_SKY,
    "Sand": AM_GREY,
    "Technical Ceramics": AM_GREEN,
    "Direct": AM_BLUE,
    "Sinter-based": AM_GREEN,
    "Other": AM_GREY,
    "Metal": AM_BLUE,
    "Polymer": AM_GREEN,
    "Ceramics": AM_SKY
}

# Predefined feedstock orders into one dictionary
feedstock_order = {
    'Metal': ['Powder', 'Wire', 'Rods', 'Other', 'Dispersion', 'Filament', 'Pellets'],
    'Ceramics': ['Dispersion', 'Powder', 'Pellets', 'Filament'],
    'Polymer': ['Liquid', 'Sheet', 'Powder', 'Tape', 'Filament', 'Pellets']
}

# Predefined order for Ceramics MAP_Feature items
FEATURE_ORDER = ["Thermoset", "Elastomer", "Thermoplastic", "Direct", "Sinter-based", "Technical Ceramics", "Clay / Concrete", "Sand"]

# Define Spacings
MARGIN_CM = 3
PADDING_CM = 0.5
LOGO_OFFSET_CM = 5
STEP = 0.01
ANGLE_OFFSET = 20
LEGEND_FONT_SIZE = Pt(18)
LEGEND_FONT = 'Roboto Light'
CIRCLE_DIAMETER = Cm(0.61)  # Circle diameter for the legend


# Define Boundaries
LOGO_MAX_AREA = 2
FEEDSTOCK_LOGO_WIDTH = Cm(2.2)
MAX_CHARS = 14
DIV_LINE_WIDTH_PX = 1
FOOTER_HEIGHT_CM = 3
MIN_WIDTH = 7
MAX_WIDTH = 20
MIN_HEIGHT = 1.5
DPI = 96