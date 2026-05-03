import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DATA_DIR = os.path.join(BASE_DIR, "data")
LOG_DIR = os.path.join(BASE_DIR, "logs")

DATABASE_PATH = os.path.join(DATA_DIR, "records.db")

DEFAULT_FONT_NAME = "宋体"
DEFAULT_FONT_SIZE = 12
DEFAULT_LINE_SPACING = 1.5
DEFAULT_INDENTATION = 2.0

APP_NAME = "Word自动化工具"
APP_VERSION = "1.0.0"
AUTHOR = "AutoWord Team"

for directory in [TEMPLATE_DIR, OUTPUT_DIR, DATA_DIR, LOG_DIR]:
    if not os.path.exists(directory):
        os.makedirs(directory)
