import sys
from src.config import config

# uno.py directory (package) path
# ***** UNO *****
sys.path.append(config['libreoffice']['python_uno_location'])
import uno