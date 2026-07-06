import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    import pyxlsb
    print("pyxlsb is installed!")
except ImportError:
    print("pyxlsb is NOT installed.")
