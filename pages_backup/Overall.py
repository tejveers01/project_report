import os
import sys
import runpy

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP_DIR = os.path.join(ROOT_DIR, "Overall")

sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, APP_DIR)

runpy.run_path(os.path.join(APP_DIR, "ol_main.py"), run_name="__main__")