# -*- coding: utf-8 -*-
import os
import shutil

import PyInstaller.__main__


ROOT = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(ROOT, "dist", "XYD_PDF_Tools_Pro")
PORTABLE_DIR = os.path.join(
    ROOT, "Output", "V4.0", "Portable", "XYD_PDF_Tools_Pro_V4.0"
)


if __name__ == "__main__":
    PyInstaller.__main__.run([
        "--noconfirm",
        "--clean",
        os.path.join(ROOT, "XYD_PDF_Tools_Pro.spec"),
    ])
    if os.path.exists(PORTABLE_DIR):
        shutil.rmtree(PORTABLE_DIR)
    os.makedirs(os.path.dirname(PORTABLE_DIR), exist_ok=True)
    shutil.copytree(DIST_DIR, PORTABLE_DIR)
    print(f"V4.0 portable build ready: {PORTABLE_DIR}")
