"""Path helpers for script and PyInstaller runtime modes."""

import os
import sys


def app_dir():
    """Directory beside the executable or script, used for config and output."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bundled_dir():
    """Directory for bundled resources in PyInstaller or source mode."""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    return os.path.join(app_dir(), "config.json")


def default_config_path():
    return os.path.join(bundled_dir(), "config.json")


def assets_dir():
    p = os.path.join(app_dir(), "assets")
    if os.path.isdir(p):
        return p
    return os.path.join(bundled_dir(), "assets")


def make_normalized_output_path(input_path):
    p = os.fspath(input_path)
    folder = os.path.dirname(p)
    stem, _ = os.path.splitext(os.path.basename(p))
    return os.path.join(folder, stem + "_规范化.docx")
