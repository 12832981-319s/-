import sys
from pathlib import Path


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def sample_data_dir() -> Path:
    return app_base_dir().parent / "data" if not getattr(sys, "frozen", False) else app_base_dir() / "data"


def config_file_path(filename: str) -> Path:
    return app_base_dir() / filename


def resource_path(*parts: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS).resolve().joinpath(*parts)
    return app_base_dir().joinpath(*parts)
