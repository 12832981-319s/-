import re
from typing import Iterable

import pandas as pd


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    return re.sub(r"\s+", " ", text)


def clean_value(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = normalize_text(value)
    if text in {"", "-", "--", "...", "…"}:
        return 0.0

    match = re.search(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    return float(match.group()) if match else 0.0


def collapse_multiline_header(parts: Iterable) -> str:
    cleaned_parts = [normalize_text(part) for part in parts if normalize_text(part)]
    return " ".join(cleaned_parts)


def finalize_dataframe(df: pd.DataFrame, text_columns: int = 1) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned = cleaned.dropna(how="all")
    cleaned.columns = [normalize_text(col) for col in cleaned.columns]
    cleaned = cleaned.loc[:, [col for col in cleaned.columns if col]]

    for index, column in enumerate(cleaned.columns):
        if index < text_columns:
            cleaned[column] = cleaned[column].map(normalize_text)
        else:
            cleaned[column] = cleaned[column].map(clean_value)

    first_column = cleaned.columns[0]
    cleaned = cleaned[cleaned[first_column] != ""]
    return cleaned.reset_index(drop=True)
