import pandas as pd

def load_table(path: str, sheet: str = None):
    if path.endswith(".csv"):
        return pd.read_csv(path)
    elif path.endswith(".xlsx") or path.endswith(".xls"):
        return pd.read_excel(path, sheet_name=sheet or 0)
    else:
        raise ValueError(f"Неподдерживаемый формат таблицы: {path}")
