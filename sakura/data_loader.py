import pandas as pd

def load_excel(path: str) -> pd.DataFrame:
    return pd.read_excel(path)

import pandas as pd

import pandas as pd

def load_csv(path: str, sep: str = ";"):
    encodings = ["utf-8", "cp1251", "latin1"]
    last_error = None

    for enc in encodings:
        try:
            df = pd.read_csv(
                path,
                sep=sep,
                encoding=enc,
                engine="python",
                on_bad_lines="skip",
                dtype=str,       # читаем всё как строки
            )

            # заменяем запятую на точку во всех строковых столбцах
            df = df.apply(
                lambda col: col.str.replace(",", ".", regex=False)
                if col.dtype == "object" else col
            )

            # даём твоему коду дальше самому приводить к float/int
            return df
        except Exception as e:
            last_error = e
            continue

    raise last_error



def load_file(path: str) -> pd.DataFrame:
    path_lower = path.lower()
    if path_lower.endswith(".xlsx") or path_lower.endswith(".xls"):
        return load_excel(path)
    elif path_lower.endswith(".csv"):
        return load_csv(path)
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {path}")
