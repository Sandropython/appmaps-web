# conversores_patch.py
from __future__ import annotations
from pathlib import Path
import os, pickle


def _load_last_dir(app) -> Path:
    cfg = getattr(app, "CONFIG_FILE", "last_dir.txt")
    try:
        p = Path(cfg)
        if p.exists():
            txt = p.read_text(encoding="utf-8").strip()
            if txt:
                return Path(txt)
    except Exception:
        pass
    return Path(__file__).resolve().parent

def _save_last_dir(app, directory: Path) -> None:
    cfg = getattr(app, "CONFIG_FILE", "last_dir.txt")
    try:
        Path(cfg).write_text(str(directory), encoding="utf-8")
    except Exception:
        pass

def _ensure_filemanager(app):
    if hasattr(app, "file_manager") and app.file_manager:
        return app.file_manager
    from kivymd.uix.filemanager import MDFileManager
    app.file_manager = MDFileManager(
        exit_manager=lambda *a, **k: app.file_manager.close(),
        select_path=lambda *a, **k: None,
        preview=False,
    )
    # extensões aceitas no diálogo (visual)
    try:
        app.file_manager.ext = [".xlsx", ".pkl"]
    except Exception:
        pass
    return app.file_manager

def _converter_xlsx_para_pkl(in_path: Path) -> Path:
    import pandas as pd
    in_path = Path(in_path)
    root = Path(__file__).resolve().parent
    out_dir = root / "bases" / "pkl"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{in_path.stem}.pkl"

    dtype_map = {"Telefone": "string", "SPX TN": "string"}
    try:
        df = pd.read_excel(in_path, dtype=dtype_map)
    except Exception:
        df = pd.read_excel(in_path)

    # normalização leve
    for col in ["Nome", "Telefone", "Local", "SPX TN"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype("string")

    df.to_pickle(out_path)
    return out_path

class _NumpyCoreRedirectUnpickler(pickle.Unpickler):
    def find_class(self, module, name):
        if module.startswith("numpy.core."):
            module = module.replace("numpy.core.", "numpy.")
        return super().find_class(module, name)

def _read_pickle_compat(path: Path):
    import pandas as pd
    try:
        with open(path, "rb") as f:
            return _NumpyCoreRedirectUnpickler(f).load()
    except Exception:
        return pd.read_pickle(path)

def _object_to_dataframe(obj):
    import pandas as pd
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    if isinstance(obj, dict) and obj and all(isinstance(v, dict) for v in obj.values()):
        df = pd.DataFrame.from_dict(obj, orient="index").reset_index()
        df = df.rename(columns={"index": "Endereco"})
        return df
    if isinstance(obj, list) and obj and all(isinstance(v, dict) for v in obj):
        return pd.DataFrame(obj)
    return pd.DataFrame(obj)

def _converter_pkl_para_xlsx(in_path: Path) -> Path:
    import pandas as pd
    in_path = Path(in_path)
    root = Path(__file__).resolve().parent
    out_dir = root / "bases" / "xlsx"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{in_path.stem}.xlsx"

    obj = _read_pickle_compat(in_path)
    df = _object_to_dataframe(obj)

    for col in ("Telefone", "SPX TN"):
        if col in df.columns:
            df[col] = df[col].astype("string")
    for col in ("Nome", "Telefone", "Local", "SPX TN"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype("string")

    df.to_excel(out_path, index=False, engine="openpyxl")
    return out_path

def abrir_xlsx_para_pkl(self):
    self._convert_mode = "xlsx2pkl"
    fm = _ensure_filemanager(self)
    fm.select_path = lambda path: _on_select_convert(self, Path(path))
    start = _load_last_dir(self)
    fm.show(str(start))

def abrir_pkl_para_xlsx(self):
    self._convert_mode = "pkl2xlsx"
    fm = _ensure_filemanager(self)
    fm.select_path = lambda path: _on_select_convert(self, Path(path))
    start = _load_last_dir(self)
    fm.show(str(start))

def _on_select_convert(self, path: Path):
    try:
        if getattr(self, "_convert_mode", None) == "xlsx2pkl":
            if path.suffix.lower() != ".xlsx":
                return self.show_dialog("Arquivo inválido", "Escolha um arquivo .xlsx")
            out = _converter_xlsx_para_pkl(path)
        elif getattr(self, "_convert_mode", None) == "pkl2xlsx":
            if path.suffix.lower() != ".pkl":
                return self.show_dialog("Arquivo inválido", "Escolha um arquivo .pkl")
            out = _converter_pkl_para_xlsx(path)
        else:
            return self.show_dialog("Erro", "Modo de conversão desconhecido.")

        _save_last_dir(self, path.parent)

        if hasattr(self, "file_manager") and self.file_manager:
            self.file_manager.close()

        self.show_dialog("Pronto", f"Convertido com sucesso!\n\nEntrada: {path}\nSaída:   {out}")
    except Exception as e:
        self.show_dialog("Erro", f"Falha na conversão:\n{e}")

def apply(AppMapsClass):
    if not hasattr(AppMapsClass, "abrir_xlsx_para_pkl"):
        AppMapsClass.abrir_xlsx_para_pkl = abrir_xlsx_para_pkl
    if not hasattr(AppMapsClass, "abrir_pkl_para_xlsx"):
        AppMapsClass.abrir_pkl_para_xlsx = abrir_pkl_para_xlsx
    return AppMapsClass