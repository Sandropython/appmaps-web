# -*- coding: utf-8 -*-
"""
gerarxlsx_fixed.py
------------------
Converte um PKL (dict de dicts, DataFrame, list[dict], etc.) para XLSX no formato "uma linha por registro".
- Corrige o caso típico: {endereco: {Bairro, Cep, Nome, Telefone}} -> DataFrame com coluna 'Endereco'.
- Preserva colunas como texto (ex.: Telefone) para não perder zeros à esquerda.
- Permite usar via CLI: python gerarxlsx_fixed.py INPUT_PKL OUTPUT_XLSX
"""

from __future__ import annotations
from pathlib import Path
import pickle
import pandas as pd
import sys

# ===================== AJUSTE PADRÃO (pode alterar aqui) =====================
# Caminhos padrão (serão ignorados se você passar via CLI)
INPUT_PKL = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\pkl\base_cel_atualizada.pkl")
OUTPUT_XLSX = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\xlsx\base_cel.xlsx")
# =============================================================================

# --- Leitor compatível para PKL antigos ou com mudanças do NumPy ---
class _NumpyCoreRedirectUnpickler(pickle.Unpickler):
    def find_class(self, module, name):
        # Redireciona módulos renomeados do NumPy 2.x, caso o PKL tenha sido feito em outra versão
        if module.startswith("numpy.core."):
            module = module.replace("numpy.core.", "numpy.")
        return super().find_class(module, name)

def _read_pickle_compat(path: Path):
    """Tenta ler PKL com compatibilidade extra (NumPy 2.x, etc.)."""
    try:
        with open(path, "rb") as f:
            return _NumpyCoreRedirectUnpickler(f).load()
    except Exception:
        # fallback: pandas.read_pickle (pode funcionar em alguns casos)
        return pd.read_pickle(path)

def to_tidy_dataframe(obj) -> pd.DataFrame:
    """Converte diferentes estruturas em um DataFrame 'tidy' (linhas = registros)."""
    # Caso já seja DataFrame
    if isinstance(obj, pd.DataFrame):
        return obj.copy()

    # Caso seja dict de dicts: {chave_externa: {campo1:..., campo2:...}}
    if isinstance(obj, dict) and obj and all(isinstance(v, dict) for v in obj.values()):
        df = pd.DataFrame.from_dict(obj, orient="index").reset_index()
        # nomeia a coluna da chave externa
        df = df.rename(columns={"index": "Endereco"})
        return df

    # Caso seja uma lista de dicts
    if isinstance(obj, list) and obj and all(isinstance(v, dict) for v in obj):
        return pd.DataFrame(obj)

    # Caso genérico: tenta converter
    try:
        return pd.DataFrame(obj)
    except Exception as e:
        raise TypeError(f"Objeto no PKL não é tabular; não foi possível converter para DataFrame: {e}")

def ajustar_tipos_e_nulos(df: pd.DataFrame) -> pd.DataFrame:
    """Ajusta colunas conhecidas para string e preenche nulos em algumas colunas."""
    df = df.copy()
    # Colunas que devem ser string para preservar zeros à esquerda
    for col in ("Telefone", "SPX TN"):
        if col in df.columns:
            df[col] = df[col].astype("string")

    # Preenche nulos com vazio em colunas de texto comuns
    for col in ("Nome", "Telefone", "Local", "SPX TN"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype("string")
    return df

def main(argv=None):
    argv = argv or sys.argv[1:]
    # Permite override via CLI
    in_pkl = Path(argv[0]) if len(argv) >= 1 else INPUT_PKL
    out_xlsx = Path(argv[1]) if len(argv) >= 2 else OUTPUT_XLSX

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    obj = _read_pickle_compat(in_pkl)
    df = to_tidy_dataframe(obj)
    df = ajustar_tipos_e_nulos(df)

    # Salva em XLSX (requer openpyxl)
    df.to_excel(out_xlsx, index=False, engine="openpyxl")

    print(f"OK! Salvo Excel: {out_xlsx}")
    print(f"Linhas: {len(df)} | Colunas: {list(df.columns)}")

if __name__ == "__main__":
    main()
