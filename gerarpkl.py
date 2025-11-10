# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import pandas as pd

# ENTRADA (seu XLSX)
INPUT_XLSX = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\xlsx\Base_Atualizada.xlsx")

# SAÍDA (sua base PKL padrão)
OUTPUT_PKL = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\pkl\base.pkl")
OUTPUT_PKL.parent.mkdir(parents=True, exist_ok=True)

def main():
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {INPUT_XLSX}")

    # Leia o Excel preservando telefones como string (ajuste os nomes de colunas conforme sua planilha)
    dtype_map = {
        "Telefone": "string",
        "SPX TN": "string",
    }
    df = pd.read_excel(INPUT_XLSX, dtype=dtype_map)

    # (Opcional) normalize NaN para strings vazias em colunas-chave
    for col in ["Nome", "Telefone", "Local", "SPX TN"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype("string")

    # Salva como PKL (DataFrame completo)
    # Obs.: PKL é dependente de versão. Ideal abrir com o mesmo Python/NumPy/Pandas.
    df.to_pickle(OUTPUT_PKL)

    print(f"OK! Salvo PKL: {OUTPUT_PKL}")
    print(f"Linhas: {len(df)} | Colunas: {list(df.columns)}")

if __name__ == "__main__":
    main()