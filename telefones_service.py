# -*- coding: utf-8 -*-
"""
Serviço para atualizar NOME/TELEFONE da planilha de entrada usando uma base PKL
(telefones) no formato:
{
  "<Local>": {"Nome": "...", "Telefone": "..."},
  "<Local>__1": {"Nome": "...", "Telefone": "..."},
  ...
}

Regras principais:
- Preencher Nome/Telefone da planilha **apenas se estiverem em branco**.
- Mesmo que Nome/Telefone já existam na planilha, verificar se o Local existe na base:
  - Se existir e a combinação (Nome, Telefone) da linha for diferente de todas da base,
    acrescentar variação com sufixo __N na base PKL.
  - Se não existir o Local na base, inserir com “tudo de direito” (par principal
    ou variação, conforme necessário).
- Consolidar “Nomes adicionais” e “Telefones adicionais” (deduplicados) com os demais
  registros do mesmo Local (base PKL), mantendo o primeiro como principal.
- Gerar uma planilha “de busca” **somente** das linhas que possuem adicionais, com:
  Sequence, SPX TN, Nome, Telefone, Local, Nomes adicionais, Telefones adicionais.

Obs.: Por padrão, a chave é apenas `Local` (igual ao seu Jupyter). Se quiser usar
Local+Complemento, chame com `usar_complemento=True`.
"""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple, List, Any
import pandas as pd
import re
import pickle


# ----------------------------- Utilitários ----------------------------- #

def _s(text: Any) -> str:
    if pd.isna(text):
        return ""
    return str(text).strip()


_phone_digits = re.compile(r"\D+")


def limpar_telefone(valor: Any) -> str:
    s = _s(valor)
    if s.endswith(".0"):
        s = s[:-2]
    return _phone_digits.sub("", s)


def _dedupe_keep_order(seq: List[str]) -> List[str]:
    seen = set()
    out = []
    for x in seq:
        if not x:
            continue
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def _merge_adicionais(existente: str, novos: List[str]) -> List[str]:
    base = []
    if existente:
        # aceita "," ou ";" como separadores
        base = [p.strip() for p in existente.replace(";", ",").split(",") if p.strip()]
    return _dedupe_keep_order(base + [p for p in novos if p])


# ---------------------- Leitura/Escrita da base PKL -------------------- #

def carregar_base_pkl(caminho: Path) -> Dict[str, Dict[str, str]]:
    caminho = Path(caminho)
    if not caminho.exists():
        return {}
    with open(caminho, "rb") as f:
        data = pickle.load(f)
    # normaliza estrutura
    out: Dict[str, Dict[str, str]] = {}
    for k, v in (data or {}).items():
        if isinstance(v, dict):
            out[str(k)] = {
                "Nome": _s(v.get("Nome")),
                "Telefone": limpar_telefone(v.get("Telefone")),
            }
    return out


def salvar_base_pkl(base: Dict[str, Dict[str, str]], caminho: Path) -> None:
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho, "wb") as f:
        pickle.dump(base, f)


# -------------------------- Operações de base -------------------------- #

def _prefixo_chaves(base: Dict[str, Dict[str, str]], chave: str) -> List[str]:
    """Retorna todas as chaves que são =chave ou começam com chave + "__"."""
    return [k for k in base.keys() if k == chave or k.startswith(chave + "__")]


def _pares_do_prefixo(base: Dict[str, Dict[str, str]], chave: str) -> List[Tuple[str, str, str]]:
    """Lista (chave_completa, Nome, Telefone) de um Local e suas variações."""
    saida = []
    for k in _prefixo_chaves(base, chave):
        d = base.get(k, {})
        saida.append((k, _s(d.get("Nome")), limpar_telefone(d.get("Telefone"))))
    return saida


def _proximo_indice_variacao(base: Dict[str, Dict[str, str]], chave: str) -> int:
    indices = {0}
    for k in _prefixo_chaves(base, chave):
        if k != chave and "__" in k:
            try:
                indices.add(int(k.split("__", 1)[1]))
            except Exception:
                pass
    n = 1
    while n in indices:
        n += 1
    return n


def garantir_par_na_base(base: Dict[str, Dict[str, str]], chave: str, nome: str, tel: str) -> str:
    """Garante que (nome, tel) está registrado na base para o prefixo `chave`.
    Retorna a chave (com ou sem sufixo) em que foi gravado.
    """
    nome, tel = _s(nome), limpar_telefone(tel)
    existentes = _pares_do_prefixo(base, chave)
    for k, n, t in existentes:
        if n == nome and t == tel:
            return k  # já existe
    # inserir
    if not any(k == chave for k, *_ in existentes):
        kfinal = chave
    else:
        kfinal = f"{chave}__{_proximo_indice_variacao(base, chave)}"
    base[kfinal] = {"Nome": nome, "Telefone": tel}
    return kfinal


# ---------------------------- Núcleo do processo ---------------------------- #

@dataclass
class Resumo:
    linhas: int = 0
    nomes_preenchidos: int = 0
    telefones_preenchidos: int = 0
    nomes_adicionados: int = 0
    telefones_adicionados: int = 0
    nao_encontrados: int = 0

    def to_dict(self) -> Dict[str, int]:
        return dict(self.__dict__)


def _garantir_colunas(df: pd.DataFrame) -> pd.DataFrame:
    cols_need = [
        "AT ID", "Sequence", "Stop", "SPX TN", "Destination Address", "Bairro", "City",
        "Zipcode/Postal code", "Latitude", "Longitude", "Nome", "Telefone", "Local",
        "Complemento", "Nomes adicionais", "Telefones adicionais",
    ]
    for c in cols_need:
        if c not in df.columns:
            df[c] = ""
    return df


def atualizar_telefones(
    df_entrada: pd.DataFrame,
    base_pkl: Dict[str, Dict[str, str]],
    usar_complemento: bool = False,
) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]], Resumo, pd.DataFrame]:
    """
    Executa a regra de atualização sobre um DataFrame e uma base de telefones (PKL já carregada).
    Retorna: (df_atualizado, base_atualizada, resumo, df_busca)
    """
    df = df_entrada.copy()
    df.columns = df.columns.str.strip()
    df = _garantir_colunas(df)

    # normalizações
    df["Nome"] = df["Nome"].fillna("").astype(str).str.strip()
    df["Telefone"] = df["Telefone"].apply(limpar_telefone)
    for col in ("Nomes adicionais", "Telefones adicionais", "Complemento", "Local"):
        df[col] = df[col].fillna("").astype(str).str.strip()

    resumo = Resumo(linhas=len(df))
    busca_rows: List[Dict[str, Any]] = []

    for idx, row in df.iterrows():
        local = _s(row.get("Local"))
        comp = _s(row.get("Complemento"))
        chave = f"{local} {comp}".strip() if usar_complemento else local

        nome_linha = _s(row.get("Nome"))
        tel_linha = limpar_telefone(row.get("Telefone"))

        pares = _pares_do_prefixo(base_pkl, chave)
        nomes_base = [n for _, n, _ in pares if n]
        tels_base = [t for _, _, t in pares if t]

        # Se não houver nada na base e a linha já tiver Nome/Tel, inserir
        if not pares and (nome_linha or tel_linha) and local:
            garantir_par_na_base(base_pkl, chave, nome_linha, tel_linha)
            pares = _pares_do_prefixo(base_pkl, chave)
            nomes_base = [n for _, n, _ in pares if n]
            tels_base = [t for _, _, t in pares if t]

        # Preencher NaN/blank com o primeiro da base
        if nomes_base and not nome_linha:
            df.at[idx, "Nome"] = nomes_base[0]
            resumo.nomes_preenchidos += 1
            nome_linha = nomes_base[0]
        if tels_base and not tel_linha:
            df.at[idx, "Telefone"] = tels_base[0]
            resumo.telefones_preenchidos += 1
            tel_linha = tels_base[0]

        # Se a linha tem Nome/Tel e eles NÃO estão na base, registrar como variação
        if local and (nome_linha or tel_linha):
            if not any((n == nome_linha and t == tel_linha) for _, n, t in pares):
                garantir_par_na_base(base_pkl, chave, nome_linha, tel_linha)
                pares = _pares_do_prefixo(base_pkl, chave)
                nomes_base = [n for _, n, _ in pares if n]
                tels_base = [t for _, _, t in pares if t]

        # Adicionais (todos da base exceto o primeiro)
        nomes_adic = _dedupe_keep_order([n for n in nomes_base[1:]])
        tels_adic = _dedupe_keep_order([t for t in tels_base[1:]])

        # Mescla/dedup com adicionais já existentes na planilha
        nomes_adic = _merge_adicionais(_s(row.get("Nomes adicionais")), nomes_adic)
        tels_adic = _merge_adicionais(_s(row.get("Telefones adicionais")), tels_adic)

        if nomes_adic:
            df.at[idx, "Nomes adicionais"] = "; ".join(nomes_adic)
            resumo.nomes_adicionados += len(nomes_adic)
        if tels_adic:
            df.at[idx, "Telefones adicionais"] = "; ".join(tels_adic)
            resumo.telefones_adicionados += len(tels_adic)

        if not pares:
            resumo.nao_encontrados += 1

        # Coleta para planilha de busca (somente se houver adicionais)
        if nomes_adic or tels_adic:
            busca_rows.append({
                "Sequence": row.get("Sequence"),
                "SPX TN": row.get("SPX TN"),
                "Nome": df.at[idx, "Nome"],
                "Telefone": df.at[idx, "Telefone"],
                "Local": local,
                "Nomes adicionais": df.at[idx, "Nomes adicionais"],
                "Telefones adicionais": df.at[idx, "Telefones adicionais"],
            })

    df_busca = pd.DataFrame(busca_rows)
    return df, base_pkl, resumo, df_busca


# -------------------------- Pipeline de caminhos -------------------------- #

def executar_processo(
    caminho_excel_entrada: Path,
    caminho_pkl_base: Path,
    saida_excel_atualizada: Path | None = None,
    saida_pkl_atualizado: Path | None = None,
    saida_excel_busca: Path | None = None,
    usar_complemento: bool = False,
) -> Dict[str, Any]:
    """Executa o processo a partir de caminhos e salva os resultados.
    Retorna dicionário com caminhos de saída e resumo.
    """
    caminho_excel_entrada = Path(caminho_excel_entrada)
    caminho_pkl_base = Path(caminho_pkl_base)

    if saida_excel_atualizada is None:
        saida_excel_atualizada = caminho_excel_entrada.with_name(
            caminho_excel_entrada.stem + "_Atualizada.xlsx"
        )
    if saida_pkl_atualizado is None:
        saida_pkl_atualizado = caminho_pkl_base.with_name(
            caminho_pkl_base.stem + "_atualizada.pkl"
        )
    if saida_excel_busca is None:
        saida_excel_busca = caminho_excel_entrada.with_name(
            caminho_excel_entrada.stem + "_Busca.xlsx"
        )

    df_in = pd.read_excel(caminho_excel_entrada, dtype={"Telefone": str})
    base = carregar_base_pkl(caminho_pkl_base)

    df_out, base_out, resumo, df_busca = atualizar_telefones(
        df_in, base, usar_complemento=usar_complemento
    )

    # salvar
    saida_excel_atualizada = Path(saida_excel_atualizada)
    saida_excel_atualizada.parent.mkdir(parents=True, exist_ok=True)
    df_out.to_excel(saida_excel_atualizada, index=False)

    salvar_base_pkl(base_out, saida_pkl_atualizado)

    busca_path = ""
    if not df_busca.empty:
        df_busca.to_excel(saida_excel_busca, index=False)
        busca_path = str(saida_excel_busca)

    return {
        "excel_atualizada": str(saida_excel_atualizada),
        "pkl_atualizado": str(saida_pkl_atualizado),
        "excel_busca": busca_path,
        "resumo": resumo.to_dict(),
    }
