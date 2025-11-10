# -*- coding: utf-8 -*-
"""
Mapeia quais arquivos .py e .kv do projeto parecem USADOS ou ORFÃO?.
Critérios:
- .py USADO: alcançado via import local a partir do main.py (grafo de imports).
- .kv USADO: referenciado por Builder.load_file("...kv"), implícito do MDApp
  (NomeApp -> nome.kv), ou incluído via #:include em outro .kv usado.

Uso:
    python mapear_usados_orfaos.py

Saídas:
    - arquivos_usados_orfaos.csv  (sempre)
    - arquivos_usados_orfaos.xlsx (se openpyxl estiver instalado)
"""

import os
import re
import sys
from pathlib import Path

try:
    import pandas as pd
except Exception as e:
    print("Instale pandas: pip install pandas")
    raise

# ---------------- Config ----------------
# Raiz do projeto (pasta atual)
ROOT = Path(__file__).resolve().parent

# Ignore estas pastas ao varrer:
IGNORE_DIRS = {
    ".git", ".idea", ".vscode", "__pycache__", "build", "dist",
    "venv", ".venv", "env", ".env", "Scripts", "Lib", "site-packages",
}

# Nome preferencial do entry point:
MAIN_NAME = "main.py"
# ----------------------------------------


def iter_files(suffixes):
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        # p.relative_to(ROOT).parts é uma tupla de pastas
        if any(part in IGNORE_DIRS for part in p.relative_to(ROOT).parts):
            continue
        if p.suffix.lower() in suffixes:
            yield p


def rtext(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""


def choose_entry_point(py_files):
    """Escolhe o entry point: main.py ou, na falta, o maior .py do projeto."""
    mains = [p for p in py_files if p.name == MAIN_NAME]
    if mains:
        return mains[0]
    if py_files:
        return max(py_files, key=lambda p: p.stat().st_size)
    return None


def build_project_map(py_files):
    """Mapeia 'pkg/sub/mod.py' -> Path."""
    return {str(p.relative_to(ROOT)).replace("\\", "/"): p for p in py_files}


def candidates_for_import(modname: str):
    """Transforma pkg.sub.mod -> [pkg/sub/mod.py, pkg/sub/__init__.py]"""
    parts = modname.split(".")
    joined = "/".join(parts)
    return [f"{joined}.py", f"{joined}/__init__.py"]


IMPORT_RE = re.compile(
    r'^\s*(?:from\s+([A-Za-z0-9_\.]+)\s+import\s+([A-Za-z0-9_\*,\s]+)|import\s+([A-Za-z0-9_\.]+))',
    re.MULTILINE,
)
BUILDER_RE = re.compile(r'Builder\.load_file\(\s*[\'"]([^\'"]+\.kv)[\'"]\s*\)')
MDAPP_RE = re.compile(r'class\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(\s*(?:MDApp|App)\s*\)\s*:', re.MULTILINE)
INCLUDE_RE = re.compile(r'#:include\s+([^\s]+)')


def find_local_imports(py_path: Path, project_map: dict) -> set:
    """Retorna módulos .py locais importados desse arquivo."""
    txt = rtext(py_path)
    locals_found = set()
    for m in IMPORT_RE.finditer(txt):
        mod = m.group(1) or m.group(3)
        if not mod:
            continue
        for cand in candidates_for_import(mod):
            if cand in project_map:
                locals_found.add(project_map[cand])
    return locals_found


def traverse_import_graph(entry: Path, py_files: list) -> set:
    """Alcança arquivos .py a partir do entry via imports locais."""
    used = set()
    if not entry:
        return used
    project_map = build_project_map(py_files)
    stack = [entry]
    used.add(entry)
    while stack:
        cur = stack.pop()
        for dep in find_local_imports(cur, project_map):
            if dep not in used:
                used.add(dep)
                stack.append(dep)
    return used


def find_used_kv(py_files, kv_files) -> set:
    """Retorna conjunto de .kv usados por Builder.load_file, implícito do MDApp, e include#."""
    kv_set = set()

    # Explicitos via Builder.load_file()
    for p in py_files:
        txt = rtext(p)
        for m in BUILDER_RE.finditer(txt):
            kvname = m.group(1)
            rel = (p.parent / kvname)
            if rel.exists():
                kv_set.add(rel.resolve())
            else:
                # matching por basename
                for k in kv_files:
                    if k.name == Path(kvname).name:
                        kv_set.add(k.resolve())

    # Implícito do MDApp
    for p in py_files:
        txt = rtext(p)
        for m in MDAPP_RE.finditer(txt):
            app_class = m.group(1)
            base = app_class[:-3] if app_class.lower().endswith("app") else app_class
            implicit = base.lower() + ".kv"
            for k in kv_files:
                if k.name == implicit:
                    kv_set.add(k.resolve())

    # Expandir include chain (#:include ...)
    def expand_includes(seed_set: set) -> set:
        expanded = set(seed_set)
        changed = True
        while changed:
            changed = False
            new_add = set()
            for kv in list(expanded):
                txt = rtext(kv)
                for m in INCLUDE_RE.finditer(txt):
                    inc = m.group(1).strip().strip("'\"")
                    cand = (kv.parent / inc)
                    if cand.exists() and cand.suffix.lower() == ".kv":
                        new_add.add(cand.resolve())
                    else:
                        # match por basename em qualquer lugar
                        for kk in kv_files:
                            if kk.name == Path(inc).name:
                                new_add.add(kk.resolve())
            add_now = new_add - expanded
            if add_now:
                expanded |= add_now
                changed = True
        return expanded

    return expand_includes(kv_set)


def main():
    all_py = list(iter_files({".py"}))
    all_kv = list(iter_files({".kv"}))

    entry = choose_entry_point(all_py)
    used_py = traverse_import_graph(entry, all_py)
    used_kv = find_used_kv(all_py, all_kv)

    rows = []
    for p in all_py:
        rows.append({
            "arquivo": str(p.relative_to(ROOT)),
            "tipo": "py",
            "status": "USADO" if p in used_py else "ORFÃO?",
            "motivo": "alcance via imports a partir do main.py" if p in used_py else "não alcançado a partir do main.py",
        })
    for k in all_kv:
        rows.append({
            "arquivo": str(k.relative_to(ROOT)),
            "tipo": "kv",
            "status": "USADO" if k.resolve() in used_kv else "ORFÃO?",
            "motivo": "referenciado por Builder/implícito/include" if k.resolve() in used_kv else "não referenciado",
        })

    df = pd.DataFrame(rows).sort_values(["tipo", "status", "arquivo"])

    out_csv = ROOT / "arquivos_usados_orfaos.csv"
    df.to_csv(out_csv, index=False, encoding="utf-8")
    print(f"[OK] CSV salvo em: {out_csv}")

    # Excel opcional
    try:
        with pd.ExcelWriter(ROOT / "arquivos_usados_orfaos.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="usado_orfao")
            pd.DataFrame(sorted([str(p.relative_to(ROOT)) for p in all_py]), columns=["py"]).to_excel(writer, index=False, sheet_name="lista_py")
            pd.DataFrame(sorted([str(p.relative_to(ROOT)) for p in all_kv]), columns=["kv"]).to_excel(writer, index=False, sheet_name="lista_kv")
        print("[OK] Excel salvo em: arquivos_usados_orfaos.xlsx")
    except Exception as e:
        print("[INFO] openpyxl não disponível; gerado apenas CSV. Instale com: pip install openpyxl")

    # Resumo
    tot_py = sum(1 for r in rows if r["tipo"] == "py")
    tot_kv = sum(1 for r in rows if r["tipo"] == "kv")
    usados_py = sum(1 for r in rows if r["tipo"] == "py" and r["status"] == "USADO")
    usados_kv = sum(1 for r in rows if r["tipo"] == "kv" and r["status"] == "USADO")
    print(f"PY: {usados_py}/{tot_py} usados | KV: {usados_kv}/{tot_kv} usados")

if __name__ == "__main__":
    main()
