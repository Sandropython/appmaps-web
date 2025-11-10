# -*- coding: utf-8 -*-
"""
main_relatorios.py
------------------
Gera:
  1) EnvioWhatsapp.xlsx
  2) Rota Completa (PDF)
  3) QRCode (PDF)

+ NOVO: opção de inserir miniatura (Maps Static) por parada no PDF de QR.

Uso (CLI):
  python main_relatorios.py "C:\\...\\minha_base.xlsx" --saida whatsapp
  python main_relatorios.py "C:\\...\\minha_base.xlsx" --saida rota_pdf
  python main_relatorios.py "C:\\...\\minha_base.xlsx" --saida qrcode_pdf [--maps-link] [--miniatura]
"""

from __future__ import annotations
import os, io, argparse, re
from typing import Dict, List, Optional, Tuple

import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

import qrcode
import requests

# --------------------------
# Utilidades de E/S
# --------------------------
def _read_any(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    if ext == ".csv":
        return pd.read_csv(path)
    if ext == ".pkl":
        return pd.read_pickle(path)
    raise ValueError(f"Extensão não suportada: {ext}")

def _stem_dir(path: str) -> Tuple[str, str]:
    outdir = os.path.dirname(path)
    stem = os.path.splitext(os.path.basename(path))[0]
    return outdir, stem

# --------------------------
# Chave Google (env/.env)
# --------------------------
def _load_google_key() -> str:
    k = os.getenv("GOOGLE_API_KEY", "") or os.getenv("GOOGLE_DIRECTIONS_KEY", "")
    if k:
        return k
    # .env ao lado do script
    env = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env):
        with open(env, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("GOOGLE_API_KEY="):
                    return line.split("=",1)[1].strip().strip('"').strip("'")
                if line.startswith("GOOGLE_DIRECTIONS_KEY="):
                    return line.split("=",1)[1].strip().strip('"').strip("'")
    return ""

# --------------------------
# Pastas padronizadas de saída
# --------------------------
APP_DIR   = os.path.dirname(__file__)
BASES_DIR = os.path.join(APP_DIR, "bases")
OUT_DIRS = {
    "html": os.path.join(BASES_DIR, "html"),
    "pdf": os.path.join(BASES_DIR, "pdf"),
    "txt": os.path.join(BASES_DIR, "txt"),
    "xlsx": os.path.join(BASES_DIR, "xlsx"),
    "xlsx_envio": os.path.join(BASES_DIR, "xlsx", "envio"),  # exceção
}
for _p in OUT_DIRS.values():
    os.makedirs(_p, exist_ok=True)

def _outpath(kind: str, filename: str) -> str:
    base = OUT_DIRS[kind]
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, filename)

# --------------------------
# Detecção de colunas (tolerante)
# --------------------------
COLS: Dict[str, List[str]] = {
    "ordem":   ["ordem", "order", "índice", "indice"],
    "sequence":["sequence", "sequência", "sequencia", "seq", "pedidos"],
    "id":      ["spx tn", "stx", "id", "rastreador", "ref", "referencia", "referência", "pedido"],
    "nome":    ["nome", "cliente", "destinatário", "destinatario", "name"],
    "fone":    ["telefone", "phone", "celular", "mobile", "whatsapp"],
    "lat":     ["latitude", "lat", "y"],
    "lon":     ["longitude", "long", "lon", "x"],
    "local":   ["destination address", "destino", "endereco", "endereço", "local", "address"],
}

def _find_col(df: pd.DataFrame, keys: List[str]) -> Optional[str]:
    low = {c.lower(): c for c in df.columns}
    for k in keys:
        if k.lower() in low:
            return low[k.lower()]
    # “contém”
    for c in df.columns:
        lc = c.lower()
        for k in keys:
            if k.lower() in lc:
                return c
    return None

def _ensure(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    cols = {k: _find_col(df, v) for k, v in COLS.items()}
    return cols

# --------------------------
# 1) Envio Whatsapp.xlsx
# --------------------------
def _clean_phone(s: pd.Series) -> pd.Series:
    """
    Normaliza telefones vindos como int/float/string.
    - Remove sufixo '.0' quando lidos como float (ex.: 51999998888.0)
    - Converte notação científica para string decimal (ex.: 5.1999998888e+10)
    - Remove tudo que não for dígito no final
    """
    def _one(x) -> str:
        if pd.isna(x):
            return ""
        if isinstance(x, (int,)):
            return str(x)
        if isinstance(x, float):
            if x.is_integer():
                return str(int(x))
            from decimal import Decimal, getcontext
            try:
                getcontext().prec = 40
                d = Decimal(str(x))
                return re.sub(r"\D", "", format(d, "f"))
            except Exception:
                return re.sub(r"\D", "", str(x))

        st = str(x).strip().replace("\u00a0", "").replace(" ", "")
        if re.search(r"^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$", st):
            from decimal import Decimal, getcontext
            try:
                getcontext().prec = 40
                d = Decimal(st)
                st = format(d, "f")
            except Exception:
                pass
        st = re.sub(r"\.0$", "", st)
        st = re.sub(r"\D", "", st)
        return st

    return s.apply(_one)

def gerar_envio_whatsapp_from_path(path: str) -> str:
    df = _read_any(path)
    cols = _ensure(df)

    need = ["sequence", "id", "nome", "fone"]
    if not all(cols.get(k) for k in need):
        raise ValueError("Planilha sem colunas suficientes (precisa de Sequence, ID, Nome, Telefone).")

    seq = cols["sequence"]; cid = cols["id"]; cnome = cols["nome"]; cfone = cols["fone"]

    data = df[[seq, cid, cnome, cfone]].copy()
    data[cfone] = _clean_phone(data[cfone])
    data = data[data[cfone].astype(str).str.len() > 0]

    data["Nome"] = "P-" + data[seq].astype(str) + " - " + data[cid].astype(str) + " - " + data[cnome].astype(str)
    data["Telefone"] = data[cfone]
    data["Grupo"] = "fam"
    data["Mensagem"] = "Bom dia cliente Shoopee!"
    data["Arquivo"] = "n"
    out = data[["Nome", "Telefone", "Grupo", "Mensagem", "Arquivo"]].reset_index(drop=True)

    rows = len(out)
    out_xlsx = _outpath("xlsx_envio", "EnvioWhatsapp.xlsx")

    if rows <= 40:
        out.to_excel(out_xlsx, index=False, sheet_name="envio")
    elif rows <= 90:
        with pd.ExcelWriter(out_xlsx) as w:
            out.iloc[:40].to_excel(w, index=False, sheet_name="envio")
            out.iloc[40:].to_excel(w, index=False, sheet_name="envio1")
    else:
        with pd.ExcelWriter(out_xlsx) as w:
            out.iloc[:40].to_excel(w, index=False, sheet_name="envio")
            out.iloc[40:80].to_excel(w, index=False, sheet_name="envio1")
            out.iloc[80:].to_excel(w, index=False, sheet_name="envio2")
    return out_xlsx

# --------------------------
# 2) Rota Completa (PDF)
# --------------------------
def _string_w(c: canvas.Canvas, text: str, font="Helvetica", size=10) -> float:
    c.setFont(font, size)
    return c.stringWidth(text, font, size)

def gerar_rota_pdf_from_df(df: pd.DataFrame, path_out_pdf: str, pagina="A4"):
    """tabela simples (Ordem, Sequence), com borda e paginação"""
    page = A4 if pagina.upper()=="A4" else letter
    W, H = page
    c = canvas.Canvas(path_out_pdf, pagesize=page)

    margem = 12*mm
    y = H - margem
    titulo = "ROTA COMPLETA"
    c.setFont("Helvetica-Bold", 16)
    c.drawString((W - _string_w(c, titulo, "Helvetica-Bold", 16))/2, y, titulo)
    y -= 12*mm

    c_ordem = _find_col(df, COLS["ordem"]) or "Ordem"
    c_seq   = _find_col(df, COLS["sequence"]) or "Sequence"
    if c_ordem not in df.columns:
        df = df.copy()
        df[c_ordem] = range(1, len(df)+1)

    ordem_w = 20*mm
    max_seq_len = max((len(str(s)) for s in df[c_seq].astype(str)), default=3)
    seq_w = min(max(40*mm, max_seq_len*3.2*mm), W - 2*margem - ordem_w)

    row_h = 9*mm

    c.setFont("Helvetica-Bold", 12)
    c.setStrokeColor(colors.black)
    c.rect(margem, y-row_h*0.8, ordem_w, row_h)
    c.rect(margem+ordem_w, y-row_h*0.8, seq_w, row_h)
    c.drawString(margem + ordem_w/2 - _string_w(c,"Ordem","Helvetica-Bold",12)/2, y-2, "Ordem")
    c.drawString(margem + ordem_w + 4, y-2, "Sequence")
    y -= row_h

    c.setFont("Helvetica", 11)

    for _, r in df.iterrows():
        if y < margem + row_h:
            c.showPage()
            y = H - margem
            c.setFont("Helvetica-Bold", 12)
            c.rect(margem, y-row_h*0.8, ordem_w, row_h)
            c.rect(margem+ordem_w, y-row_h*0.8, seq_w, row_h)
            c.drawString(margem + ordem_w/2 - _string_w(c,"Ordem","Helvetica-Bold",12)/2, y-2, "Ordem")
            c.drawString(margem + ordem_w + 4, y-2, "Sequence")
            y -= row_h
            c.setFont("Helvetica", 11)

        c.rect(margem, y-row_h*0.8, ordem_w, row_h)
        c.rect(margem+ordem_w, y-row_h*0.8, seq_w, row_h)
        c.drawString(margem + ordem_w/2 - _string_w(c,str(r[c_ordem]))/2, y-2, str(r[c_ordem]))
        c.drawString(margem + ordem_w + 4, y-2, str(r[c_seq]))
        y -= row_h

    c.save()

def gerar_rota_pdf_from_path(path: str, pagina="A4") -> str:
    df = _read_any(path)
    _, stem = _stem_dir(path)
    out = _outpath("pdf", f"{stem}_Rota_Completa.pdf")
    gerar_rota_pdf_from_df(df, out, pagina=pagina)
    return out

# --------------------------
# 3) QRCode (PDF) + Miniatura de mapa (opcional)
# --------------------------
def _make_qr_image(data: str, box=4, border=1) -> ImageReader:
    qr = qrcode.QRCode(
        version=None, error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box, border=border
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return ImageReader(bio)

def _maps_static_image(lat: float, lon: float, key: str,
                       zoom: int = 16, size_px: Tuple[int,int] = (160, 160)) -> Optional[ImageReader]:
    """
    Baixa uma miniatura da parada via Maps Static API e retorna ImageReader.
    """
    if not key:
        return None
    url = "https://maps.googleapis.com/maps/api/staticmap"
    params = {
        "center": f"{lat},{lon}",
        "zoom": str(zoom),
        "size": f"{size_px[0]}x{size_px[1]}",
        "maptype": "roadmap",
        "markers": f"color:red|{lat},{lon}",
        "key": key,
        "language": "pt-BR",
        "scale": "2"  # mais nitidez
    }
    try:
        r = requests.get(url, params=params, timeout=30)
        r.raise_for_status()
        bio = io.BytesIO(r.content)
        bio.seek(0)
        return ImageReader(bio)
    except Exception as e:
        print(f"[Maps Static] falhou ({e})")
        return None

def gerar_qrcode_pdf_from_path(path: str, usar_sequence=True, usar_link_maps=False,
                               usar_miniatura: bool = False) -> str:
    """
    Gera PDF com colunas:
      N_Pedido | SPX TN | Local (QR) | [Mapa] | SPX TN (QR)

    - 'Mapa' (miniatura) aparece se usar_miniatura=True E houver lat/lon.
    """
    df = _read_any(path)
    cols = _ensure(df)
    c_seq = cols["sequence"]
    c_id  = cols["id"]
    c_lat = cols["lat"]; c_lon = cols["lon"]
    if not c_id:
        raise ValueError("Coluna de ID (SPX TN / STX / ID) não encontrada.")
    if usar_sequence and not c_seq:
        raise ValueError("Coluna Sequence não encontrada.")

    _, stem = _stem_dir(path)
    out_pdf = _outpath("pdf", f"{stem}_QRCODE.pdf")
    c = canvas.Canvas(out_pdf, pagesize=A4)
    W, H = A4
    margem = 15*mm
    y = H - margem

    c.setFont("Helvetica-Bold", 14)
    title = "QR Codes"
    c.drawString((W - _string_w(c, title, "Helvetica-Bold", 14))/2, y, title)
    y -= 10*mm

    # --- Larguras (A4 útil ~180mm) ---
    col1 = 25*mm   # N_Pedido (reduzido)
    col2 = 55*mm   # SPX TN texto (reduzido)
    col3 = 25*mm   # Local (QR)
    col_map = 30*mm if (usar_miniatura and c_lat and c_lon) else 0*mm  # Miniatura opcional
    col4 = 25*mm   # SPX TN (QR)
    row_h = 12*mm
    total_w = col1 + col2 + col3 + col_map + col4

    # --- Cabeçalho ---
    c.setFont("Helvetica-Bold", 12)
    x = margem
    def _head(label, width):
        nonlocal x
        c.rect(x, y-row_h*0.8, width, row_h)
        c.drawString(x + width/2 - _string_w(c,label,"Helvetica-Bold",12)/2, y-2, label)
        x += width

    _head("N_Pedido", col1)
    _head("SPX TN", col2)
    _head("Local (QR)", col3)
    if col_map:
        _head("Mapa", col_map)
    _head("SPX TN (QR)", col4)
    y -= row_h
    c.setFont("Helvetica", 11)

    def maps_link(lat: float, lon: float) -> str:
        return f"https://www.google.com/maps/dir/?api=1&destination={lat:.6f},{lon:.6f}&travelmode=driving"

    api_key = _load_google_key()

    for _, r in df.iterrows():
        if y < margem + row_h + 40:
            c.showPage(); y = H - margem
            c.setFont("Helvetica-Bold", 12)
            x = margem
            _head("N_Pedido", col1)
            _head("SPX TN", col2)
            _head("Local (QR)", col3)
            if col_map:
                _head("Mapa", col_map)
            _head("SPX TN (QR)", col4)
            y -= row_h
            c.setFont("Helvetica", 11)

        n_pedido = str(r[c_seq]) if (usar_sequence and c_seq) else ""
        id_ped   = str(r[c_id])

        # conteúdo do QR (Local)
        if usar_link_maps and c_lat and c_lon and pd.notna(r.get(c_lat)) and pd.notna(r.get(c_lon)):
            content_local = maps_link(float(r[c_lat]), float(r[c_lon]))
        else:
            # mesmo sem maps-link, mantemos um QR (ex.: endereço)
            local_txt = str(r.get(cols["local"], "")) if cols.get("local") else id_ped
            content_local = local_txt

        qr_local = _make_qr_image(content_local)
        qr_spx   = _make_qr_image(id_ped)

        # células
        x = margem
        # N_Pedido
        c.rect(x, y-row_h*0.8, col1, row_h)
        c.drawString(x + col1/2 - _string_w(c, n_pedido)/2, y-2, n_pedido)
        x += col1

        # SPX TN (texto)
        c.rect(x, y-row_h*0.8, col2, row_h)
        c.drawString(x + 4, y-2, id_ped)
        x += col2

        # Local (QR)
        c.rect(x, y-row_h*0.8, col3, row_h)
        qr_w = qr_h = 10*mm
        c.drawImage(qr_local, x + (col3-qr_w)/2, y - qr_h + 2, qr_w, qr_h, preserveAspectRatio=True, mask='auto')
        x += col3

        # Miniatura (opcional)
        if col_map and usar_miniatura and c_lat and c_lon and pd.notna(r.get(c_lat)) and pd.notna(r.get(c_lon)):
            c.rect(x, y-row_h*0.8, col_map, row_h)
            thumb = _maps_static_image(float(r[c_lat]), float(r[c_lon]), api_key, zoom=16, size_px=(140,140))
            if thumb:
                # altura da linha é 12mm -> miniatura bem pequena; vamos encaixar como "selo"
                th = 10*mm; tw = 10*mm
                c.drawImage(thumb, x + (col_map-tw)/2, y - th + 2, tw, th, preserveAspectRatio=True, mask='auto')
            x += col_map

        # SPX TN (QR)
        c.rect(x, y-row_h*0.8, col4, row_h)
        c.drawImage(qr_spx, x + (col4-qr_w)/2, y - qr_h + 2, qr_w, qr_h, preserveAspectRatio=True, mask='auto')

        y -= row_h

    c.save()
    return out_pdf

# --------------------------
# CLI
# --------------------------
def _parse():
    p = argparse.ArgumentParser()
    p.add_argument("arquivo", help="xlsx/xls/csv/pkl de entrada")
    p.add_argument("--saida", choices=["whatsapp","rota_pdf","qrcode_pdf"], required=True)
    p.add_argument("--maps-link", action="store_true", help="QR usa link do Google Maps (se tiver lat/lon)")
    p.add_argument("--miniatura", action="store_true", help="Inclui miniatura do Maps Static no PDF de QR (se tiver lat/lon)")
    p.add_argument("--pagina", choices=["A4","LETTER"], default="A4")
    return p.parse_args()

def main():
    a = _parse()
    if a.saida == "whatsapp":
        out = gerar_envio_whatsapp_from_path(a.arquivo)
    elif a.saida == "rota_pdf":
        out = gerar_rota_pdf_from_path(a.arquivo, pagina=a.pagina)
    else:
        out = gerar_qrcode_pdf_from_path(
            a.arquivo,
            usar_sequence=True,
            usar_link_maps=a.maps_link,
            usar_miniatura=a.miniatura
        )
    print(out)

if __name__ == "__main__":
    main()
