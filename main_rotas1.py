# -*- coding: utf-8 -*-
"""
Engine de geração do mapa (Folium) com coloração por critério.

NÃO ALTERA PONTOS DA BASE.
- sequence/melhor: 0 (depósito) laranja; 1-2 roxo; depois muda a cor a cada 15 pontos
- cep/bairro: mesma cor por grupo
- Camada 'Resumo da rota (KM)'
- Camadas 'Rota (ida/volta)', 'Rota (ida)', 'Rota (volta)'
- Botões Google Maps/Waze no popup
- (Opcional) Google Directions para linha/KM reais
- 'melhor': vizinho mais próximo + 2-OPT (round-trip) usando custo de TEMPO previsto (opcional)
- LayerControl colapsado + botão "Desmarcar" fixo
- Persistência via localStorage (camadas + view)

Dep.: pandas, folium, openpyxl, requests, polyline (se usar Google).
"""
from __future__ import annotations
import os
from math import radians, sin, cos, sqrt, atan2
from typing import Dict, Optional, List, Tuple

import pandas as pd
import folium
from folium.plugins import Draw, Fullscreen
import requests
try:
    import polyline as _poly
except Exception:
    _poly = None

from branca.element import Template, MacroElement

# ===== Integração opcional com TEMPO previsto (trânsito) =====
try:
    from directions_predictive import (
        build_time_cost_matrix,
        route_time_cost,
        two_opt_on_time_matrix,
    )
except Exception:
    build_time_cost_matrix = None
    route_time_cost = None
    two_opt_on_time_matrix = None

# ==========================
# PASTAS DO APP (saídas)
# ==========================
APP_DIR   = os.path.dirname(__file__)
BASES_DIR = os.path.join(APP_DIR, "bases")
HTML_DIR  = os.path.join(BASES_DIR, "html")
LOGS_DIR  = os.path.join(BASES_DIR, "logs")
os.makedirs(HTML_DIR, exist_ok=True)
os.makedirs(LOGS_DIR,  exist_ok=True)

# ==========================
# CONFIG PADRÃO (modo script)
# ==========================
INPUT_PATH = r"C:\Users\sandr\OneDrive\2025\Shopee\Rota\14-08-2025.xlsx"
CRITERIO   = "sequence"     # "ordem" | "sequence" | "bairro" | "cep" | "melhor"
BASE_NAME  = "Mapa base"

MOSTRAR_LINHA          = False
MOSTRAR_ROTA_IDA_VOLTA = True
MOSTRAR_PIN_DEPOSITO   = True

# ---------- Validação LEVE de chave (sem custo) ----------
def _is_probably_google_key(k: str) -> bool:
    if not isinstance(k, str):
        return False
    k = k.strip()
    return bool(k) and k.startswith("AIza") and (30 <= len(k) <= 60)

# IMPORTANTÍSSIMO: nada de pegar chave de .env/ENV aqui no engine.
# O engine só respeita a chave recebida nos parâmetros da função pública.
GOOGLE_API_KEY = ""          # <- propositalmente vazio (NÃO USADO)
GOOGLE_LANG    = "pt-BR"
GOOGLE_MAX_POINTS = 25

# ===== FLAGS =====
USE_GOOGLE_GEOCODING        = False  # mantido OFF
USE_GOOGLE_DISTANCE_MATRIX  = False  # opcional; ainda assim só com chave válida
USE_GOOGLE_ROADS            = False  # mantido OFF

INCLUIR_DEPOSITO_NO_KM = True
CALCULAR_KM_OTIMIZADO  = True

# ===== Usar tempo previsto (só com chave válida) =====
USE_PREDICTIVE_TIME   = True
PREDICTIVE_LOG_PATH   = os.path.join(LOGS_DIR, "predictive_times.csv")

# ==========================
# Colunas (tolerante)
# ==========================
COLS: Dict[str, List[str]] = {
    "ordem":   ["ordem", "order", "índice", "indice"],
    "sequence":["sequence", "sequência", "sequencia", "seq", "pedidos"],
    "bairro":  ["bairro", "district", "neighborhood", "neighbourhood"],
    "cep":     ["cep", "zip", "zipcode", "postal", "código postal", "codigo postal"],
    "lat":     ["latitude", "lat", "y"],
    "lon":     ["longitude", "long", "lon", "x"],
    "local":   ["destination address", "destino", "endereco", "endereço", "local", "address"],
    "cidade":  ["cidade", "city", "município", "municipio"],
    "nome":    ["nome", "cliente", "name"],
    "fone":    ["telefone", "phone", "celular", "mobile"],
    "id":      ["spx tn", "stx", "id", "rastreador", "ref", "referencia", "referência"],
    "parada":  ["parada", "stop"],
}

def _find_col(df: pd.DataFrame, keys: List[str]) -> Optional[str]:
    low = {c.lower(): c for c in df.columns}
    for k in keys:
        if k.lower() in low:
            return low[k.lower()]
    for c in df.columns:
        lc = c.lower()
        for k in keys:
            if k.lower() in lc:
                return c
    return None

def read_any(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"): return pd.read_excel(path)
    if ext == ".csv":            return pd.read_csv(path)
    if ext == ".pkl":            return pd.read_pickle(path)
    raise ValueError(f"Extensão não suportada: {ext}")

def ensure_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    cols = {k: _find_col(df, v) for k, v in COLS.items()}
    return cols

def _to_num(s):
    try:
        return pd.to_numeric(s, errors="coerce")
    except Exception:
        return pd.Series([None] * len(s))

def clean_base(df: pd.DataFrame, c: Dict[str, str]) -> pd.DataFrame:
    dd = df.copy()
    if c.get("lat") in dd and c.get("lon") in dd:
        dd[c["lat"]] = _to_num(dd[c["lat"]]); dd[c["lon"]] = _to_num(dd[c["lon"]])
    if c.get("nome"): dd[c["nome"]] = dd[c["nome"]].astype(str).str.title()
    if c.get("fone"): dd[c["fone"]] = dd[c["fone"]].astype(str).str.replace(r"\D", "", regex=True)
    if c.get("nome") and c.get("fone"):
        dd["INFO"] = dd[c["nome"]] + "-" + dd[c["fone"]].fillna("")
    else:
        dd["INFO"] = dd.get("INFO", "")
    if c.get("sequence"): dd[c["sequence"]] = _to_num(dd[c["sequence"]])
    if c.get("ordem"):    dd[c["ordem"]]    = _to_num(dd[c["ordem"]])
    return dd

# ---------- Distâncias ----------
def haversine_km(lat1, lon1, lat2, lon2) -> float:
    R = 6371.0088
    p1, p2 = radians(lat1), radians(lat2)
    dphi = radians(lat2 - lat1); dl = radians(lon2 - lon1)
    a = sin(dphi/2)**2 + cos(p1)*cos(p2)*sin(dl/2)**2
    return 2*R*atan2(sqrt(a), sqrt(1-a))

def path_len_haversine(coords: List[Tuple[float,float]]) -> float:
    if len(coords) < 2: return 0.0
    tot = 0.0
    for a, b in zip(coords, coords[1:]):
        tot += haversine_km(a[0], a[1], b[0], b[1])
    return tot

# ---------- Ordenação ----------
def ordenar(df: pd.DataFrame, c: Dict[str, str], criterio: str, ascending: bool=True) -> pd.DataFrame:
    dd = df.copy(); crit = (criterio or "").lower()
    if   crit == "ordem"    and c.get("ordem"):
        dd = dd.sort_values(c["ordem"], ascending=ascending, kind="mergesort")
    elif crit == "sequence" and c.get("sequence"):
        dd = dd.sort_values(c["sequence"], ascending=ascending, kind="mergesort")
    elif crit == "bairro"   and c.get("bairro"):
        extra = [x for x in [c.get("ordem"), c.get("sequence")] if x]
        asc_list = [ascending] + [True]*len(extra)
        dd = dd.sort_values([c["bairro"]] + extra, ascending=asc_list, kind="mergesort")
    elif crit == "cep"      and c.get("cep"):
        extra = [x for x in [c.get("ordem"), c.get("sequence")] if x]
        asc_list = [ascending] + [True]*len(extra)
        dd = dd.sort_values([c["cep"]] + extra, ascending=asc_list, kind="mergesort")
    return dd.reset_index(drop=True)

def cat_uniq(series: pd.Series) -> str:
    vals = [str(v) for v in series.dropna().astype(str)]
    seen, out = set(), []
    for v in vals:
        if v not in seen: out.append(v); seen.add(v)
    return " ".join(out) if out else ""

def agrupar_por_coord_preservando_ordem(df_ordenado: pd.DataFrame, c: Dict[str, str]) -> pd.DataFrame:
    key_cols = [c["lat"], c["lon"]]
    g = (df_ordenado.groupby(key_cols, sort=False, as_index=False)
         .agg({
            c["sequence"]: cat_uniq if c.get("sequence") else "first",
            c["id"]:       cat_uniq if c.get("id")       else "first",
            "INFO":        cat_uniq,
            c["local"]:    cat_uniq if c.get("local")    else "first",
            c["bairro"]:   "first" if c.get("bairro") else (lambda s: ""),
            c["cidade"]:   "first" if c.get("cidade") else (lambda s: ""),
            c["parada"]:   "first" if c.get("parada") else (lambda s: ""),
            c["cep"]:      "first" if c.get("cep") else (lambda s: ""),
         }))
    col_map = {
        c.get("sequence","Sequence"): "Sequence",
        c.get("id","ID"):             "ID",
        "INFO":                       "INFO",
        c.get("local","LOCAL"):       "LOCAL",
        c.get("bairro","Bairro"):     "Bairro",
        c.get("cidade","Cidade"):     "Cidade",
        c.get("parada","Parada"):     "Parada",
        c.get("cep","CEP"):           "CEP",
        c["lat"]:                     "Latitude",
        c["lon"]:                     "Longitude",
    }
    g = g.rename(columns=col_map)
    g.insert(0, "Ordem", range(1, len(g)+1))
    for need in ["Sequence","ID","LOCAL","Bairro","Cidade","Parada","CEP"]:
        if need not in g.columns: g[need] = ""
    return g[["Ordem","Sequence","Parada","ID","INFO","LOCAL","CEP","Bairro","Cidade","Latitude","Longitude"]]

# ---------- Depósito ----------
def detectar_deposito_por_sequence(df_base_limpo: pd.DataFrame, c: Dict[str, str]) -> Optional[Tuple[float, float]]:
    if c.get("sequence") and df_base_limpo[c["sequence"]].notna().any():
        df_ok = df_base_limpo.dropna(subset=[c["sequence"], c["lat"], c["lon"]]).copy()
        if len(df_ok):
            df_ok = df_ok.sort_values(c["sequence"], kind="mergesort")
            r0 = df_ok.iloc[0]
            return float(r0[c["lat"]]), float(r0[c["lon"]])
    df_any = df_base_limpo.dropna(subset=[c["lat"], c["lon"]]) if c.get("lat") and c.get("lon") else df_base_limpo
    if len(df_any) and c.get("lat") and c.get("lon"):
        r0 = df_any.iloc[0]
        return float(r0[c["lat"]]), float(r0[c["lon"]])
    return None

# ---------- Heurística "melhor": NN + 2-OPT ----------
def _nearest_neighbor_order(coords: List[Tuple[float,float]], start_idx: int, time_matrix: Optional[List[List[float]]] = None) -> List[int]:
    n = len(coords)
    if n <= 1: return list(range(n))
    unvis = set(range(n))
    cur = start_idx
    order = [cur]
    unvis.remove(cur)

    def cost(i, j):
        if time_matrix is not None:
            return time_matrix[i][j]
        A, B = coords[i], coords[j]
        return haversine_km(A[0], A[1], B[0], B[1])

    while unvis:
        nxt = min(unvis, key=lambda j: cost(cur, j))
        order.append(nxt)
        unvis.remove(nxt)
        cur = nxt
    return order

def _two_opt(order: List[int], coords: List[Tuple[float,float]]) -> List[int]:
    improved = True
    best = order[:]
    def seg_len(a,b):
        A,B = coords[a], coords[b]
        return haversine_km(A[0],A[1],B[0],B[1])
    while improved:
        improved = False
        for i in range(1, len(best)-2):
            for k in range(i+1, len(best)-1):
                a,b = best[i-1], best[i]
                c,d = best[k], best[k+1]
                delta = (seg_len(a,c) + seg_len(b,d)) - (seg_len(a,b) + seg_len(c,d))
                if delta < -1e-6:
                    best[i:k+1] = reversed(best[i:k+1])
                    improved = True
    return best

def _ordem_melhor(grp: pd.DataFrame, deposito: Tuple[float,float], google_key: str, usar_tempo: bool) -> pd.DataFrame:
    pts = [(float(r["Latitude"]), float(r["Longitude"])) for _, r in grp.iterrows()]
    if len(pts) <= 2:
        out = grp.reset_index(drop=True).copy()
        out["Ordem"] = range(1, len(out)+1)
        return out

    d0 = [haversine_km(deposito[0], deposito[1], p[0], p[1]) for p in pts]
    seed = int(min(range(len(pts)), key=lambda i: d0[i]))

    # Matriz de tempo só se chave for válida
    time_matrix = None
    if usar_tempo and build_time_cost_matrix and _is_probably_google_key(google_key):
        try:
            time_matrix = build_time_cost_matrix([deposito] + pts, google_key, PREDICTIVE_LOG_PATH)
            time_matrix = [row[1:] for row in time_matrix[1:]]
        except Exception as e:
            print(f"[predictive] falhou: {e}")
            time_matrix = None

    nn = _nearest_neighbor_order(pts, seed, time_matrix=time_matrix)
    if time_matrix is not None and two_opt_on_time_matrix:
        better = two_opt_on_time_matrix(nn, time_matrix)
    else:
        better = _two_opt(nn, pts)

    out = grp.iloc[better].reset_index(drop=True).copy()
    out["Ordem"] = range(1, len(out)+1)
    return out

# ---------- Google Directions ----------
def _google_one_call(points: List[Tuple[float, float]],
                     api_key: str,
                     idioma: str = "pt-BR",
                     optimize: bool = False) -> Tuple[float, List[Tuple[float, float]]]:
    if _poly is None:
        raise RuntimeError("Instale 'polyline' (pip install polyline).")
    if len(points) < 2:
        return 0.0, []
    origin = f"{points[0][0]},{points[0][1]}"
    destination = f"{points[-1][0]},{points[-1][1]}"
    wps = points[1:-1]
    waypoints = "|".join(f"{lat},{lon}" for lat, lon in wps) if wps else ""
    url = "https://maps.googleapis.com/maps/api/directions/json"
    params = {
        "origin": origin, "destination": destination,
        "mode": "driving", "language": idioma, "key": api_key,
        "departure_time": "now",
    }
    if waypoints:
        params["waypoints"] = ("optimize:true|" if optimize else "optimize:false|") + waypoints
    r = requests.get(url, params=params, timeout=60); r.raise_for_status()
    data = r.json()
    if not data.get("routes"):
        raise RuntimeError(f"Directions sem rotas. status={data.get('status')} {data.get('error_message','')}")
    route = data["routes"][0]
    dist_m = sum(leg["distance"]["value"] for leg in route["legs"])
    poly_str = route["overview_polyline"]["points"]
    line = [(lat, lon) for lat, lon in _poly.decode(poly_str)]
    return dist_m / 1000.0, line

def google_directions_roundtrip_chunked(coords: List[Tuple[float, float]],
                                        api_key: str,
                                        idioma: str = "pt-BR",
                                        optimize: bool = False) -> Tuple[float, List[Tuple[float, float]]]:
    n = len(coords)
    if n < 2: return 0.0, []
    total = 0.0
    full_line: List[Tuple[float, float]] = []
    i = 0
    while i < n - 1:
        j = min(i + GOOGLE_MAX_POINTS - 1, n - 1)
        chunk = coords[i:j+1]
        km, line = _google_one_call(chunk, api_key, idioma, optimize=optimize)
        total += km
        if not full_line: full_line.extend(line)
        else:
            if line: full_line.extend(line[1:])
        i = j
    return round(total, 2), full_line

# ---------- Botões ----------
def nav_buttons(lat: float, lon: float, deposito: Optional[Tuple[float, float]] = None) -> str:
    gm_cur = f"https://www.google.com/maps/dir/?api=1&destination={lat},{lon}&travelmode=driving"
    wz_cur = f"https://waze.com/ul?ll={lat},{lon}&navigate=yes"
    html = [
        '<div style="margin-top:8px;">',
        '<div style="display:flex; gap:6px; flex-wrap:wrap;">',
        f'<a href="{gm_cur}" target="_blank" rel="noopener" '
        'style="text-decoration:none;padding:6px 8px;border-radius:6px;'
        'background:#1a73e8;color:#fff;font-weight:600;">Google Maps</a>',
        f'<a href="{wz_cur}" target="_blank" rel="noopener" '
        'style="text-decoration:none;padding:6px 8px;border-radius:6px;'
        'background:#0b8043;color:#fff;font-weight:600;">Waze</a>',
    ]
    if deposito:
        gm_dep = f"https://www.google.com/maps/dir/?api=1&origin={deposito[0]},{deposito[1]}&destination={lat},{lon}&travelmode=driving"
        wz_dep = f"https://waze.com/ul?ll={lat},{lon}&from={deposito[0]},{deposito[1]}&navigate=yes"
        html += [
            f'<a href="{gm_dep}" target="_blank" rel="noopener" '
            'style="text-decoration:none;padding:6px 8px;border-radius:6px;'
            'background:#1a73e8;color:#fff;">Google (depósito)</a>',
            f'<a href="{wz_dep}" target="_blank" rel="noopener" '
            'style="text-decoration:none;padding:6px 8px;border-radius:6px;'
            'background:#0b8043;color:#fff;">Waze (depósito)</a>',
        ]
    html += ['</div>', '</div>']
    return "".join(html)

# ---------- UI helpers ----------
class FixedUncheckAllControl(MacroElement):
    _tpl = Template(r"""
    {% macro header(this, kwargs) %}
    <style>
      .pin-uncheck-btn{
        position: fixed;
        right: 18px;
        bottom: 18px;
        z-index: 10000;
        background: #fff;
        border: 1px solid #bbb;
        border-radius: 6px;
        padding: 6px 10px;
        font: 600 12px/1 Arial, sans-serif;
        box-shadow: 0 1px 3px rgba(0,0,0,.35);
        color: #111;
        text-decoration: none;
        user-select: none;
      }
      .pin-uncheck-btn:hover{ background:#f7f7f7; }
      @media print { .pin-uncheck-btn{ display:none; } }
    </style>
    {% endmacro %}
    {% macro html(this, kwargs) %}
      <a class="pin-uncheck-btn" href="#" id="pin-uncheck">Desmarcar</a>
    {% endmacro %}
    {% macro script(this, kwargs) %}
      (function(){
        var a = document.getElementById('pin-uncheck');
        if(!a) return;
        a.addEventListener('click', function(e){
          e.preventDefault(); e.stopPropagation();
          const boxes=document.querySelectorAll('.leaflet-control-layers-overlays input[type=checkbox]');
          boxes.forEach(cb => { if (cb.checked) cb.click(); });
        }, {passive:false});
      })();
    {% endmacro %}
    """)
    def __init__(self):
        super().__init__()
        self._template = self._tpl

def add_uncheck_all_button(m: folium.Map):
    m.get_root().add_child(FixedUncheckAllControl())

class PersistLayerStateControl(MacroElement):
    def __init__(self, storage_key: str):
        super().__init__()
        self.storage_key = storage_key
        self._template = Template(r"""
        {% macro script(this, kwargs) %}
        (function(){
          var map = {{ this._parent.get_name() }};
          var KEY = {{ this.storage_key | tojson }};
          var KLAY = KEY + ":layers";
          var KVIEW = KEY + ":view";

          function readJSON(k, def){ try{ return JSON.parse(localStorage.getItem(k)) ?? def; }catch(_){ return def; } }
          function writeJSON(k, v){ try{ localStorage.setItem(k, JSON.stringify(v)); }catch(_){ } }

          function getOverlayItems(){
            const items = [];
            const root = document.querySelector('.leaflet-control-layers-overlays');
            if(!root) return items;
            root.querySelectorAll('label').forEach(function(lbl){
              const name = (lbl.textContent || '').trim();
              const cb = lbl.querySelector('input[type=checkbox]');
              if (cb && name) items.push({ name, cb });
            });
            return items;
          }

          function saveLayers(){
            const obj = {};
            getOverlayItems().forEach(it => obj[it.name] = !!it.cb.checked);
            writeJSON(KLAY, obj);
          }

          function applySavedLayers(){
            const saved = readJSON(KLAY, null);
            if (!saved) return false;
            const items = getOverlayItems();
            if (!items.length) return false;
            items.forEach(function(it){
              const want = saved.hasOwnProperty(it.name) ? !!saved[it.name] : it.cb.checked;
              if (!!it.cb.checked !== want){
                it.cb.click();
              }
            });
            return true;
          }

          function saveView(){
            try{
              var c = map.getCenter();
              var z = map.getZoom();
              writeJSON(KVIEW, {center:[c.lat, c.lng], zoom:z});
            }catch(_){}
          }

          function applySavedView(){
            const view = readJSON(KVIEW, null);
            if (view && Array.isArray(view.center) && typeof view.zoom === 'number'){
              try{ map.setView(view.center, view.zoom, {animate:false}); }catch(_){}
            }
          }

          var applied = false;
          var tries = 0;
          var maxTries = 60;
          var i = setInterval(function(){
            tries += 1;
            var ok = document.querySelectorAll('.leaflet-control-layers-overlays label').length > 0;
            if (ok){
              applySavedView();
              var done = applySavedLayers();
              if (done){
                applied = true;
                wireEvents();
                clearInterval(i);
              }
            }
            if (tries >= maxTries && !applied){
              wireEvents();
              clearInterval(i);
            }
          }, 200);

          function wireEvents(){
            if (wireEvents._wired) return;
            wireEvents._wired = true;

            function bindBoxes(){
              getOverlayItems().forEach(it => {
                it.cb.addEventListener('change', saveLayers, {passive:true});
              });
            }
            bindBoxes();

            var obs = new MutationObserver(function(){ bindBoxes(); });
            var root = document.querySelector('.leaflet-control-layers-overlays');
            if (root){ obs.observe(root, {childList:true, subtree:true}); }

            map.on('overlayadd', saveLayers);
            map.on('overlayremove', saveLayers);
            map.on('moveend', saveView);
            map.on('zoomend', saveView);
          }

          window.addEventListener('beforeunload', function(){
            saveLayers();
            saveView();
          });
        })();
        {% endmacro %}
        """)

def add_persist_layer_state(m: folium.Map, storage_key: str):
    m.get_root().add_child(PersistLayerStateControl(storage_key))

# ---------- Card KM ----------
def add_km_card_layer(map_obj: folium.Map, center_latlon: List[float],
                      km_atual: float, fonte_atual: str = "Local",
                      km_otimizado: Optional[float] = None,
                      show: bool = True):
    lat_km_card = center_latlon[0] - 0.0021
    lon_km_card = center_latlon[1] - 0.0077
    linhas = [
        '<p style="margin:6px 0 0 0;"><b style="color:black;">ROTA:</b></p>',
        f'<p style="margin:0; color:black; font-weight:bold;">{km_atual:.2f} km</p>',
        f'<p style="margin:0; font-size:11px; color:#555;">({fonte_atual})</p>',
    ]
    if km_otimizado is not None:
        linhas += [
            '<hr style="margin:4px 16px;border:0;border-top:1px solid #ddd;">',
            f'<p style="margin:0; color:black;"><b>{km_otimizado:.2f} km</b> '
            '<span style="font-size:11px;color:#555;">(Google – otim.)</span></p>'
        ]
    html = f"""
    <div style="
        width: 150px; min-height: 90px; font-size: 14px;
        background-color: white; border: 2px solid grey; border-radius: 6px;
        text-align: center; opacity: 0.92;">
        {''.join(linhas)}
    </div>
    """
    fg = folium.FeatureGroup(name="Resumo da rota (KM)", show=show)
    folium.Marker([lat_km_card, lon_km_card], icon=folium.DivIcon(html=html)).add_to(fg)
    fg.add_to(map_obj)

# ==========================================================
# >>> FUNÇÃO PÚBLICA CHAMADA PELO APP <<<
# ==========================================================
def gerar_mapa_from_path(
    input_path: str,
    criterio: str = "sequence",
    base_name: str = "Mapa base",
    use_google: bool = False,
    google_key: str = "",
    desenhar_rota_google: bool = True,
    mostrar_linha: bool = False,
    mostrar_ida_volta: bool = True,
    mostrar_deposito: bool = True,
    sort_asc: bool = True,
    incluir_deposito_no_km: bool = INCLUIR_DEPOSITO_NO_KM,
    calcular_km_otimizado: bool = CALCULAR_KM_OTIMIZADO,
    start_layers_unchecked: bool = False,
    forcar_layer_google: bool = False,
    show_google_layer: bool = False,
) -> str:

    # --------- TRAVA 1 (no engine): só há Google se houver CHAVE VÁLIDA ---------
    key = (google_key or "").strip()
    key_valid = _is_probably_google_key(key)

    # Se a chave não for válida, zera tudo que poderia consumir API.
    if not key_valid:
        use_google = False
        desenhar_rota_google = False
        forcar_layer_google = False
        show_google_layer = False

    stem = os.path.splitext(os.path.basename(input_path))[0]
    out_html = os.path.join(HTML_DIR, f"{stem}_MAPA_{criterio.upper()}.html")
    cfg_fingerprint = (
        f"dg{int(desenhar_rota_google)}"
        f"-ug{int(bool(use_google and key_valid))}"
        f"-su{int(bool(start_layers_unchecked))}"
        f"-flg{int(bool(forcar_layer_google))}"
        f"-sgl{int(bool(show_google_layer))}"
    )
    storage_key = f"appmaps:{stem}:{criterio.lower()}:{cfg_fingerprint}"

    raw  = read_any(input_path)
    cols = ensure_columns(raw)
    base = clean_base(raw, cols)

    if not cols.get("lat") or not cols.get("lon"):
        raise ValueError("Planilha sem Latitude/Longitude.")
    base = base.dropna(subset=[cols["lat"], cols["lon"]])

    base = ordenar(base, cols, criterio, ascending=sort_asc)
    grp  = agrupar_por_coord_preservando_ordem(base, cols)

    deposito = detectar_deposito_por_sequence(clean_base(raw, cols), cols)

    if criterio.lower() == "melhor" and deposito and len(grp) > 1:
        grp = _ordem_melhor(grp, deposito, key, usar_tempo=(USE_PREDICTIVE_TIME and key_valid))

    PALETA_SEQ = [
        "blue","red","green","black","darkblue","darkred","cadetblue",
        "darkgreen","darkpurple","pink","gray","lightgreen","lightblue","beige","lightgray"
    ]
    ICO_PREF, ICO_NAME, ICO_INT = "fa","home","white"

    def _colors_sequence_like(grp_: pd.DataFrame) -> pd.DataFrame:
        ext = []
        for ordem in grp_["Ordem"]:
            if ordem in (1,2): ext.append("purple")
            else:
                k = max(0, int((ordem - 3) // 15))
                ext.append(PALETA_SEQ[k % len(PALETA_SEQ)])
        return pd.DataFrame({"COR_EXT": ext, "COR_INT": ICO_INT, "PREF": ICO_PREF, "ICO": ICO_NAME})

    def _colors_group_by(grp_: pd.DataFrame, key_: str) -> pd.DataFrame:
        keys = grp_[key_].fillna("").astype(str)
        uniq = list(dict.fromkeys(keys))
        mapa = {v: PALETA_SEQ[i % len(PALETA_SEQ)] for i, v in enumerate(uniq)}
        ext = [mapa[v] for v in keys]
        return pd.DataFrame({"COR_EXT": ext, "COR_INT": ICO_INT, "PREF": ICO_PREF, "ICO": ICO_NAME})

    crit = criterio.lower()
    if crit in ("sequence", "ordem", "melhor"):
        cores_df = _colors_sequence_like(grp)
    elif crit == "cep":
        cores_df = _colors_group_by(grp, "CEP")
    elif crit == "bairro":
        cores_df = _colors_group_by(grp, "Bairro")
    else:
        cores_df = _colors_sequence_like(grp)
    grp = pd.concat([grp.reset_index(drop=True), cores_df.reset_index(drop=True)], axis=1)

    mid = max(0, min(len(grp)-1, len(grp)//2))
    center = [grp.iloc[mid]["Latitude"], grp.iloc[mid]["Longitude"]] if len(grp) else [-22.0, -47.6]

    m = folium.Map(location=center, zoom_start=16, tiles=None, control_scale=True)
    folium.TileLayer(tiles="OpenStreetMap", name=base_name, control=True, show=True).add_to(m)
    Draw().add_to(m); Fullscreen().add_to(m)

    rota_coords = [(float(r["Latitude"]), float(r["Longitude"])) for _, r in grp.iterrows()]

    coords_para_km = rota_coords[:]
    if incluir_deposito_no_km and deposito:
        coords_para_km = [deposito] + coords_para_km
        if mostrar_ida_volta:
            coords_para_km = coords_para_km + [deposito]

    km_total = None
    google_line: List[Tuple[float, float]] = []
    km_otimizado = None

    # ——— Directions apenas com chave válida ———
    if use_google and desenhar_rota_google and key_valid:
        try:
            km_total, google_line = google_directions_roundtrip_chunked(
                coords_para_km, key, idioma=GOOGLE_LANG, optimize=False
            )
            if calcular_km_otimizado:
                km_otimizado, _ = google_directions_roundtrip_chunked(
                    coords_para_km, key, idioma=GOOGLE_LANG, optimize=True
                )
        except Exception as e:
            print(f"Google Directions falhou ({e}). Usando alternativa.")
            km_total = None

    # ——— Fallback Distance Matrix (ainda assim só com chave válida) ———
    if km_total is None and USE_GOOGLE_DISTANCE_MATRIX and key_valid and len(coords_para_km) > 1:
        from collections import defaultdict
        pairs = list(zip(coords_para_km[:-1], coords_para_km[1:]))
        def dist_pairs(pairs_, api_key, language="pt-BR") -> float:
            if not pairs_: return 0.0
            total_km_ = 0.0
            by_origin = defaultdict(list)
            for o, d in pairs_:
                by_origin[o].append(d)
            for origin, dests in by_origin.items():
                o_str = f"{origin[0]},{origin[1]}"
                d_str = "|".join(f"{d[0]},{d[1]}" for d in dests)
                try:
                    r = requests.get("https://maps.googleapis.com/maps/api/distancematrix/json",
                                     params={"origins":o_str,"destinations":d_str,"mode":"driving",
                                             "language":language,"key":api_key}, timeout=60)
                    r.raise_for_status()
                    data = r.json()
                    if data.get("rows"):
                        for el in data["rows"][0]["elements"]:
                            if el.get("status") == "OK":
                                total_km_ += (el["distance"]["value"] / 1000.0)
                except Exception as e:
                    print(f"Distance Matrix falhou ({e}).")
            return round(total_km_, 2)
        km_total = dist_pairs(pairs, key)

    if km_total is None:
        km_total = round(path_len_haversine(coords_para_km), 2)
    if calcular_km_otimizado and km_otimizado is None and (google_line or (use_google and key_valid)):
        km_otimizado = km_total

    fonte = "Google" if google_line else "Local"
    add_km_card_layer(m, center, km_total, fonte, km_otimizado, show=True)

    for _, row in grp.iterrows():
        nome_fg = f"{int(row['Ordem'])}P - {str(row['Sequence']).strip()}"
        fg = folium.FeatureGroup(name=nome_fg, show=True)
        ic = folium.Icon(
            color=str(row.get("COR_EXT") or "blue"),
            icon_color=str(row.get("COR_INT") or "white"),
            icon=str(row.get("ICO") or "home"),
            prefix=str(row.get("PREF") or "fa"),
        )
        folium.Marker(
            location=[row["Latitude"], row["Longitude"]],
            popup=folium.Popup(html=f"""
                <div>
                  <b style="font-size:12px;">Ordem:</b><b style="font-size:14px;"> {int(row['Ordem'])}</b><br>
                  <b style="font-size:12px;">Pedidos:</b> {str(row['Sequence']).strip()}<br>
                  <b style="font-size:12px;">Parada:</b> {row['Parada']}<br>
                  <b style="font-size:12px;">STX:</b> {row['ID']}<br>
                  <b style="font-size:12px;">Nome:</b> {row['INFO']}<br>
                  <b style="font-size:12px;">Destino:</b><b style="font-size:16px;"> {row['LOCAL']}</b><br>
                  <b style="font-size:12px;">CEP:</b> {row['CEP']}<br>
                  <b style="font-size:12px;">Bairro:</b> {row['Bairro']}<br>
                  <b style="font-size:12px;">Cidade:</b> {row['Cidade']}
                  {nav_buttons(float(row['Latitude']), float(row['Longitude']), deposito)}
                </div>
            """, max_width=360),
            icon=ic,
        ).add_to(fg)
        fg.add_to(m)

    if mostrar_linha and len(grp) > 1:
        fg_ln = folium.FeatureGroup(name="Linha simples", show=False)
        folium.PolyLine(grp[["Latitude","Longitude"]].to_numpy().tolist(),
                        weight=3, opacity=0.6).add_to(fg_ln)
        fg_ln.add_to(m)

    if mostrar_ida_volta and deposito and len(rota_coords) > 0:
        ida_coords = [deposito] + rota_coords
        fg_ida = folium.FeatureGroup(name="Rota (ida)", show=False)
        folium.PolyLine(ida_coords, weight=4, opacity=0.8, color="blue").add_to(fg_ida)
        fg_ida.add_to(m)

        volta_coords = rota_coords + [deposito]
        fg_volta = folium.FeatureGroup(name="Rota (volta)", show=False)
        folium.PolyLine(volta_coords, weight=4, opacity=0.9, color="red", dash_array="6,8").add_to(fg_volta)
        fg_volta.add_to(m)

        fg_iv = folium.FeatureGroup(name="Rota (ida/volta)", show=False)
        folium.PolyLine([deposito] + rota_coords + [deposito], weight=4, opacity=0.7).add_to(fg_iv)
        fg_iv.add_to(m)

    if google_line:
        google_fg = folium.FeatureGroup(name="Rota (Google)", show=bool(show_google_layer))
        folium.PolyLine(google_line, weight=5, opacity=0.85, color="blue").add_to(google_fg)
        google_fg.add_to(m)
    elif forcar_layer_google and key_valid:
        # checkbox aparece só se houver possibilidade real de usar Google
        folium.FeatureGroup(name="Rota (Google)", show=False).add_to(m)

    if mostrar_deposito and deposito:
        depo_fg = folium.FeatureGroup(name="Depósito", show=True)
        folium.Marker(
            location=[deposito[0], deposito[1]],
            popup="Depósito (ponto zero)",
            icon=folium.Icon(color="orange", icon="flag", prefix="fa")
        ).add_to(depo_fg)
        depo_fg.add_to(m)

    folium.LayerControl(collapsed=True).add_to(m)
    add_uncheck_all_button(m)
    add_persist_layer_state(m, storage_key)

    m.save(out_html)
    print(f"✔ Mapa gerado com persistência em: {out_html}")
    return out_html


if __name__ == "__main__":
    # Execução direta: SEM chave; não deve consumir Google.
    gerar_mapa_from_path(
        INPUT_PATH,
        criterio=CRITERIO,
        base_name=BASE_NAME,
        use_google=False,          # segurança
        google_key="",             # segurança
        desenhar_rota_google=False,
        mostrar_linha=MOSTRAR_LINHA,
        mostrar_ida_volta=MOSTRAR_ROTA_IDA_VOLTA,
        mostrar_deposito=MOSTRAR_PIN_DEPOSITO,
        sort_asc=True,
        incluir_deposito_no_km=INCLUIR_DEPOSITO_NO_KM,
        calcular_km_otimizado=CALCULAR_KM_OTIMIZADO,
        forcar_layer_google=False,
        show_google_layer=False,
    )

