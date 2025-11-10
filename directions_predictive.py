# -*- coding: utf-8 -*-
"""
Ordenação 'melhor' usando TEMPO PREVISTO (trânsito) via Google Distance Matrix.

- Não altera as coordenadas dos pontos (NÃO faz geocoding e NÃO faz snap-to-roads).
- Usa duration_in_traffic quando disponível (departure_time=now, traffic_model=best_guess).
- Mantém cache local (CSV) para reduzir chamadas à API.
- Heurística: Vizinho Mais Próximo por tempo + 2-OPT por tempo (round-trip).

Dep.: requests, pandas (somente para conveniência no dataframe de entrada), math.
"""
from __future__ import annotations
import csv
import os
import time
from math import radians, sin, cos, sqrt, atan2
from typing import List, Tuple, Optional, Dict

import requests
import pandas as pd

# velocidade de fallback (quando não há API): 35 km/h (valor conservador urbano)
FALLBACK_SPEED_KMH = 35.0

def _haversine_km(a: Tuple[float, float], b: Tuple[float, float]) -> float:
    (lat1, lon1), (lat2, lon2) = a, b
    R = 6371.0088
    p1, p2 = radians(lat1), radians(lat2)
    dphi = radians(lat2 - lat1); dl = radians(lon2 - lon1)
    h = sin(dphi/2)**2 + cos(p1)*cos(p2)*sin(dl/2)**2
    return 2*R*atan2(sqrt(h), sqrt(1-h))

def _fallback_secs(a: Tuple[float, float], b: Tuple[float, float]) -> int:
    km = _haversine_km(a, b)
    hours = km / max(FALLBACK_SPEED_KMH, 1e-6)
    return int(round(hours * 3600))

def _key_pair(a: Tuple[float, float], b: Tuple[float, float], prec: int = 5) -> Tuple[str, str]:
    ar = f"{round(a[0],prec)},{round(a[1],prec)}"
    br = f"{round(b[0],prec)},{round(b[1],prec)}"
    return ar, br

def _cache_load(cache_path: Optional[str]) -> Dict[Tuple[str,str], int]:
    d: Dict[Tuple[str,str], int] = {}
    if cache_path and os.path.exists(cache_path):
        with open(cache_path, "r", newline="", encoding="utf-8") as f:
            rd = csv.DictReader(f)
            for r in rd:
                k = (r["o"], r["d"])
                try:
                    d[k] = int(float(r["secs"]))
                except Exception:
                    pass
    return d

def _cache_save(cache_path: Optional[str], store: Dict[Tuple[str,str], int]) -> None:
    if not cache_path:
        return
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    with open(cache_path, "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, fieldnames=["o","d","secs","ts"])
        wr.writeheader()
        ts = int(time.time())
        for (o,d), secs in store.items():
            wr.writerow({"o":o, "d":d, "secs":secs, "ts":ts})

def dm_time_seconds(origins: List[Tuple[float,float]],
                    destinations: List[Tuple[float,float]],
                    api_key: str,
                    timeout: int = 60) -> List[List[int]]:
    """
    Retorna matriz de TEMPO (segundos) entre 'origins' e 'destinations' usando
    Distance Matrix (driving, duration_in_traffic quando disponível).
    """
    if not origins or not destinations:
        return []

    ori = "|".join(f"{lat},{lon}" for lat,lon in origins)
    dst = "|".join(f"{lat},{lon}" for lat,lon in destinations)

    url = "https://maps.googleapis.com/maps/api/distancematrix/json"
    params = {
        "origins": ori,
        "destinations": dst,
        "mode": "driving",
        "departure_time": "now",        # habilita duration_in_traffic
        "traffic_model": "best_guess",
        "language": "pt-BR",
        "key": api_key
    }
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    data = r.json()

    rows = data.get("rows", [])
    out: List[List[int]] = []
    for i, row in enumerate(rows):
        line = []
        elems = row.get("elements", [])
        for el in elems:
            if el.get("status") == "OK":
                # prioriza traffic; se não existir, usa duration
                dur = el.get("duration_in_traffic") or el.get("duration") or {}
                secs = int(dur.get("value", 0))
                line.append(max(1, secs))
            else:
                # se não vier OK, preenche com 0 (trataremos depois)
                line.append(0)
        out.append(line)
    return out

def _time_matrix(points: List[Tuple[float,float]],
                 api_key: Optional[str],
                 cache_path: Optional[str]) -> List[List[int]]:
    """
    Gera matriz NxN de tempo (s). Usa cache por par. Se api_key for None/'' usa fallback.
    """
    n = len(points)
    M = [[0]*n for _ in range(n)]
    cache = _cache_load(cache_path)

    # primeiro tenta preencher via cache
    miss_pairs = []
    for i in range(n):
        for j in range(n):
            if i == j:
                M[i][j] = 0
                continue
            k = _key_pair(points[i], points[j])
            if k in cache:
                M[i][j] = cache[k]
            else:
                miss_pairs.append((i,j))

    if api_key and miss_pairs:
        # para reduzir chamadas, agrupamos por origem
        by_origin: Dict[int, List[int]] = {}
        for i,j in miss_pairs:
            by_origin.setdefault(i, []).append(j)

        for i, js in by_origin.items():
            origins = [points[i]]
            destinations = [points[j] for j in js]
            mat = dm_time_seconds(origins, destinations, api_key)
            if mat and len(mat[0]) == len(js):
                for col, j in enumerate(js):
                    secs = int(mat[0][col] or 0)
                    if secs <= 0:
                        secs = _fallback_secs(points[i], points[j])
                    M[i][j] = secs
                    cache[_key_pair(points[i], points[j])] = secs
            else:
                # fallback total para este origin
                for j in js:
                    secs = _fallback_secs(points[i], points[j])
                    M[i][j] = secs
                    cache[_key_pair(points[i], points[j])] = secs

        _cache_save(cache_path, cache)
    else:
        # não há API: tudo por fallback
        for i in range(n):
            for j in range(n):
                if i != j and M[i][j] == 0:
                    M[i][j] = _fallback_secs(points[i], points[j])

    return M

def _nearest_neighbor_by_time(M: List[List[int]], start: int) -> List[int]:
    n = len(M)
    unvis = set(range(n))
    cur = start
    order = [cur]
    unvis.remove(cur)
    while unvis:
        nxt = min(unvis, key=lambda j: M[cur][j])
        order.append(nxt)
        unvis.remove(nxt)
        cur = nxt
    return order

def _two_opt_by_time(order: List[int], M: List[List[int]]) -> List[int]:
    best = order[:]
    improved = True
    def seg(a,b): return M[a][b]
    while improved:
        improved = False
        for i in range(1, len(best)-2):
            for k in range(i+1, len(best)-1):
                a,b = best[i-1], best[i]
                c,d = best[k], best[k+1]
                delta = (seg(a,c)+seg(b,d)) - (seg(a,b)+seg(c,d))
                if delta < 0:
                    best[i:k+1] = reversed(best[i:k+1])
                    improved = True
    return best

def order_by_predictive_time(grp: pd.DataFrame,
                             deposito: Tuple[float,float],
                             api_key: Optional[str],
                             cache_path: Optional[str]) -> pd.DataFrame:
    """
    Recebe o DataFrame das paradas (colunas Latitude/Longitude).
    Retorna o mesmo DF reordenado (coluna Ordem 1..N) usando TEMPO como custo.
    """
    pts = [(float(r["Latitude"]), float(r["Longitude"])) for _, r in grp.iterrows()]
    if len(pts) <= 2:
        out = grp.reset_index(drop=True).copy()
        out["Ordem"] = range(1, len(out)+1)
        return out

    # índice inicial: ponto mais rápido (menor tempo) a partir do depósito
    M0 = _time_matrix([deposito] + pts, api_key, cache_path)  # (N+1)x(N+1)
    # tempos do depósito (linha 0) para cada parada (1..N)
    times_from_depot = [M0[0][i+1] for i in range(len(pts))]
    start = int(min(range(len(pts)), key=lambda i: times_from_depot[i]))

    # matriz só entre paradas:
    M = _time_matrix(pts, api_key, cache_path)

    nn = _nearest_neighbor_by_time(M, start)
    better = _two_opt_by_time(nn, M)

    out = grp.iloc[better].reset_index(drop=True).copy()
    out["Ordem"] = range(1, len(out)+1)
    return out
