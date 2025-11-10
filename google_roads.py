# -*- coding: utf-8 -*-
"""
google_roads.py — Snap to Roads para o AppMaps.

Uso:
    from google_roads import snap_points_with_roads, snap_dataframe_with_roads

- snap_points_with_roads: recebe lista de (lat, lon) e devolve lista "grudada" na via
- snap_dataframe_with_roads: recebe DataFrame com colunas Latitude/Longitude e devolve DF ajustado

Requisitos:
    - Roads API habilitada no Google Cloud
    - GOOGLE_API_KEY válido (mesmo do Directions)
"""

from __future__ import annotations
import time
import requests
from typing import List, Tuple, Optional
import pandas as pd

ROADS_ENDPOINT = "https://roads.googleapis.com/v1/snapToRoads"

def _http_get_json(url: str, params: dict, timeout: int = 60) -> dict:
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.json()

def snap_points_with_roads(points: List[Tuple[float, float]],
                           api_key: str,
                           interpolate: bool = True,
                           batch_size: int = 100,
                           sleep_between: float = 0.05) -> List[Tuple[float, float]]:
    """
    Aplica Snap to Roads a uma lista de pontos (lat, lon).
    - Roads API aceita até 100 pontos por chamada (path=lat,lon|...)
    - Se interpolate=True, o Google pode inserir pontos intermediários entre os de entrada
      (nós retornamos apenas os pontos "snapped" na mesma cadência de entrada).

    Retorna: lista de (lat, lon) ajustados. Se algum bloco falhar, retorna os originais daquele bloco.
    """
    if not points:
        return []

    snapped: List[Tuple[float, float]] = []
    n = len(points)
    i = 0
    while i < n:
        j = min(i + batch_size, n)
        chunk = points[i:j]
        path = "|".join(f"{lat},{lon}" for lat, lon in chunk)
        try:
            data = _http_get_json(
                ROADS_ENDPOINT,
                {"path": path, "interpolate": str(interpolate).lower(), "key": api_key},
                timeout=60
            )
            # A resposta vem em "snappedPoints". Precisamos reconstruir na mesma cadência de entrada.
            # O Google devolve points com placeId/locations. Vamos mapear por "originalIndex" quando existir;
            # se interpolado, não terá originalIndex (pulamos).
            idx_to_latlon = {}
            for sp in data.get("snappedPoints", []):
                if "originalIndex" in sp:
                    idx = int(sp["originalIndex"])
                    loc = sp.get("location", {})
                    lat = float(loc.get("latitude"))
                    lon = float(loc.get("longitude"))
                    idx_to_latlon[idx] = (lat, lon)
            # Remonta o bloco na ordem original:
            for k in range(len(chunk)):
                snapped.append(idx_to_latlon.get(k, chunk[k]))
        except Exception as e:
            print(f"[ROADS] falha ao snappar bloco {i}:{j} — {e}. Mantendo coordenadas originais do bloco.")
            snapped.extend(chunk)
        i = j
        if i < n and sleep_between:
            time.sleep(sleep_between)
    return snapped

def snap_dataframe_with_roads(df: pd.DataFrame,
                              lat_col: str,
                              lon_col: str,
                              api_key: str,
                              interpolate: bool = True) -> pd.DataFrame:
    """
    Recebe um DataFrame com lat/lon e devolve outro com as colunas ajustadas (snap to roads).
    """
    base = df.copy()
    pts = list(zip(base[lat_col].astype(float), base[lon_col].astype(float)))
    pts_snapped = snap_points_with_roads(pts, api_key=api_key, interpolate=interpolate)
    base[lat_col] = [p[0] for p in pts_snapped]
    base[lon_col] = [p[1] for p in pts_snapped]
    return base
