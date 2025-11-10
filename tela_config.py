# -*- coding: utf-8 -*-
import json, os
from pathlib import Path
from kivymd.uix.screen import MDScreen
from kivy.properties import StringProperty, NumericProperty, BooleanProperty
from kivymd.app import MDApp


# ---------- util: localizar raiz que contém ui/ ----------
def _locate_root_with_ui(start: Path, max_hops: int = 6) -> Path | None:
    """
    Sobe até max_hops níveis a partir de 'start' procurando um diretório que contenha 'ui'.
    Retorna o Path encontrado ou None.
    """
    cur = start
    hops = 0
    while cur and hops <= max_hops:
        if (cur / "ui").exists() and (cur / "ui").is_dir():
            return cur
        cur = cur.parent
        hops += 1
    return None


class TelaConfig(MDScreen):
    """
    Tela de Configurações.
    - Lê/salva config.json (prioriza ui/config.json; mantém compat. com ./config.json)
    - Exibe e atualiza switches/campos
    - Testa a chave Google (desktop)
    """

    # ---------- Caminhos base (agora corretos) ----------
    # 1) tenta descobrir automaticamente a raiz que contém 'ui/'
    _THIS_FILE = Path(__file__).resolve()
    _AUTO_ROOT = _locate_root_with_ui(_THIS_FILE.parent)

    # 2) fallback EXATO que você informou
    _FALLBACK_ROOT = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps")

    # 3) raiz final
    APP_ROOT = _AUTO_ROOT if _AUTO_ROOT is not None else _FALLBACK_ROOT

    UI_DIR = APP_ROOT / "ui"
    CONFIG_UI = UI_DIR / "config.json"          # padrão principal
    CONFIG_ROOT = APP_ROOT / "config.json"      # compatibilidade

    # ---------- espelhos ----------
    api_key = StringProperty("")
    lang = StringProperty("pt-BR")
    max_points = NumericProperty(25)
    use_google = BooleanProperty(True)
    desenhar_rota_google = BooleanProperty(True)
    incluir_deposito_no_km = BooleanProperty(True)
    calcular_km_otimizado = BooleanProperty(True)
    start_layers_unchecked = BooleanProperty(False)
    use_predictive_time = BooleanProperty(False)

    DEFAULTS = {
        "GOOGLE_API_KEY": "",
        "GOOGLE_LANG": "pt-BR",
        "GOOGLE_MAX_POINTS": 25,
        "USE_GOOGLE": True,
        "DESENHAR_ROTA_GOOGLE": True,
        "INCLUIR_DEPOSITO_NO_KM": True,
        "CALCULAR_KM_OTIMIZADO": True,
        "START_LAYERS_UNCHECKED": False,
        "USE_PREDICTIVE_TIME": False
    }

    # ---------- ciclo de vida ----------
    def on_pre_enter(self, *args):
        self._load_into_ui()

    # ---------- notificador ----------
    def _notify(self, msg: str, title: str = "Info"):
        try:
            app = MDApp.get_running_app()
            if hasattr(app, "show_dialog"):
                app.show_dialog(title, msg)
            else:
                print(f"[{title}] {msg}")
        except Exception:
            print(f"[{title}] {msg}")

    # ---------- leitura/escrita de config ----------
    def _resolve_read_path(self) -> Path:
        """
        Preferência de leitura:
        1) APP_ROOT/ui/config.json
        2) APP_ROOT/config.json
        3) APP_ROOT/ui/config.json (para futura criação)
        """
        if self.CONFIG_UI.exists():
            return self.CONFIG_UI
        if self.CONFIG_ROOT.exists():
            return self.CONFIG_ROOT
        return self.CONFIG_UI

    def _read_config(self) -> dict:
        path = self._resolve_read_path()
        cfg = self.DEFAULTS.copy()
        try:
            if path.exists():
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f) or {}
                for k in cfg:
                    if k in data:
                        cfg[k] = data[k]
        except Exception as e:
            print(f"[TelaConfig] Falha ao ler config: {e}")
        return cfg

    def _safe_write_json(self, dest: Path, data: dict):
        """
        Escrita segura (arquivo temporário + replace) para evitar corrupção.
        """
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_suffix(dest.suffix + ".tmp")
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, dest)

    def _write_config(self, cfg: dict) -> bool:
        """
        Grava em ui/config.json e em ./config.json (ambos dentro de
        C:\ProjetoApp\Maps\AppMaps\AppMaps\... )
        """
        try:
            self._safe_write_json(self.CONFIG_UI, cfg)
            self._safe_write_json(self.CONFIG_ROOT, cfg)
            print(f"[TelaConfig] Config salvas em:\n  {self.CONFIG_UI}\n  {self.CONFIG_ROOT}")
            return True
        except Exception as e:
            self._notify(f"Falha ao salvar: {e}", "Erro")
            return False

    # ---------- sincronização UI <-> JSON ----------
    def _load_into_ui(self):
        cfg = self._read_config()

        # espelhos
        self.api_key = cfg["GOOGLE_API_KEY"]
        self.lang = cfg["GOOGLE_LANG"]
        self.max_points = cfg["GOOGLE_MAX_POINTS"]
        self.use_google = cfg["USE_GOOGLE"]
        self.desenhar_rota_google = cfg["DESENHAR_ROTA_GOOGLE"]
        self.incluir_deposito_no_km = cfg["INCLUIR_DEPOSITO_NO_KM"]
        self.calcular_km_otimizado = cfg["CALCULAR_KM_OTIMIZADO"]
        self.start_layers_unchecked = cfg["START_LAYERS_UNCHECKED"]
        self.use_predictive_time = cfg["USE_PREDICTIVE_TIME"]

        ids = self.ids
        try:
            if "tf_api_key" in ids:      ids["tf_api_key"].text = self.api_key
            if "tf_lang" in ids:         ids["tf_lang"].text = self.lang
            if "tf_max_pts" in ids:      ids["tf_max_pts"].text = str(self.max_points)
            if "sw_use_google" in ids:   ids["sw_use_google"].active = self.use_google
            if "sw_draw_google" in ids:  ids["sw_draw_google"].active = self.desenhar_rota_google
            if "sw_incluir_dep" in ids:  ids["sw_incluir_dep"].active = self.incluir_deposito_no_km
            if "sw_km_otim" in ids:      ids["sw_km_otim"].active = self.calcular_km_otimizado
            if "sw_layers_unchecked" in ids: ids["sw_layers_unchecked"].active = self.start_layers_unchecked
            if "sw_predictive" in ids:   ids["sw_predictive"].active = self.use_predictive_time
        except Exception as e:
            print(f"[TelaConfig] Falha ao sincronizar widgets: {e}")

    # ---------- ações ----------
    def salvar_config(self):
        ids = self.ids

        def _int_from(id_key, default):
            try:
                return max(2, int(ids[id_key].text)) if id_key in ids else default
            except Exception:
                return default

        def _sw(id_key, default):
            try:
                return bool(ids[id_key].active) if id_key in ids else default
            except Exception:
                return default

        cfg = {
            "GOOGLE_API_KEY": (ids["tf_api_key"].text if "tf_api_key" in ids else "").strip(),
            "GOOGLE_LANG": (ids["tf_lang"].text if "tf_lang" in ids else "").strip() or "pt-BR",
            "GOOGLE_MAX_POINTS": _int_from("tf_max_pts", 25),

            "USE_GOOGLE": _sw("sw_use_google", True),
            "DESENHAR_ROTA_GOOGLE": _sw("sw_draw_google", True),
            "INCLUIR_DEPOSITO_NO_KM": _sw("sw_incluir_dep", True),
            "CALCULAR_KM_OTIMIZADO": _sw("sw_km_otim", True),
            "START_LAYERS_UNCHECKED": _sw("sw_layers_unchecked", False),
            "USE_PREDICTIVE_TIME": _sw("sw_predictive", False),
        }

        if self._write_config(cfg):
            # espelhos
            self.api_key = cfg["GOOGLE_API_KEY"]
            self.lang = cfg["GOOGLE_LANG"]
            self.max_points = cfg["GOOGLE_MAX_POINTS"]
            self.use_google = cfg["USE_GOOGLE"]
            self.desenhar_rota_google = cfg["DESENHAR_ROTA_GOOGLE"]
            self.incluir_deposito_no_km = cfg["INCLUIR_DEPOSITO_NO_KM"]
            self.calcular_km_otimizado = cfg["CALCULAR_KM_OTIMIZADO"]
            self.start_layers_unchecked = cfg["START_LAYERS_UNCHECKED"]
            self.use_predictive_time = cfg["USE_PREDICTIVE_TIME"]

            # avisa app principal (se existir)
            try:
                app = MDApp.get_running_app()
                if hasattr(app, "atualizar_cfg"):
                    app.atualizar_cfg(cfg)
            except Exception:
                pass

            self._notify(
                f"Configurações salvas em:\n{self.CONFIG_UI}\n{self.CONFIG_ROOT}",
                "Sucesso"
            )

    def restaurar_padrao(self):
        if self._write_config(self.DEFAULTS.copy()):
            self._notify("Padrões restaurados.", "Sucesso")
            self._load_into_ui()

    # ---------- teste de chave Google ----------
    def testar_chave_google(self):
        try:
            import googlemaps
        except Exception:
            self._notify("Pacote 'googlemaps' não instalado.", "Erro")
            return

        key = (self.ids["tf_api_key"].text if "tf_api_key" in self.ids else "").strip()
        if not key:
            self._notify("Informe a GOOGLE_API_KEY antes de testar.", "Aviso")
            return

        try:
            gmaps = googlemaps.Client(key=key)
        except Exception as e:
            self._notify(f"Falha ao iniciar cliente Google: {e}", "Erro")
            return

        ok, fail = [], []

        try:
            r = gmaps.geocode("Av. Paulista, 1000 - São Paulo")
            (ok if r else fail).append("Geocoding" if r else "Geocoding (vazio)")
        except Exception as e:
            fail.append(f"Geocoding ({e})")

        try:
            r = gmaps.directions("Av. Paulista, 1000 - São Paulo", "Praça da Sé, São Paulo", mode="driving")
            (ok if r else fail).append("Directions" if r else "Directions (vazio)")
        except Exception as e:
            fail.append(f"Directions ({e})")

        try:
            r = gmaps.distance_matrix(["Av. Paulista, 1000 - São Paulo"], ["Praça da Sé, São Paulo"], mode="driving")
            valid = bool(r and r.get("rows"))
            (ok if valid else fail).append("DistanceMatrix" if valid else "DistanceMatrix (vazio)")
        except Exception as e:
            fail.append(f"DistanceMatrix ({e})")

        msg = "✅ OK: " + ", ".join(ok) if ok else "Nenhuma OK."
        if fail:
            msg += "\n⚠️ Falhas: " + "; ".join(fail)
        self._notify(msg, "Teste Google")




