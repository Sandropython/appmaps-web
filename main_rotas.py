# -*- coding: utf-8 -*-
import os
import json
import threading
import inspect
from pathlib import Path

from kivy.lang import Builder
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.properties import StringProperty, BooleanProperty
from kivy.graphics import Color, Rectangle, InstructionGroup

from kivymd.app import MDApp
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.filemanager import MDFileManager

# ---------------- Fallback p/ DropdownMenu legados (viewclass) ----------------
MENU_FALLBACK_CLASS_NAME = None
try:
    from kivymd.uix.list import OneLineListItem  # noqa: F401
    MENU_FALLBACK_CLASS_NAME = "OneLineListItem"
except Exception:
    try:
        from kivymd.uix.list import MDListItem, MDListItemHeadlineText
        from kivy.factory import Factory

        class _MenuItemCompat(MDListItem):
            def __init__(self, **kwargs):
                text = kwargs.pop("text", "")
                on_release_cb = kwargs.pop("on_release", None)
                super().__init__(**kwargs)
                self.add_widget(MDListItemHeadlineText(text=text))
                if on_release_cb:
                    self.bind(on_release=lambda *a: on_release_cb())

        Factory.register("_MenuItemCompat", cls=_MenuItemCompat)
        MENU_FALLBACK_CLASS_NAME = "_MenuItemCompat"
    except Exception:
        MENU_FALLBACK_CLASS_NAME = None

# ---------------- Janela/paths ----------------
APP_DIR = os.path.dirname(__file__)
UI_DIR = os.path.join(APP_DIR, "ui")
KV_FILE = os.path.join(UI_DIR, "main_rotas.kv")

Window.size = (360, 640)

for _icon in ("app_icon.ico", "app_icon.png", "icon.ico", "icon.png", "LogoFinal.png"):
    _p = os.path.join(UI_DIR, _icon)
    if os.path.exists(_p):
        try:
            Window.set_icon(_p)
            break
        except Exception:
            pass

ALLOWED_EXTS = {".xlsx", ".xls", ".csv", ".pkl"}
LAST_DIR_FILE = os.path.join(str(Path.home()), ".rotas_last_dir.txt")
DEFAULT_SEED_DIRS = [
    r"C:\Users\sandr\OneDrive\2025\Shopee\Rota",
    str(Path.home()),
]

# ---------------- Config ----------------
DEFAULT_CFG = {
    "GOOGLE_API_KEY": "",
    "GOOGLE_LANG": "pt-BR",
    "GOOGLE_MAX_POINTS": 25,
    "USE_GOOGLE": True,
    "DESENHAR_ROTA_GOOGLE": True,
    "INCLUIR_DEPOSITO_NO_KM": False,
    "CALCULAR_KM_OTIMIZADO": True,
    "START_LAYERS_UNCHECKED": False,
    "USE_PREDICTIVE_TIME": False,
}

def _load_cfg():
    """Carrega config.json (APP_DIR ou UI_DIR), mescla com DEFAULT_CFG e tolera arquivo ausente/corrompido."""
    cfg = DEFAULT_CFG.copy()
    for base in (APP_DIR, UI_DIR):
        p = os.path.join(base, "config.json")
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f) or {}
                for k in cfg:
                    if k in data:
                        cfg[k] = data[k]
                break
            except Exception:
                pass
    return cfg

# ---------------- Snackbar/print “seguro” ----------------
def _toast_safe(msg: str):
    try:
        from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
        bar = MDSnackbar(MDSnackbarText(text=str(msg)), y=dp(16), pos_hint={"center_x": 0.5})
        bar.open()
    except Exception:
        print(f"[TOAST] {msg}")

# ---------------- Validação leve de chave ----------------
def _is_probably_google_key(k: str) -> bool:
    """Validação sem custo: chaves do Google costumam começar com 'AIza' e ter ~39-45 chars."""
    if not k or not isinstance(k, str):
        return False
    k = k.strip()
    if not k:
        return False
    return k.startswith("AIza") and (30 <= len(k) <= 60)

# ---------------- Key Management (NÃO expor a chave) ----------------
def _load_google_key_safe():
    """
    Busca a chave do Google SOMENTE de fontes controladas pelo usuário:
      1) ENV GOOGLE_API_KEY
      2) ui/config.json ou ./config.json
    (NÃO usamos mais constantes em código para evitar “vazar” chave ao distribuir executável)
    """
    k = os.getenv("GOOGLE_API_KEY", "").strip()
    if k:
        return k

    for base in (UI_DIR, APP_DIR):
        cfg = os.path.join(base, "config.json")
        if os.path.exists(cfg):
            try:
                with open(cfg, "r", encoding="utf-8") as f:
                    data = json.load(f)
                k = str(data.get("GOOGLE_API_KEY", "")).strip()
                if k:
                    return k
            except Exception:
                pass
    return ""

class RotasApp(MDApp):
    input_path = StringProperty("")
    criterio_atual = StringProperty("sequence")  # sequence | bairro | cep | melhor
    sort_asc = BooleanProperty(True)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.file_manager: MDFileManager | None = None
        self.menu_criterio: MDDropdownMenu | None = None
        self.cfg = _load_cfg()  # carga inicial

    def build(self):
        if not os.path.exists(KV_FILE):
            raise FileNotFoundError(f"KV não encontrado: {KV_FILE}")
        root = Builder.load_file(KV_FILE)
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.theme_style = "Light"
        self.logo_small = os.path.join(UI_DIR, "LogoFinal.png")
        return root

    @property
    def ids(self):
        return self.root.ids

    # ---------------- Helpers simples de UI ----------------
    def _status(self, msg: str):
        Clock.schedule_once(lambda *_: self._set_lbl("lbl_status", f"Status: {msg}"), 0)

    def _set_lbl(self, lbl_id: str, text: str):
        try:
            self.ids[lbl_id].text = text
        except Exception:
            pass

    def _set_file_label(self, path: str):
        Clock.schedule_once(lambda *_: self._set_lbl("file_label", f"Arquivo: {path}"), 0)

    def _enable_gerar(self, enable: bool):
        Clock.schedule_once(lambda *_: setattr(self.ids["btn_gerar"], "disabled", not enable), 0)

    # ---------------- Navegação ----------------
    def voltar_home(self):
        try:
            from kivy.uix.screenmanager import ScreenManager
            if hasattr(self.root, "manager") and isinstance(self.root.manager, ScreenManager):
                sm = self.root.manager
                if "home" in sm.screen_names:
                    sm.transition.direction = "right"
                    sm.current = "home"
                    return
                elif sm.screen_names:
                    sm.transition.direction = "right"
                    sm.current = sm.screen_names[0]
                    return
        except Exception:
            pass
        try:
            self.stop()
        except Exception:
            pass

    def mostrar_info_criterio(self):
        _toast_safe("Critérios: sequence, bairro, cep ou melhor (TSP).")

    # ---------------- lembrar diretório ----------------
    def _get_last_dir(self) -> str:
        try:
            if os.path.exists(LAST_DIR_FILE):
                with open(LAST_DIR_FILE, "r", encoding="utf-8") as f:
                    d = f.read().strip()
                    if d and os.path.isdir(d):
                        return d
        except Exception:
            pass
        for d in DEFAULT_SEED_DIRS:
            if os.path.isdir(d):
                return d
        return str(Path.home())

    def _set_last_dir(self, d: str):
        try:
            with open(LAST_DIR_FILE, "w", encoding="utf-8") as f:
                f.write(d)
        except Exception:
            pass

    # =====================================================================
    # FileManager – abrir e ESTILIZAR
    # =====================================================================
    def abrir_file_manager(self):
        initial_path = self._get_last_dir()
        self._status(f"Abrindo explorador em: {initial_path}")

        if self.file_manager is None:
            self.file_manager = MDFileManager(
                exit_manager=self._close_file_manager,
                select_path=self._select_file,
                preview=False,
                search="all",
                use_access=True,
            )

        self.file_manager.show(initial_path)
        self._fm_style_tries = 0
        Clock.schedule_interval(self._tentar_tweak_file_manager, 0.1)

    def _tentar_tweak_file_manager(self, dt):
        fm = getattr(self, "file_manager", None)
        if not fm or not fm.parent:
            return False

        self._fm_style_tries += 1
        if self._fm_style_tries > 60:
            print("[FileManager] Não consegui aplicar estilo (timeout).")
            return False

        tweaked = False
        self._pintar_fundo_widget(fm, (0.97, 0.97, 1.00, 1))
        tweaked = True

        tb = self._find_toolbar(fm)
        if not tb and hasattr(fm, "ids"):
            tb = fm.ids.get("toolbar") or fm.ids.get("appbar") or fm.ids.get("top_app_bar")

        if tb:
            if hasattr(tb, "theme_bg_color"):
                try:
                    tb.theme_bg_color = "Custom"
                except Exception:
                    pass
            if hasattr(tb, "md_bg_color"):
                tb.md_bg_color = (0.13, 0.59, 0.95, 1)
            elif hasattr(tb, "bg_color"):
                tb.bg_color = (0.13, 0.59, 0.95, 1)

            for attr in ("specific_text_color", "title_color"):
                if hasattr(tb, attr):
                    setattr(tb, attr, (1, 1, 1, 1))
                    break

            if hasattr(tb, "anchor_title"):
                tb.anchor_title = "left"
            if hasattr(tb, "type"):
                tb.type = "small"
            if hasattr(tb, "elevation"):
                try:
                    tb.elevation = 0
                except Exception:
                    pass

            cur = getattr(fm, "current_path", "") or ""
            if hasattr(tb, "title"):
                tb.title = self._shorten_path(cur, 42)
            self._pintar_fundo_widget(tb, (0.13, 0.59, 0.95, 1))
            tweaked = True

        path_label = None
        try:
            if hasattr(fm, "ids"):
                path_label = fm.ids.get("path") or fm.ids.get("current_path") or fm.ids.get("label")
            if not path_label:
                for w in fm.walk(restrict=False):
                    txt = getattr(w, "text", "")
                    if isinstance(txt, str) and ("\\" in txt or "/" in txt):
                        path_label = w
                        break
        except Exception:
            path_label = None

        if path_label:
            self._apply_font_style(path_label, ["bodySmall", "body-medium", "Body2", "Caption", "Body1"])
            if hasattr(path_label, "theme_text_color"):
                path_label.theme_text_color = "Custom"
            if hasattr(path_label, "text_color"):
                path_label.text_color = (1, 1, 1, 1)
            if hasattr(path_label, "shorten"):
                path_label.shorten = True
                path_label.shorten_from = "left"
            tweaked = True

        ok_btn = self._find_fab(fm)
        if ok_btn:
            if hasattr(ok_btn, "icon"):
                ok_btn.icon = "check"
            if hasattr(ok_btn, "md_bg_color"):
                ok_btn.md_bg_color = (0.10, 0.40, 0.80, 1)
            if hasattr(ok_btn, "elevation"):
                ok_btn.elevation = 2
            tweaked = True

        if tweaked:
            print("[FileManager] Estilo aplicado.")
            return False
        return True

    # ---- helpers de estilo ----
    def _apply_font_style(self, label, candidates):
        try:
            theme = getattr(self, "theme_cls", None)
            valid = set(getattr(theme, "font_styles", [])) if theme else set()
        except Exception:
            valid = set()
        for name in candidates:
            try:
                if name in valid:
                    label.font_style = name
                    return True
            except Exception:
                pass
        try:
            label.font_size = sp(12)
        except Exception:
            pass
        return False

    def _pintar_fundo_widget(self, widget, rgba):
        """Pinta fundo sem clear() – evita crash RenderContext."""
        bg_group = getattr(widget, "_bg_group", None)
        if bg_group is not None:
            try:
                widget.canvas.before.remove(bg_group)
            except Exception:
                pass

        bg_group = InstructionGroup()
        color = Color(*rgba)
        rect = Rectangle(pos=widget.pos, size=widget.size)
        bg_group.add(color)
        bg_group.add(rect)
        widget.canvas.before.add(bg_group)
        widget._bg_group = bg_group

        def _sync(*_):
            rect.pos = widget.pos
            rect.size = widget.size

        widget.bind(pos=_sync, size=_sync)

    def _find_toolbar(self, root_widget):
        for w in root_widget.walk(restrict=False):
            if w.__class__.__name__ in ("MDTopAppBar", "MDToolbar"):
                return w
        return None

    def _find_fab(self, root_widget):
        for w in root_widget.walk(restrict=False):
            if w.__class__.__name__ in ("MDFloatingActionButton", "MDFabButton"):
                return w
        return None

    def _shorten_path(self, path: str, max_chars: int = 42) -> str:
        if not path or len(path) <= max_chars:
            return path
        return "…" + path[-(max_chars - 1):]

    def _close_file_manager(self, *args):
        if self.file_manager:
            self.file_manager.close()

    def _select_file(self, path: str):
        if os.path.isdir(path):
            self._set_last_dir(path)
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in ALLOWED_EXTS:
            _toast_safe("Escolha um arquivo .xlsx, .xls, .csv ou .pkl")
            return
        self.input_path = path
        self._set_last_dir(os.path.dirname(path))
        self._set_file_label(path)
        self._status("Arquivo selecionado. Pronto para gerar.")
        self._enable_gerar(True)
        self._close_file_manager()

    # ---------------- Menu Critério ----------------
    def _menu_usa_viewclass(self) -> bool:
        try:
            sig = inspect.signature(MDDropdownMenu.__init__)
            return "viewclass" in sig.parameters
        except Exception:
            return False

    def abrir_menu_criterio(self, caller_btn):
        labels = ("sequence", "bairro", "cep", "melhor")
        usa_viewclass = self._menu_usa_viewclass()

        if usa_viewclass and MENU_FALLBACK_CLASS_NAME:
            itens = [{
                "viewclass": MENU_FALLBACK_CLASS_NAME,
                "text": lbl,
                "on_release": (lambda x=lbl: self._set_criterio(x)),
            } for lbl in labels]
            kwargs = dict(caller=caller_btn, items=itens, width_mult=3, max_height="240dp")
        else:
            itens = [{
                "text": lbl,
                "on_release": (lambda x=lbl: self._set_criterio(x)),
            } for lbl in labels]
            kwargs = dict(caller=caller_btn, items=itens, width_mult=3, max_height="240dp")

        if self.menu_criterio:
            self.menu_criterio.dismiss()
        self.menu_criterio = MDDropdownMenu(**kwargs)
        self.menu_criterio.open()

    def _set_criterio(self, criterio: str):
        self.criterio_atual = criterio
        try:
            self.ids["btn_criterio_text"].text = f"Critério: {criterio}"
        except Exception:
            pass
        if self.menu_criterio:
            self.menu_criterio.dismiss()
        self._status(f"Critério definido: {criterio}")

    # ---------------- Ordenação ----------------
    def toggle_sort(self):
        self.sort_asc = not self.sort_asc
        try:
            self.ids.sort_button.icon = "sort-ascending" if self.sort_asc else "sort-descending"
        except Exception:
            pass

    # ---------------- Geração do mapa (thread) ----------------
    def gerar_mapa_html(self):
        if not self.input_path:
            _toast_safe("Importe uma base primeiro.")
            return

        # Recarrega o config do disco sempre que gerar
        self.cfg = _load_cfg()

        self._enable_gerar(False)
        self._status("Gerando mapa…")
        _toast_safe("Gerando…")

        def _worker():
            try:
                import main_rotas1 as rc1

                cfg = self.cfg or DEFAULT_CFG
                key = _load_google_key_safe().strip()
                valid_key = _is_probably_google_key(key)
                have_key = bool(key) and valid_key
                use_google = have_key and bool(cfg.get("USE_GOOGLE", True))

                # Estado desejado:
                desenhar_google = bool(cfg.get("DESENHAR_ROTA_GOOGLE", True))
                start_unchecked = bool(cfg.get("START_LAYERS_UNCHECKED", False))

                if desenhar_google and not have_key:
                    _toast_safe("Chave Google ausente/inválida — usando modo local (sem consumo de API).")

                kwargs = dict(
                    criterio=self.criterio_atual,
                    use_google=use_google,
                    google_key=(key if have_key else ""),  # só passa se válida

                    # Geração da linha do Google (só gera se toggle estiver ON e chave for válida)
                    desenhar_rota_google=(desenhar_google and have_key),

                    # Outras camadas:
                    mostrar_linha=False,
                    mostrar_ida_volta=True,
                    mostrar_deposito=True,

                    base_name="Mapa base",
                    sort_asc=self.sort_asc,

                    start_layers_unchecked=False,

                    # Mostrar checkbox “Rota (Google)” só se for possível gerar
                    forcar_layer_google=(desenhar_google and have_key),

                    # Se gerar a rota Google, já abrir marcada
                    show_google_layer=(desenhar_google and use_google and not start_unchecked),
                )

                out_html = rc1.gerar_mapa_from_path(
                    self.input_path,
                    **kwargs,
                    incluir_deposito_no_km=bool(cfg.get("INCLUIR_DEPOSITO_NO_KM", False)),
                    calcular_km_otimizado=bool(cfg.get("CALCULAR_KM_OTIMIZADO", True)),
                )

                self._status(f"Mapa gerado: {os.path.basename(out_html)}")
                _toast_safe("Mapa gerado. Use os checkboxes para mostrar as rotas.")
                try:
                    os.startfile(out_html)
                except Exception:
                    pass

            except Exception as e:
                self._status(f"Erro: {e}")
                _toast_safe("Falha ao gerar mapa. Veja o status.")
            finally:
                self._enable_gerar(True)

        threading.Thread(target=_worker, daemon=True).start()

    # ---- demais geradores ----
    def gerar_pdf_rota(self):
        if not self.input_path:
            _toast_safe("Importe uma base primeiro.")
            return
        self._status("Gerando PDF da Rota...")

        def _w():
            try:
                import os, shutil
                import main_relatorios as rel

                out = rel.gerar_rota_pdf_from_path(self.input_path, pagina="A4")
                if not out or not os.path.exists(out):
                    self._status("Erro: arquivo PDF não encontrado")
                    _toast_safe("Falha ao gerar o PDF da Rota.")
                    return

                bases_pdf = os.path.join(os.path.dirname(__file__), "bases", "pdf")
                os.makedirs(bases_pdf, exist_ok=True)
                fname = os.path.basename(out)
                dst = os.path.join(bases_pdf, fname)

                norm = lambda p: os.path.normcase(os.path.normpath(p))
                if norm(out) != norm(dst):
                    try:
                        os.replace(out, dst)
                    except Exception:
                        shutil.copy2(out, dst)
                        try:
                            os.remove(out)
                        except Exception:
                            pass
                else:
                    dst = out

                self._status(f"Gerado: {os.path.basename(dst)}")
                _toast_safe("PDF da Rota gerado.")
            except Exception as e:
                self._status(f"Erro: {e}")
                _toast_safe("Falha ao gerar o PDF da Rota.")

        threading.Thread(target=_w, daemon=True).start()

    def gerar_pdf_qrcode(self, usar_link_maps=True):
        if not self.input_path:
            _toast_safe("Importe uma base primeiro.")
            return
        self._status("Gerando PDF com QR Codes...")

        def _w():
            try:
                import os, shutil
                import main_relatorios as rel

                out = rel.gerar_qrcode_pdf_from_path(
                    self.input_path,
                    usar_sequence=True,
                    usar_link_maps=usar_link_maps
                )
                if not out or not os.path.exists(out):
                    self._status("Erro: arquivo PDF não encontrado")
                    _toast_safe("Falha ao gerar o PDF com QR.")
                    return

                bases_pdf = os.path.join(os.path.dirname(__file__), "bases", "pdf")
                os.makedirs(bases_pdf, exist_ok=True)
                fname = os.path.basename(out)
                dst = os.path.join(bases_pdf, fname)

                norm = lambda p: os.path.normcase(os.path.normpath(p))
                if norm(out) != norm(dst):
                    try:
                        os.replace(out, dst)
                    except Exception:
                        shutil.copy2(out, dst)
                        try:
                            os.remove(out)
                        except Exception:
                            pass
                else:
                    dst = out

                self._status(f"Gerado: {os.path.basename(dst)}")
                _toast_safe("PDF com QR Codes gerado.")
            except Exception as e:
                self._status(f"Erro: {e}")
                _toast_safe("Falha ao gerar o PDF com QR.")

        threading.Thread(target=_w, daemon=True).start()

    def gerar_pdf_qrcode_static(self):
        if not self.input_path:
            _toast_safe("Importe uma base primeiro.")
            return

        self._status("Gerando PDF (QR + miniaturas do mapa)...")

        def _w():
            try:
                import os, shutil
                import main_relatorios as rel

                key = _load_google_key_safe().strip()
                valid_key = _is_probably_google_key(key)

                out = rel.gerar_qrcode_pdf_from_path(
                    self.input_path,
                    usar_sequence=True,
                    usar_link_maps=True,
                    usar_miniatura=bool(valid_key)  # << se não tiver chave válida, NÃO usa Maps Static
                )
                if not out or not os.path.exists(out):
                    self._status("Erro: arquivo PDF não encontrado")
                    _toast_safe("Falha ao gerar o PDF com miniaturas.")
                    return

                bases_pdf = os.path.join(os.path.dirname(__file__), "bases", "pdf")
                os.makedirs(bases_pdf, exist_ok=True)
                fname = os.path.basename(out)
                dst = os.path.join(bases_pdf, fname)

                norm = lambda p: os.path.normcase(os.path.normpath(p))
                if norm(out) != norm(dst):
                    try:
                        os.replace(out, dst)
                    except Exception:
                        shutil.copy2(out, dst)
                        try:
                            os.remove(out)
                        except Exception:
                            pass
                else:
                    dst = out

                label = " (miniatura ligada)" if valid_key else " (miniatura desativada – sem chave)"
                self._status(f"Gerado: {os.path.basename(dst)}{label}")
                _toast_safe("PDF gerado.")
            except Exception as e:
                if "403" in str(e) or "Forbidden" in str(e):
                    self._status("Maps Static bloqueado (403). Habilite a API e refaça.")
                else:
                    self._status(f"Erro: {e}")
                _toast_safe("Falha ao gerar o PDF com miniaturas.")

        threading.Thread(target=_w, daemon=True).start()

    def gerar_planilha_whatsapp(self):
        if not self.input_path:
            _toast_safe("Importe uma base primeiro.")
            return
        self._status("Gerando planilha de WhatsApp...")

        def _w():
            try:
                import os, shutil
                import main_relatorios as rel

                out = rel.gerar_envio_whatsapp_from_path(self.input_path)
                if not out or not os.path.exists(out):
                    self._status("Erro: planilha não encontrada")
                    _toast_safe("Falha ao gerar a planilha.")
                    return

                bases_envio = os.path.join(os.path.dirname(__file__), "bases", "xlsx", "envio")
                os.makedirs(bases_envio, exist_ok=True)
                fname = os.path.basename(out) if os.path.basename(out) else "EnvioWhatsapp.xlsx"
                dst = os.path.join(bases_envio, fname)

                norm = lambda p: os.path.normcase(os.path.normpath(p))
                if norm(out) != norm(dst):
                    try:
                        os.replace(out, dst)
                    except Exception:
                        shutil.copy2(out, dst)
                        try:
                            os.remove(out)
                        except Exception:
                            pass
                else:
                    dst = out

                self._status(f"Gerado: {os.path.basename(dst)}")
                _toast_safe("Planilha WhatsApp gerada.")
            except Exception as e:
                self._status(f"Erro: {e}")
                _toast_safe("Falha ao gerar a planilha.")

        threading.Thread(target=_w, daemon=True).start()

    def gerar(self):
        self.gerar_mapa_html()


if __name__ == "__main__":
    RotasApp().run()

