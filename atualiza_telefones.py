# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
from pathlib import Path

from kivy.core.window import Window
from kivy.lang import Builder
from kivy.metrics import dp
from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.screen import MDScreen
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivy.properties import StringProperty

# services reais do seu projeto
from telefones_service import executar_processo, carregar_base_pkl

Window.size = (360, 640)

KV_PATH = "ui/atualiza_telefones.kv"

# Caminho padrão da base
PADRAO_BASE_PKL = r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\pkl\base_cel.pkl"

# Saídas
ENVIO_DIR = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\xlsx\envio")

# Preferências simples
PREFS_DIR = Path.home() / "AppMaps"
PREFS_DIR.mkdir(parents=True, exist_ok=True)
PREFS_PATH = PREFS_DIR / "prefs_atualiza_telefones.json"


class AtualizaTelefonesScreen(MDScreen):
    pkl_path = StringProperty(PADRAO_BASE_PKL)
    excel_in = StringProperty("")


def _try_executar_processo(**kwargs):
    """Tolerante a assinaturas diferentes do service."""
    variants = [
        kwargs,
        {
            **{k: v for k, v in kwargs.items() if k != "saida_excel_atualizada"},
            "saida_cel_atualizada": kwargs.get("saida_excel_atualizada"),
        },
        {
            "caminho_excel_entrada": kwargs.get("caminho_excel_entrada"),
            "caminho_pkl_base": kwargs.get("caminho_pkl_base"),
        },
    ]
    last_exc = None
    for payload in variants:
        try:
            return executar_processo(**payload)
        except TypeError as e:
            last_exc = e
            continue
    if last_exc:
        raise last_exc
    raise RuntimeError("Falha ao chamar executar_processo().")


class _Controller:
    def __init__(self, app: MDApp):
        self.app = app
        self.file_manager: MDFileManager | None = None
        self._which: str | None = None  # "pkl" | "xlsx"
        self._target_field = None
        self._last_dir: Path | None = None

        # prefs
        prefs = self._read_prefs()
        if prefs.get("last_dir"):
            try:
                self._last_dir = Path(prefs["last_dir"])
            except Exception:
                self._last_dir = None

    # ---------- util ----------
    def _snack(self, text: str):
        try:
            MDSnackbar(MDSnackbarText(text=text), y=dp(24), pos_hint={"center_x": 0.5}).open()
        except Exception:
            print(text)

    @staticmethod
    def _read_prefs() -> dict:
        try:
            if PREFS_PATH.exists():
                return json.loads(PREFS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    @staticmethod
    def _write_prefs(prefs: dict):
        try:
            PREFS_PATH.write_text(json.dumps(prefs, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_prefs_into_screen(self):
        prefs = self._read_prefs()
        root = self.app.root.get_screen("AtualizaTelefonesScreen")
        root.ids.pkl_path.text = prefs.get("pkl_path") or PADRAO_BASE_PKL
        root.ids.excel_in.text = prefs.get("excel_in", "")

    # ---------- FILE MANAGER ----------
    def _abrir_filemanager(self, which: str, target_widget):
        """Abre o seletor compatível com KivyMD 2.x (sem 'ext')."""
        self._which = which
        self._target_field = target_widget

        # ponto de partida
        start = Path.home()
        try:
            texto = (target_widget.text or "").strip()
            if texto:
                p = Path(texto)
                start = p if p.is_dir() else p.parent
        except Exception:
            pass
        if not start.exists():
            start = Path.home()
        self._last_dir = start

        # GUARDA a instância no objeto (evita GC)
        self.file_manager = MDFileManager(
            select_path=self._select_path,
            exit_manager=self._close_manager,
            preview=False,
            search="all",
            use_access=True,
        )
        self.file_manager.show(str(start))

    def _close_manager(self, *args):
        if self.file_manager:
            try:
                self.file_manager.close()
            except Exception:
                pass
        self.file_manager = None
        self._which = None
        self._target_field = None

    def _select_path(self, path: str):
        """Valida extensão aqui (filtragem estável no 2.x) e preenche o TextField."""
        try:
            root = self.app.root.get_screen("AtualizaTelefonesScreen")
            p = Path(path)
            if self._which == "pkl":
                if p.suffix.lower() != ".pkl":
                    self._snack("Selecione um arquivo .pkl")
                    self._close_manager()
                    return
                root.ids.pkl_path.text = path
            elif self._which == "xlsx":
                if p.suffix.lower() != ".xlsx":
                    self._snack("Selecione um arquivo .xlsx")
                    self._close_manager()
                    return
                root.ids.excel_in.text = path

            # prefs e last_dir
            prefs = self._read_prefs()
            prefs["last_dir"] = str(p.parent)
            prefs["pkl_path"] = root.ids.pkl_path.text
            prefs["excel_in"] = root.ids.excel_in.text
            self._write_prefs(prefs)
            self._snack("Arquivo selecionado.")
        finally:
            self._close_manager()

    # ---------- helpers preview ----------
    def _conta_registros_base(self, base) -> int:
        try:
            if isinstance(base, dict):
                return len(base)
            if getattr(base, "__class__", None) and base.__class__.__name__ == "DataFrame":
                return int(len(base))
            if hasattr(base, "__len__"):
                return int(len(base))
        except Exception:
            pass
        return 0

    def _find_local_col(self, df) -> str | None:
        try:
            for c in df.columns:
                if str(c).strip().lower() == "local":
                    return c
        except Exception:
            pass
        return None

    def _get_from_resumo(self, resumo, keys, default=0):
        if isinstance(keys, (str, bytes)):
            keys = [keys]
        for k in keys:
            try:
                return getattr(resumo, k)
            except Exception:
                try:
                    return resumo.get(k, default)
                except Exception:
                    pass
        return default

    # ---------- ações ----------
    def alerta_info_atualiza_telefones(self):
        self._snack(
            "Preenche vazios; cria variação __N para pares; insere novos; "
            "gera planilha de BUSCA quando há múltiplos."
        )

    def preview_atualiza_telefones(self):
        import pandas as pd
        from telefones_service import atualizar_telefones

        try:
            root = self.app.root.get_screen("AtualizaTelefonesScreen")
            pkl_path = Path(root.ids.pkl_path.text.strip())
            excel_in = Path(root.ids.excel_in.text.strip())
            if not pkl_path.exists() or not excel_in.exists():
                raise FileNotFoundError("Selecione Base PKL e Planilha XLSX válidas.")

            base = carregar_base_pkl(pkl_path)
            df_in = pd.read_excel(excel_in, dtype={"Telefone": str})
            df_out, base_out, resumo, df_busca = atualizar_telefones(df_in, base, usar_complemento=False)

            local_col = self._find_local_col(df_in)
            total_locais = int(df_in[local_col].astype(str).nunique()) if local_col else 0
            qtd_base = self._conta_registros_base(base)

            linhas = self._get_from_resumo(resumo, ["linhas", "total_linhas", "qtd_linhas"], 0)
            nomes_p = self._get_from_resumo(resumo, ["nomes_preenchidos", "nomes_preench", "nomes_preench_qtd"], 0)
            tels_p = self._get_from_resumo(resumo, ["telefones_preenchidos", "tels_preenchidos", "telefones_preench"], 0)
            nomes_add = self._get_from_resumo(resumo, ["nomes_adicionais", "nomes_adicionados", "nomes_extra"], 0)
            tels_add = self._get_from_resumo(resumo, ["telefones_adicionais", "telefones_adicionados", "tels_extra"], 0)
            nao_base = self._get_from_resumo(resumo, ["nao_encontrados", "sem_base", "nao_encontrado"], 0)

            texto = (
                f"Linhas: {linhas}\n"
                f"Locais únicos: {total_locais}\n"
                f"Registros na base (PKL): {qtd_base}\n"
                f"Nomes preenchidos: {nomes_p}\n"
                f"Telefones preenchidos: {tels_p}\n"
                f"Nomes adicionais: {nomes_add}\n"
                f"Telefones adicionais: {tels_add}\n"
                f"Não encontrados (sem base): {nao_base}\n"
                f"Linhas com adicionais (BUSCA): {len(df_busca)}"
            )

            root.ids.resumo_lbl.text = texto
            self._snack("Prévia gerada.")
        except Exception as e:
            self._snack(f"Erro no Preview: {e}")

    def executar_atualiza_telefones(self):
        try:
            root = self.app.root.get_screen("AtualizaTelefonesScreen")
            pkl_path = Path(root.ids.pkl_path.text.strip())
            excel_in = Path(root.ids.excel_in.text.strip())
            if not pkl_path.exists() or not excel_in.exists():
                raise FileNotFoundError("Selecione Base PKL e Planilha XLSX válidas.")

            ENVIO_DIR.mkdir(parents=True, exist_ok=True)
            excel_out = ENVIO_DIR / f"{excel_in.stem}_Cel.xlsx"
            excel_busca = ENVIO_DIR / f"{excel_in.stem}_Busca.xlsx"

            result = _try_executar_processo(
                caminho_excel_entrada=excel_in,
                caminho_pkl_base=pkl_path,
                saida_excel_atualizada=excel_out,
                saida_excel_busca=excel_busca,
                usar_complemento=False,
            )

            # Contagem atual da base
            qtd_base = 0
            try:
                base_count_src = result.get("pkl_atualizado") or str(pkl_path)
                base_now = carregar_base_pkl(Path(base_count_src))
                qtd_base = self._conta_registros_base(base_now)
            except Exception:
                pass

            res = result.get("resumo", {}) or {}
            linhas = res.get("linhas") or res.get("total_linhas") or res.get("qtd_linhas") or 0
            nomes_p = res.get("nomes_preenchidos") or res.get("nomes_preench") or res.get("nomes_preench_qtd") or 0
            tels_p = res.get("telefones_preenchidos") or res.get("tels_preenchidos") or res.get("telefones_preench") or 0
            nomes_add = res.get("nomes_adicionais") or res.get("nomes_adicionados") or res.get("nomes_extra") or 0
            tels_add = res.get("telefones_adicionais") or res.get("telefones_adicionados") or res.get("tels_extra") or 0
            nao_base = res.get("nao_encontrados") or res.get("sem_base") or res.get("nao_encontrado") or 0

            texto = (
                f"Linhas: {linhas}\n"
                f"Registros na base (PKL): {qtd_base}\n"
                f"Nomes preenchidos: {nomes_p}\n"
                f"Telefones preenchidos: {tels_p}\n"
                f"Nomes adicionais: {nomes_add}\n"
                f"Telefones adicionais: {tels_add}\n"
                f"Não encontrados (sem base): {nao_base}"
            )
            root.ids.resumo_lbl.text = texto

            # salvar prefs
            prefs = self._read_prefs()
            prefs.update(
                {"pkl_path": root.ids.pkl_path.text.strip(), "excel_in": root.ids.excel_in.text.strip(),
                 "last_dir": prefs.get("last_dir")}
            )
            self._write_prefs(prefs)

            print(
                "\n==== Resultado Atualiza Telefones ====\n"
                f"Planilha atualizada: {result.get('excel_atualizada') or excel_out}\n"
                f"Base PKL atualizada: {result.get('pkl_atualizado')}\n"
                f"Planilha de busca: {result.get('excel_busca') or excel_busca}\n"
                f"Resumo: {result.get('resumo')}"
            )
            self._snack("Concluído.")
        except Exception as e:
            self._snack(f"Erro no Processo: {e}")


# ---------- bootstrap ----------
def load_screen(app: MDApp):
    if not app.root.has_screen("AtualizaTelefonesScreen"):
        Builder.load_file(KV_PATH)
        app.root.add_widget(AtualizaTelefonesScreen(name="AtualizaTelefonesScreen"))

    if not hasattr(app, "_atl_controller"):
        app._atl_controller = _Controller(app)
        app._abrir_filemanager = app._atl_controller._abrir_filemanager
        app.preview_atualiza_telefones = app._atl_controller.preview_atualiza_telefones
        app.executar_atualiza_telefones = app._atl_controller.executar_atualiza_telefones
        app.alerta_info_atualiza_telefones = app._atl_controller.alerta_info_atualiza_telefones
        try:
            app._atl_controller._load_prefs_into_screen()
        except Exception:
            pass

    return app.root.get_screen("AtualizaTelefonesScreen")


# ---------- execução standalone para teste ----------
if __name__ == "__main__":
    from kivymd.uix.screenmanager import MDScreenManager

    class _MiniApp(MDApp):
        def build(self):
            sm = MDScreenManager()
            self.root = sm
            load_screen(self)
            sm.current = "AtualizaTelefonesScreen"
            self.title = "Telefones"
            # ícone opcional
            self.logo_small = "ui/LogoFinal.png"
            self.placeholder = "ui/placeholder.png"
            if os.path.exists(self.logo_small):
                Window.set_icon(self.logo_small)

        def voltar_tela_anterior(self):
            self.root.current = "AtualizaTelefonesScreen"

    _MiniApp().run()


