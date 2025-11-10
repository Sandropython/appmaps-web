# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import json
from shutil import copy2
from datetime import datetime
from kivymd.app import MDApp
from kivy.properties import StringProperty
from kivymd.uix.screen import MDScreen
from kivymd.uix.filemanager import MDFileManager
from kivy.lang import Builder
from kivymd.uix.snackbar import MDSnackbar, MDSnackbarText
from kivy.metrics import dp
from kivy.core.window import Window
from telefones_service import executar_processo, carregar_base_pkl

Window.size = (360, 640)
KV_PATH = "ui/atualiza_telefones.kv"
PADRAO_BASE_PKL = r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\pkl\base_cel.pkl"
ENVIO_DIR = Path(r"C:\ProjetoApp\Maps\AppMaps\AppMaps\bases\xlsx\envio")
PREFS_DIR = Path.home() / "AppMaps"
PREFS_DIR.mkdir(parents=True, exist_ok=True)
PREFS_PATH = PREFS_DIR / "prefs_atualiza_telefones.json"

class AtualizaTelefonesScreen(MDScreen):
    pkl_path = StringProperty(PADRAO_BASE_PKL)
    excel_in = StringProperty("")

def _try_executar_processo(**kwargs):
    variants = [
        kwargs,
        {**{k: v for k, v in kwargs.items() if k != "saida_excel_atualizada"},
         "saida_cel_atualizada": kwargs.get("saida_excel_atualizada")},
        {"caminho_excel_entrada": kwargs.get("caminho_excel_entrada"),
         "caminho_pkl_base": kwargs.get("caminho_pkl_base")},
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

    def __init__(self, app: MDApp):
        self.app = app
        self.file_manager: MDFileManager | None = None
        self._target_field = None
        self._last_dir: Path | None = None
        try:
            prefs = self._read_prefs()
            if prefs.get("last_dir"):
                self._last_dir = Path(prefs["last_dir"])
        except Exception:
            pass

    def _snack(self, text: str) -> None:
        try:
            MDSnackbar(MDSnackbarText(text=text), y=dp(24), pos_hint={"center_x": 0.5}).open()
        except Exception as e:
            print("SNACKBAR:", text, "(erro:", e, ")")

    @staticmethod
    def _read_prefs() -> dict:
        try:
            if PREFS_PATH.exists():
                return json.loads(PREFS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    @staticmethod
    def _write_prefs(prefs: dict) -> None:
        try:
            PREFS_PATH.write_text(json.dumps(prefs, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_prefs_into_screen(self) -> None:
        prefs = self._read_prefs()
        root = self.app.root.get_screen("AtualizaTelefonesScreen")
        root.ids.pkl_path.text = prefs.get("pkl_path") or PADRAO_BASE_PKL
        if prefs.get("excel_in"):
            root.ids.excel_in.text = prefs["excel_in"]

    def _abrir_filemanager(self, tipo: str, target_widget) -> None:
        self._target_field = target_widget
        if self.file_manager is None:
            self.file_manager = MDFileManager(
                select_path=self._select_path,
                exit_manager=self._close_manager,
                ext=[f".{tipo}"] if tipo else None,
            )
        self.file_manager.ext = [f".{tipo}"] if tipo else None
        self.file_manager.show(str(self._last_dir or Path.home()))

    def _select_path(self, path: str) -> None:
        root = self.app.root.get_screen("AtualizaTelefonesScreen")
        if self._target_field is not None:
            self._target_field.text = path
        try:
            p = Path(path)
            self._last_dir = p if p.is_dir() else p.parent
        except Exception:
            pass
        prefs = self._read_prefs()
        if self._last_dir:
            prefs["last_dir"] = str(self._last_dir)
        try:
            if self._target_field is root.ids.excel_in:
                prefs["excel_in"] = path
            elif self._target_field is root.ids.pkl_path:
                prefs["pkl_path"] = path
        except Exception:
            pass
        self._write_prefs(prefs)
        self._close_manager()

    def _close_manager(self, *args) -> None:
        if self.file_manager:
            self.file_manager.close()

    def alerta_info_atualiza_telefones(self) -> None:
        self._snack("Preenche vazios; cria variação __N quando par diferente; insere quando novo; BUSCA só com adicionais.")

    def preview_atualiza_telefones(self) -> None:
        from telefones_service import atualizar_telefones
        import pandas as pd
        try:
            root = self.app.root.get_screen("AtualizaTelefonesScreen")
            pkl_path = Path(root.ids.pkl_path.text.strip())
            excel_in = Path(root.ids.excel_in.text.strip())
            if not pkl_path.exists() or not excel_in.exists():
                raise FileNotFoundError("Selecione Base PKL e Planilha de entrada válidas.")
            base = carregar_base_pkl(pkl_path)
            df_in = pd.read_excel(excel_in, dtype={"Telefone": str})
            df_out, base_out, resumo, df_busca = atualizar_telefones(df_in, base, usar_complemento=False)
            local_col = self._find_local_col(df_in)
            total_locais = int(df_in[local_col].astype(str).nunique()) if local_col else 0
            qtd_base = self._conta_registros_base(base)
            linhas    = self._get_from_resumo(resumo, ["linhas","total_linhas","qtd_linhas"], 0)
            nomes_p   = self._get_from_resumo(resumo, ["nomes_preenchidos","nomes_preench","nomes_preench_qtd"], 0)
            tels_p    = self._get_from_resumo(resumo, ["telefones_preenchidos","tels_preenchidos","telefones_preench"], 0)
            nomes_add = self._get_from_resumo(resumo, ["nomes_adicionais","nomes_adicionados","nomes_extra"], 0)
            tels_add  = self._get_from_resumo(resumo, ["telefones_adicionais","telefones_adicionados","tels_extra"], 0)
            nao_base  = self._get_from_resumo(resumo, ["nao_encontrados","sem_base","nao_encontrado"], 0)
            texto = (f"Linhas: {linhas}\n"
                     f"Locais únicos: {total_locais}\n"
                     f"Registros na base (PKL): {qtd_base}\n"
                     f"Nomes preenchidos: {nomes_p}\n"
                     f"Telefones preenchidos: {tels_p}\n"
                     f"Nomes adicionais: {nomes_add}\n"
                     f"Telefones adicionais: {tels_add}\n"
                     f"Não encontrados (sem base): {nao_base}\n"
                     f"Linhas com adicionais (BUSCA): {len(df_busca)}")
            root.ids.resumo_lbl.text = texto
            self._snack("Prévia gerada.")
        except Exception as e:
            self._snack(f"Erro no Preview: {e}")

    def executar_atualiza_telefones(self) -> None:
        try:
            root = self.app.root.get_screen("AtualizaTelefonesScreen")
            pkl_path = Path(root.ids.pkl_path.text.strip())
            excel_in = Path(root.ids.excel_in.text.strip())
            if not pkl_path.exists() or not excel_in.exists():
                raise FileNotFoundError("Selecione Base PKL e Planilha de entrada válidas.")
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
            # Backup + atualização da base fixa
            src_updated_pkl = result.get("pkl_atualizado") or str(pkl_path)
            if pkl_path.exists():
                ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                bckp = pkl_path.with_name(f"{pkl_path.stem}_bckp_{ts}{pkl_path.suffix}")
                try:
                    copy2(str(pkl_path), str(bckp))
                except Exception as e:
                    print(f"[ATENÇÃO] Não foi possível criar backup: {e}")
            try:
                if Path(src_updated_pkl).resolve() != pkl_path.resolve():
                    copy2(str(src_updated_pkl), str(pkl_path))
                print(f"[OK] Base fixa atualizada em: {pkl_path}")
            except Exception as e:
                print(f"[ERRO] Falha ao atualizar base fixa: {e}")
            res = result.get("resumo", {})
            qtd_base = 0
            try:
                base_now = carregar_base_pkl(pkl_path)
                qtd_base = self._conta_registros_base(base_now)
            except Exception:
                pass
            linhas    = res.get("linhas") or res.get("total_linhas") or res.get("qtd_linhas") or 0
            nomes_p   = res.get("nomes_preenchidos") or res.get("nomes_preench") or res.get("nomes_preench_qtd") or 0
            tels_p    = res.get("telefones_preenchidos") or res.get("tels_preenchidos") or res.get("telefones_preench") or 0
            nomes_add = res.get("nomes_adicionais") or res.get("nomes_adicionados") or res.get("nomes_extra") or 0
            tels_add  = res.get("telefones_adicionais") or res.get("telefones_adicionados") or res.get("tels_extra") or 0
            nao_base  = res.get("nao_encontrados") or res.get("sem_base") or res.get("nao_encontrado") or 0
            texto = (f"Linhas: {linhas}\n"
                     f"Registros na base (PKL): {qtd_base}\n"
                     f"Nomes preenchidos: {nomes_p}\n"
                     f"Telefones preenchidos: {tels_p}\n"
                     f"Nomes adicionais: {nomes_add}\n"
                     f"Telefones adicionais: {tels_add}\n"
                     f"Não encontrados (sem base): {nao_base}")
            root.ids.resumo_lbl.text = texto
            prefs = self._read_prefs()
            prefs.update({"pkl_path": root.ids.pkl_path.text.strip(),
                          "excel_in": root.ids.excel_in.text.strip(),
                          "last_dir": prefs.get("last_dir")})
            self._write_prefs(prefs)
            print("\n==== Resultado Atualiza Telefones ====\n"
                  f"Planilha atualizada: {result.get('excel_atualizada') or excel_out}\n"
                  f"Base PKL atualizada: {str(pkl_path)}\n"
                  f"Planilha de busca: {result.get('excel_busca') or excel_busca}\n"
                  f"Resumo: {result.get('resumo')}")
            self._snack("Concluído.")
        except Exception as e:
            self._snack(f"Erro no Processo: {e}")

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
