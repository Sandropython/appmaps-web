# -*- coding: utf-8 -*-
"""
WhatsApp - Envio em Lote (Kivy/KivyMD + Selenium)
Fluxo: importar base (xlsx) -> escolher aba -> pré-visualizar -> enviar (Selenium)
- Uma única sessão do WhatsApp Web (mais rápido/estável)
- Verifica número inválido
- Removidos: inserir variáveis na UI e anexos, modo simples
"""

import os
import re
import time
import threading
from pathlib import Path

import pandas as pd

from kivy.lang import Builder
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.properties import StringProperty, ListProperty, BooleanProperty

from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.menu import MDDropdownMenu

APP_DIR = os.path.dirname(__file__)
UI_DIR = os.path.join(APP_DIR, "ui")
KV_FILE = os.path.join(UI_DIR, "main_whatsapp.kv")

Window.size = (360, 640)

for _icon in ("app_icon.ico", "app_icon.png", "icon.ico", "icon.png", "LogoFinal.png"):
    _p = os.path.join(UI_DIR, _icon)
    if os.path.exists(_p):
        try:
            Window.set_icon(_p)
            break
        except Exception:
            pass

# lembrar diretório do último uso
LAST_DIR_FILE = os.path.join(str(Path.home()), ".whats_last_dir.txt")

# diretório padrão
DEFAULT_ENVIO_DIR = os.path.join(APP_DIR, "bases", "xlsx", "envio")


class WhatsAppApp(MDApp):
    input_whats = StringProperty("")
    whats_preview_rows = ListProperty([])
    open_in_web = BooleanProperty(True)
    selected_sheet = StringProperty("envio")
    _sheet_names: list[str] = []

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.file_manager: MDFileManager | None = None
        self._menu_sheets: MDDropdownMenu | None = None
        self.df_envio: pd.DataFrame | None = None
        self._stop_flag = False

    def build(self):
        if not os.path.exists(KV_FILE):
            raise FileNotFoundError(f"KV não encontrado: {KV_FILE}")
        root = Builder.load_file(KV_FILE)
        self.theme_cls.theme_style = "Light"
        self.theme_cls.primary_palette = "Blue"
        return root

    @property
    def ids(self):
        return getattr(self.root, "ids", {})

    def _set_status(self, msg: str):
        Clock.schedule_once(lambda *_: self._set_lbl("lbl_status_whats", f"Status: {msg}"))

    def _set_lbl(self, lbl_id: str, text: str):
        try:
            self.ids[lbl_id].text = text
        except Exception:
            pass

    # AppBar
    def voltar(self, *args):
        try:
            self.stop()  # fecha esta janela/app de envio WhatsApp
        except Exception:
            pass
    try:
        from kivy.app import App
        App.get_running_app().stop()
    except Exception:
        pass

    def whats_info(self, *args):
        self._set_status("Fluxo: importar → escolher aba → pré-visualizar → enviar (Selenium).")

    def whats_menu(self, *args):
        self._set_status("Menu (placeholder).")

    # contador
    def _refresh_count(self):
        try:
            if not self.input_whats or not self.selected_sheet:
                self.ids["lbl_contagem"].text = "Contatos: 0"
                return
            df = pd.read_excel(self.input_whats, sheet_name=self.selected_sheet)
            self.ids["lbl_contagem"].text = f"Contatos: {len(df)}"
        except Exception as e:
            self._set_status(f"Erro ao contar linhas: {e}")

    # file manager
    def _get_last_dir(self) -> str:
        if os.path.isdir(DEFAULT_ENVIO_DIR):
            return DEFAULT_ENVIO_DIR
        try:
            if os.path.exists(LAST_DIR_FILE):
                with open(LAST_DIR_FILE, "r", encoding="utf-8") as f:
                    d = f.read().strip()
                    if d and os.path.isdir(d):
                        return d
        except Exception:
            pass
        return str(Path.home())

    def _set_last_dir(self, d: str):
        try:
            with open(LAST_DIR_FILE, "w", encoding="utf-8") as f:
                f.write(d)
        except Exception:
            pass

    def whats_importar_base(self):
        initial_path = self._get_last_dir()
        self._set_status(f"Abrindo base em: {initial_path}")
        if self.file_manager is None:
            self.file_manager = MDFileManager(
                exit_manager=self._close_file_manager,
                select_path=self._select_file,
                preview=False,
                search="all",
                use_access=True,
            )
        self.file_manager.show(initial_path)

    def _close_file_manager(self, *args):
        if self.file_manager:
            self.file_manager.close()

    def _select_file(self, path: str):
        if os.path.isdir(path):
            self._set_last_dir(path)
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".xls"):
            self._set_status("Escolha um arquivo .xlsx ou .xls")
            return
        self.input_whats = path
        self._set_last_dir(os.path.dirname(path))
        self._set_status("Base de contatos selecionada.")
        self.whats_preview_rows = []
        self.df_envio = None
        self._load_sheet_names()
        self._refresh_count()
        self._close_file_manager()

    # planilha
    def _load_sheet_names(self):
        self._sheet_names = []
        try:
            with pd.ExcelFile(self.input_whats) as xls:
                self._sheet_names = xls.sheet_names
        except Exception as e:
            self._set_status(f"Erro lendo abas: {e}")
            return

        if "envio" in [s.lower() for s in self._sheet_names]:
            for nm in self._sheet_names:
                if nm.lower() == "envio":
                    self.selected_sheet = nm
                    break
        elif self._sheet_names:
            self.selected_sheet = self._sheet_names[0]
        else:
            self.selected_sheet = ""

        try:
            self.ids["btn_sheet_text"].text = self.selected_sheet or "Selecionar"
        except Exception:
            pass

    def open_sheet_menu(self, caller_widget):
        if not self._sheet_names:
            self._load_sheet_names()
        items = [{"text": nm, "on_release": lambda x=nm: self._choose_sheet(x)} for nm in self._sheet_names]
        if self._menu_sheets:
            self._menu_sheets.dismiss()
        self._menu_sheets = MDDropdownMenu(caller=caller_widget, items=items, width_mult=3)
        self._menu_sheets.open()

    def _choose_sheet(self, name: str):
        self.selected_sheet = name
        try:
            self.ids["btn_sheet_text"].text = self.selected_sheet or "Selecionar"
        except Exception:
            pass
        if self._menu_sheets:
            self._menu_sheets.dismiss()
        self._set_status(f"Aba selecionada: {name}")
        self._refresh_count()

    # utils
    def _find_col(self, df: pd.DataFrame, keys: list[str]) -> str | None:
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

    def _normalize_phone_br(self, raw: str) -> str:
        digits = re.sub(r"\D", "", str(raw or ""))
        if not digits:
            return ""
        if digits.startswith("55"):
            digits = digits[2:]
        digits = digits.lstrip("0")
        if len(digits) not in (10, 11):
            return ""
        return "55" + digits

    def _compose_message(self, template_ui: str, row: dict) -> str:
        """
        Regra:
          - Se houver texto na coluna 'Mensagem' da planilha (saudação), ele vem PRIMEIRO.
          - Em seguida vem o texto do campo da tela (template_ui), em nova linha.
          - Substitui 'cliente' -> Nome e {Nome} -> Nome em AMBAS as partes.
          - Se só existir um dos dois, envia só aquele.
        """
        nome = str(row.get("Nome", "")).strip()

        def sub_cli(txt: str) -> str:
            if not txt:
                return ""
            # 'cliente' (case-insensitive, palavra inteira) -> Nome
            txt = re.sub(r"\bcliente\b", lambda m: nome if nome else m.group(0), txt, flags=re.IGNORECASE)
            txt = txt.replace("{Nome}", nome)
            return txt.replace("\r\n", "\n").replace("\r", "\n")

        msg_planilha = sub_cli(str(row.get("Mensagem", "")).strip())
        msg_tela = sub_cli((template_ui or "").strip())

        if msg_planilha and msg_tela:
            return f"{msg_planilha}\n{msg_tela}"
        return msg_planilha or msg_tela

    # preview
    def whats_preview(self):
        if not self.input_whats:
            self._set_status("Importe a base primeiro.")
            return
        if not self.selected_sheet:
            self._set_status("Selecione a aba.")
            return

        try:
            df = pd.read_excel(self.input_whats, sheet_name=self.selected_sheet)
        except Exception as e:
            self._set_status(f"Erro lendo a aba: {e}")
            return

        c_nome = self._find_col(df, ["nome"])
        c_tel = self._find_col(df, ["telefone", "fone", "whatsapp", "mobile", "celular"])
        c_grp = self._find_col(df, ["grupo", "group"])
        c_msg = self._find_col(df, ["mensagem", "message"])

        if not all([c_nome, c_tel]):
            self._set_status("Colunas mínimas não encontradas (Nome/Telefone).")
            return

        df_out = pd.DataFrame({
            "Nome": df[c_nome].astype(str).fillna("").str.strip(),
            "Telefone": df[c_tel].astype(str),
            "Grupo": df[c_grp].astype(str) if c_grp else "",
            "Mensagem": df[c_msg].astype(str) if c_msg else "",
        })
        df_out["Telefone"] = df_out["Telefone"].map(self._normalize_phone_br)

        self.df_envio = df_out
        # preview = [f'{r["Nome"]} | {r["Telefone"] or "inválido"} | {r["Grupo"]}'
        preview = [f'{r["Nome"]}'
                   for _, r in df_out.iterrows()]
        self.whats_preview_rows = preview

        try:
            self.ids["lbl_contagem"].text = f"Contatos: {len(df_out)}"
        except Exception:
            pass
        try:
            rv = self.ids.get("rv_preview")
            if rv:
                rv.data = [{"text": row} for row in preview]
        except Exception:
            pass

        self._set_status("Pré-visualização atualizada.")

    # envio com Selenium (uma sessão)
    def whats_enviar_selenium(self):
        if self.df_envio is None or self.df_envio.empty:
            self._set_status("Faça a pré-visualização primeiro.")
            return

        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.common.keys import Keys
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from webdriver_manager.chrome import ChromeDriverManager
            from selenium.webdriver.chrome.service import Service as ChromeService
            from selenium.webdriver.chrome.options import Options as ChromeOptions
        except Exception as e:
            self._set_status(f"Selenium não instalado: {e}")
            return

        try:
            template = self.ids["tf_msg"].text or ""
        except Exception:
            template = ""

        try:
            delay_ms = int(self.ids["tf_delay_ms"].text.strip())
        except Exception:
            delay_ms = 800
        delay = max(200, delay_ms) / 1000.0

        chrome_options = ChromeOptions()
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--no-default-browser-check")
        # perfil local ao projeto (evita conflito com perfis do Chrome em uso)
        user_data_dir = os.path.join(APP_DIR, ".wa_profile")
        os.makedirs(user_data_dir, exist_ok=True)
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")

        try:
            driver = webdriver.Chrome(
                service=ChromeService(ChromeDriverManager().install()),
                options=chrome_options
            )
        except Exception as e:
            self._set_status(f"Falha ao iniciar ChromeDriver: {e}")
            return

        self._set_status("Abrindo WhatsApp Web...")
        driver.get("https://web.whatsapp.com/")

        try:
            WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#app')))
            WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[id="side"]')))
        except Exception:
            self._set_status("Não consegui abrir/logar no WhatsApp Web.")
            try:
                driver.quit()
            except Exception:
                pass
            return

        self._stop_flag = False
        self._set_status("Enviando... (Selenium)")

        def _find_input_box(drv):
            """
            Tenta localizar a caixa de mensagem com seletores alternativos.
            """
            selectors = [
                'footer [contenteditable="true"]',
                'div[contenteditable="true"][data-tab="10"]',
                'div[contenteditable="true"][data-tab="6"]',
                'footer div[contenteditable="true"]',
            ]
            last_exc = None
            for sel in selectors:
                try:
                    return WebDriverWait(drv, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                    )
                except Exception as e:
                    last_exc = e
            if last_exc:
                raise last_exc

        def _worker():
            ok = 0
            fail = 0
            total = len(self.df_envio)

            for idx, row in self.df_envio.iterrows():
                if self._stop_flag:
                    break

                phone = row.get("Telefone", "")
                if not phone:
                    fail += 1
                    Clock.schedule_once(lambda *_: self._set_status(f"Telefone inválido (linha {idx+1})."))
                    continue

                # COMPOSIÇÃO: Mensagem (planilha) primeiro + template da tela abaixo
                msg = self._compose_message(template, row.to_dict()).strip()
                if not msg:
                    fail += 1
                    Clock.schedule_once(lambda *_: self._set_status(f"Nenhuma mensagem definida (linha {idx+1})."))
                    continue

                # 1) Abre conversa sem texto na URL (mais estável)
                chat_url = f"https://web.whatsapp.com/send?phone={phone}"
                try:
                    driver.get(chat_url)
                except Exception:
                    fail += 1
                    Clock.schedule_once(lambda *_: self._set_status(f"Falha ao abrir conversa (linha {idx+1})."))
                    continue

                # 2) Aguarda conversa carregar
                try:
                    WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[id="side"]')))
                    input_box = _find_input_box(driver)
                    time.sleep(0.8)
                except Exception:
                    # número inválido? tenta detectar popup
                    try:
                        invalid_popup = driver.find_elements(
                            By.XPATH,
                            '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]'
                        )
                        if invalid_popup:
                            fail += 1
                            Clock.schedule_once(lambda *_: self._set_status(f"Número inválido (linha {idx+1})."))
                            time.sleep(1.0)
                            continue
                    except Exception:
                        pass

                    fail += 1
                    Clock.schedule_once(lambda *_: self._set_status(f"Timeout ao carregar conversa (linha {idx+1})."))
                    continue

                # 3) Digita a mensagem (com quebras de linha)
                try:
                    input_box.click()
                    try:
                        input_box.clear()  # pode não funcionar em contenteditable, não tem problema
                    except Exception:
                        pass

                    for i, line in enumerate(msg.split("\n")):
                        if i > 0:
                            input_box.send_keys(Keys.SHIFT, Keys.ENTER)
                            time.sleep(0.03)
                        if line:
                            input_box.send_keys(line)
                        time.sleep(0.02)
                except Exception:
                    fail += 1
                    Clock.schedule_once(lambda *_: self._set_status(f"Falha ao digitar (linha {idx+1})."))
                    continue

                # 4) Envia (Enter) com fallback de botão
                try:
                    input_box.send_keys(Keys.ENTER)
                    ok += 1
                except Exception:
                    try:
                        send_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '//*[@id="main"]/footer//button[@aria-label][last()]')
                            )
                        )
                        send_btn.click()
                        ok += 1
                    except Exception:
                        fail += 1
                        Clock.schedule_once(lambda *_: self._set_status(f"Falha ao enviar (linha {idx+1})."))
                        continue

                Clock.schedule_once(lambda *_: self._set_status(f"Enviado {ok}/{total} (falhas {fail})"))
                time.sleep(delay)

            # NÃO fecha o WhatsApp ao terminar (mantém sessão aberta)
            Clock.schedule_once(lambda *_: self._set_status(f"Concluído: {ok}/{total} enviados, {fail} falhas."))

        threading.Thread(target=_worker, daemon=True).start()

    def whats_limpar(self):
        self.input_whats = ""
        self.whats_preview_rows = []
        self.df_envio = None
        try:
            self.ids["lbl_contagem"].text = "Contatos: 0"
        except Exception:
            pass
        try:
            rv = self.ids.get("rv_preview")
            if rv:
                rv.data = []
        except Exception:
            pass
        self._set_status("Campos limpos.")


if __name__ == "__main__":
    WhatsAppApp().run()






