# -*- coding: utf-8 -*-
"""
Arquivo: main.py (versão comentada)
Descrição: Aplicativo KivyMD para processamento de planilhas, busca de coordenadas
           via Google Maps e gerenciamento de uma base local.

Obs.: Todos os comentários foram escritos em português para facilitar a leitura.
      Mantive a estrutura original, apenas acrescentei comentários linha a linha
      (ou em blocos) sempre que possível sem poluir demais.
"""

# ------------------------------ Imports padrão Kivy/KivyMD ------------------------------
from kivy.core.window import Window  # controla propriedades da janela (tamanho, ícone)
from kivy.lang import Builder  # carrega arquivos .kv (templates das telas)
from kivymd.app import MDApp  # classe base para apps KivyMD
from kivymd.uix.screen import MDScreen  # telas individuais
from kivymd.uix.screenmanager import MDScreenManager  # gerenciador de telas
from kivymd.uix.filemanager import MDFileManager  # seletor de arquivos (explorador)

# Componentes de AppBar (topo) e botões de ação
from kivymd.uix.appbar import (
    MDTopAppBar, MDTopAppBarTitle,
    MDTopAppBarLeadingButtonContainer, MDTopAppBarTrailingButtonContainer,
    MDActionTopAppBarButton
)

# Componentes de lista (itens com título e texto de apoio)
from kivymd.uix.list import (
    MDList, MDListItem, MDListItemHeadlineText, MDListItemSupportingText
)

from kivy.core.clipboard import Clipboard  # acesso à área de transferência do SO

# >>> ADIÇÕES (para o tweak do MDFileManager)
from kivy.clock import Clock
from kivy.metrics import dp

# ------------------------------ Imports de sistema/bibliotecas ------------------------------
import os  # manipulação de caminhos/arquivos
import pandas as pd  # manipulação de planilhas/DataFrames
import googlemaps  # cliente oficial Google Maps API
import kivy  # núcleo do Kivy (para ajustar logs)
import subprocess  # abrir processo separado (p/ main_rotas.py)
import sys  # informações do interpretador Python atual
import re  # expressões regulares para normalização de logradouro

# Reduz verbosidade do logger do Kivy (opcional)
kivy.logger.Logger.setLevel('WARNING')

# Define tamanho padrão da janela (útil no desktop)
Window.size = (360, 640)

# ------------------------------ Constantes/Configuração ------------------------------
CONFIG_FILE = "last_dir.txt"  # armazena último diretório acessado no seletor de arquivos
GOOGLE_API_KEY = "AIzaSyDNwadAVMmLfK7Lt-kpPmf2VbwKsh7fQXQ"  # sua chave local (sugere-se usar variável de ambiente em produção)

# ------------------------------ Esquema padrão da Base ------------------------------
# Mantém consistência das colunas da base local (PKL/XLSX)
BASE_COLUMNS = [
    'ID', 'Local', 'Bairro', 'CEP', 'Nome', 'Telefone',
    'Latitude', 'Longitude',
    'Bairro_Maps', 'Cep_Maps', 'Cidade_Maps', 'Endereco_Formatado'
]


# ------------------------------ Telas Simples ------------------------------
class HomeScreen(MDScreen):
    """Tela inicial (apenas declarada; layout em arquivo .kv)."""
    pass


class ProcessarPlanilhaScreen(MDScreen):
    """Tela para processar planilhas (layout em .kv)."""
    pass


class BaseScreen(MDScreen):
    """Tela para visualizar/operar sobre a base local (layout em .kv)."""
    pass


# ------------------------------ Tela: Resultados da Busca ------------------------------
class ResultadosBuscaScreen(MDScreen):
    """
    Apresenta resultados de uma busca na base com:
      - Lista rolável
      - Coluna Id (respeitando coluna de origem; normaliza para 'Id')
      - Ações: exportar para Excel, copiar IDs, abrir correção de coordenadas
      - Botão de ordenação asc/desc por Id
    """

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        from kivymd.uix.boxlayout import MDBoxLayout  # layout vertical/horizontal
        from kivy.uix.scrollview import ScrollView  # contêiner com rolagem
        from kivymd.uix.button import MDButton, MDButtonText  # botões KivyMD 2.x

        self.df_resultados = None  # DataFrame atual exibido
        self.sort_asc = True  # estado de ordenação (True = ascendente)

        # Layout raiz da tela
        root = MDBoxLayout(
            orientation="vertical",
            spacing="8dp",
            padding=("8dp", "8dp", "8dp", "8dp"),
        )

        # ---------------- AppBar (topo) ----------------
        self.appbar = MDTopAppBar(type="small")  # barra superior compacta
        # Botão de voltar (leading)
        self.appbar.add_widget(MDTopAppBarLeadingButtonContainer(
            MDActionTopAppBarButton(icon="arrow-left",
                                    on_release=lambda *_: self._voltar_impl())
        ))
        # Título
        self.appbar.add_widget(MDTopAppBarTitle(text="Resultados da Busca"))
        # Botão de ordenação (trailing)
        self.sort_button = MDActionTopAppBarButton(icon="sort",
                                                   on_release=lambda *_: self.toggle_sort())
        self.appbar.add_widget(MDTopAppBarTrailingButtonContainer(
            self.sort_button
        ))
        root.add_widget(self.appbar)

        # ---------------- Lista com rolagem ----------------
        self.scroll = ScrollView()  # viewport rolável
        self.list_container = MDList()  # lista de itens (será preenchida dinamicamente)
        self.scroll.add_widget(self.list_container)
        root.add_widget(self.scroll)

        # ---------------- Barra de ações (rodapé) ----------------
        from kivymd.uix.boxlayout import MDBoxLayout as _MDBox
        actions = _MDBox(
            orientation="horizontal",
            size_hint_y=None,
            height="48dp",
            spacing="8dp",
        )

        # Botão: exportar para Excel (usa callback na App)
        self.btn_exportar = MDButton(
            MDButtonText(text="Excel"),
            style="filled",
            on_release=lambda *_: self._exportar_impl(),
        )
        # Botão: copiar IDs para área de transferência
        self.btn_copiar_ids = MDButton(
            MDButtonText(text="Cop IDs"),
            style="tonal",
            on_release=lambda *_: self._copiar_ids_impl(),
        )
        # Botão: abrir diálogo de correção de coordenadas
        self.btn_corrigir = MDButton(
            MDButtonText(text="><"),
            style="tonal",
            on_release=lambda *_: self._corrigir_impl(),
        )
        # Botão: voltar para a tela "base"
        self.btn_voltar = MDButton(
            MDButtonText(text="Voltar"),
            style="tonal",
            on_release=lambda *_: self._voltar_impl(),
        )

        # Adiciona botões na barra de ações
        actions.add_widget(self.btn_exportar)
        actions.add_widget(self.btn_copiar_ids)
        actions.add_widget(self.btn_corrigir)
        actions.add_widget(self.btn_voltar)
        root.add_widget(actions)

        # Finaliza montagem da tela
        self.add_widget(root)

    # ---------- API p/ App ----------
    def set_results(self, df: pd.DataFrame):
        """
        Recebe um DataFrame e prepara para exibição:
          1) Identifica coluna de ID existente e normaliza o nome para 'Id'.
          2) Garante tipo numérico em 'Id' para ordenação.
          3) Ordena 'Id' de forma ascendente por padrão.
        """
        if df is None or df.empty:
            # Caso vazio, prepara DataFrame com colunas padrão usadas na renderização
            self.df_resultados = pd.DataFrame(columns=["Id", "Local", "Complemento", "Latitude", "Longitude"])
            self.sort_asc = True
            self._render_list()
            return

        df = df.copy()  # evita alterar DataFrame externo

        # Detecta nome da coluna de ID (variações comuns)
        id_col = next((c for c in ["Id", "ID", "id", "iD"] if c in df.columns), None)

        if id_col is None:
            # Se não existir coluna de Id, cria um sequencial apenas para exibição
            df = df.reset_index(drop=True)
            df.insert(0, "Id", df.index + 1)
        else:
            # Renomeia a coluna detectada para 'Id'
            if id_col != "Id":
                df.rename(columns={id_col: "Id"}, inplace=True)

        # Garante tipo inteiro em 'Id' (coerção segura)
        df["Id"] = pd.to_numeric(df["Id"], errors="coerce").fillna(0).astype(int)

        # Ordenação inicial por Id ascendente
        self.sort_asc = True
        df = df.sort_values(by="Id", ascending=self.sort_asc).reset_index(drop=True)

        self.df_resultados = df
        self._render_list()  # desenha na tela

    def toggle_sort(self):
        """Alterna ordenação por 'Id' (asc/desc) e re-renderiza a lista."""
        if self.df_resultados is None or self.df_resultados.empty:
            return
        self.sort_asc = not self.sort_asc
        self.df_resultados = self.df_resultados.sort_values(by='Id', ascending=self.sort_asc).reset_index(drop=True)
        self._render_list()

    # ---------- Internos ----------
    def _render_list(self):
        """Popula a MDList com os itens (cabeçalho + texto de apoio)."""
        self.list_container.clear_widgets()  # limpa a lista atual

        if self.df_resultados is None or self.df_resultados.empty:
            # Mensagem padrão quando não há dados
            item = MDListItem()
            item.add_widget(MDListItemHeadlineText(text="Nenhum resultado"))
            item.add_widget(MDListItemSupportingText(text="Refaça a busca"))
            self.list_container.add_widget(item)
            return

        # Cria um item para cada linha do DataFrame
        for _, row in self.df_resultados.iterrows():
            _id = int(row.get('Id', 0))  # Id normalizado
            local = str(row.get('Local', ''))  # endereço/descrição
            comp = row.get('Complemento', '')
            comp = '' if (pd.isna(comp) or comp is None) else str(comp)
            prim = f"[{_id}] {local if not comp else f'{local} | {comp}'}"  # título
            lat = row.get('Latitude', '')
            lng = row.get('Longitude', '')
            sec = f"Lat: {lat}  Lon: {lng}" if (lat != '' or lng != '') else "Sem coordenadas"  # subtítulo

            item = MDListItem()
            item.add_widget(MDListItemHeadlineText(text=prim))
            item.add_widget(MDListItemSupportingText(text=sec))
            self.list_container.add_widget(item)

        # Atualiza ícone do botão de ordenação conforme estado atual
        self.sort_button.icon = "sort-ascending" if self.sort_asc else "sort-descending"

    # ---------- Callbacks (delegam p/ App) ----------
    def _exportar_impl(self):
        # Chama exportação na instância do App
        MDApp.get_running_app().exportar_resultados()

    def _copiar_ids_impl(self):
        # Chama cópia de IDs na instância do App
        MDApp.get_running_app().copiar_ids_resultados()

    def _corrigir_impl(self):
        # Abre diálogo de correção de coordenadas
        MDApp.get_running_app().abrir_corrigir_coordenadas()

    def _voltar_impl(self):
        # Troca para a tela "base"
        MDApp.get_running_app().root.current = "base"


# ------------------------------ Aplicativo principal ------------------------------
class AppMaps(MDApp):
    # Atributos de classe (estado compartilhado na instância)
    selected_file = None  # caminho da planilha escolhida
    last_directory = os.path.expanduser("~")  # último diretório acessado
    dialog = None  # referência ao diálogo atual (para fechar/abrir)
    df_temp = None  # usado para armazenar DataFrame temporário entre confirmações
    base_dict_temp = None  # dicionário de base para busca confiável (cache temporário)

    def build(self):
        """Configura tema, carrega .kv, cria e retorna o ScreenManager."""
        # ---- Aparência do tema ----
        self.title = "App Maps"
        self.theme_cls.theme_style = "Light"
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.primary_hue = "500"

        # ---- Infra (pastas e último diretório) ----
        self.carregar_ultimo_diretorio()
        self.configurar_pastas()

        # ---- Ícones e imagens do app ----
        self.logo_small = "ui/LogoFinal.png"
        self.logo_large = "ui/Imagem_LogoMFinal.png"
        self.placeholder = "ui/placeholder.png"

        # Seta ícone da janela (se existir); senão usa placeholder
        Window.set_icon(self.logo_small if os.path.exists(self.logo_small) else self.placeholder)

        # ---- Carrega arquivos KV de cada tela ----
        Builder.load_file("ui/home_screen.kv")
        Builder.load_file("ui/processar_planilha.kv")
        Builder.load_file("ui/base_screen.kv")

        # ---- Cria gerenciador de telas e registra as telas ----
        sm = MDScreenManager()
        sm.add_widget(HomeScreen(name="home"))
        sm.add_widget(ProcessarPlanilhaScreen(name="processar_planilha"))
        sm.add_widget(BaseScreen(name="base"))
        sm.add_widget(ResultadosBuscaScreen(name="resultados_busca"))
        return sm  # retorna o root widget do app

    # ------------------ Infra ------------------
    def configurar_pastas(self):
        """Garante a existência das pastas usadas pelo app."""
        os.makedirs("bases/pkl", exist_ok=True)
        os.makedirs("bases/xlsx", exist_ok=True)
        os.makedirs("logs", exist_ok=True)

    def carregar_ultimo_diretorio(self):
        """Lê do arquivo CONFIG_FILE o último diretório usado, se ainda existir."""
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                path = f.read().strip()
                if os.path.exists(path):
                    self.last_directory = path

    def salvar_ultimo_diretorio(self):
        """Persiste o último diretório acessado no seletor de arquivos."""
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(self.last_directory)

    # ---- Navegação entre telas ----
    def abrir_processar_planilha(self):
        self.root.current = "processar_planilha"

    def abrir_tratamento_base(self):
        self.root.current = "base"

    def voltar_home(self):
        self.root.current = "home"

    # ---- Seletor de planilha (.xlsx) ----
    def selecionar_planilha(self):
        """Abre o MDFileManager iniciando no último diretório usado."""
        initial_path = self.last_directory if os.path.exists(self.last_directory) else os.path.expanduser("~")

        # Cria e mostra o gerenciador de arquivos (SEU CÓDIGO original mantido)
        self.file_manager = MDFileManager(
            exit_manager=self.close_file_manager,  # callback ao fechar
            select_path=self.select_file,  # callback ao selecionar caminho
            preview=False,
            search="all",
            use_access=True,
            # background_color_toolbar="brown",
            background_color_selection_button=(0.1, 0.4, 0.8, 1),
            icon_selection_button="check",
            icon_color=(1, 0.6, 0, 1),
        )
        self.file_manager.show(initial_path)

        # >>> ÚNICA linha nova aqui: tenta aplicar estilo assim que abrir
        Clock.schedule_once(self._tweak_file_manager_ui, 0)

    # >>> NOVO: pequena rotina que tenta aplicar estilo no MDFileManager por até ~2s
    def _tweak_file_manager_ui(self, *args):
        self._fm_style_tries = 0
        Clock.schedule_interval(self._tentar_tweak_file_manager, 0.1)

    def _tentar_tweak_file_manager(self, dt):
        fm = getattr(self, "file_manager", None)
        if not fm:
            return True  # segue tentando

        self._fm_style_tries += 1
        if self._fm_style_tries > 20:  # ~2 segundos
            print("[FileManager] Não consegui aplicar estilo (timeout).")
            return False

        # --- localizar toolbar ---
        toolbar = None
        try:
            if hasattr(fm, "ids"):
                toolbar = fm.ids.get("toolbar") or fm.ids.get("appbar") or fm.ids.get("top_app_bar")
            if not toolbar:
                for w in fm.walk():
                    if w.__class__.__name__ in ("MDTopAppBar", "MDToolbar"):
                        toolbar = w
                        break
        except Exception:
            toolbar = None

        # --- localizar label do caminho ---
        path_label = None
        try:
            if hasattr(fm, "ids"):
                path_label = fm.ids.get("path") or fm.ids.get("current_path") or fm.ids.get("label")
            if not path_label:
                for w in fm.walk():
                    txt = getattr(w, "text", "")
                    if isinstance(txt, str) and ("\\" in txt or "/" in txt):
                        path_label = w
                        break
        except Exception:
            path_label = None

        tweaked = False

        # --- estiliza toolbar ---
        if toolbar:
            try:
                if hasattr(toolbar, "theme_bg_color"):
                    toolbar.theme_bg_color = "Custom"
                toolbar.md_bg_color = (0.13, 0.59, 0.95, 1)  # fundo
            except Exception:
                pass
            if hasattr(toolbar, "specific_text_color"):
                toolbar.specific_text_color = (0, 0, 0, 0)  # texto/ícones
            try:
                toolbar.height = dp(100)  # compacto
            except Exception:
                pass
            try:
                left = [["arrow-left", lambda x: fm.back()]]
                right = [["close", lambda x: self.close_file_manager()]]
                if hasattr(toolbar, "left_action_items") and not toolbar.left_action_items:
                    toolbar.left_action_items = left
                if hasattr(toolbar, "right_action_items") and not toolbar.right_action_items:
                    toolbar.right_action_items = right
            except Exception:
                pass
            tweaked = True

        # --- estiliza label do caminho ---
        if path_label:
            try:
                if hasattr(path_label, "font_style"):
                    path_label.font_style = "bodySmall"  # fonte menor
                if hasattr(path_label, "theme_text_color"):
                    path_label.theme_text_color = "Custom"
                    path_label.text_color = (1, 1, 1, 1)
                if hasattr(path_label, "shorten"):
                    path_label.shorten = True
                    path_label.shorten_from = "left"  # truncar pelo começo
            except Exception:
                pass
            tweaked = True

        # --- itens da lista mais baixos (quando possível) ---
        try:
            rv = getattr(fm, "ids", {}).get("recycleview") or getattr(fm, "ids", {}).get("rv")
            if rv and hasattr(rv, "data") and rv.data:
                data = rv.data
                for d in data:
                    d.setdefault("height", dp(56))
                rv.data = data
                tweaked = True
        except Exception:
            pass

        if tweaked:
            print("[FileManager] Estilo aplicado.")
            return False  # para o schedule_interval
        return True

    def close_file_manager(self, *args):
        """Fecha o seletor de arquivos se estiver aberto."""
        if hasattr(self, 'file_manager'):
            self.file_manager.close()

    def select_file(self, path):
        """Valida extensão .xlsx, atualiza estado e o label na tela de processamento."""
        if path.endswith(".xlsx"):
            self.selected_file = path
            self.last_directory = os.path.dirname(path)
            self.salvar_ultimo_diretorio()
            screen = self.root.get_screen("processar_planilha")
            screen.ids.file_label.text = f"Arquivo: {os.path.basename(path)}"
        else:
            self.show_dialog("Erro", "Por favor, selecione um arquivo válido com extensão .xlsx.")
        self.close_file_manager()

    # ------------------ Diálogo genérico ------------------
    def show_dialog(self, title, text):
        """
        Exibe um MDDialog padrão (KivyMD 2.x) com título, texto e botão Fechar.
        Reaproveita self.dialog para manter apenas um diálogo aberto por vez.
        """
        from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogSupportingText, MDDialogButtonContainer
        from kivymd.uix.button import MDButton, MDButtonText

        # Fecha diálogo anterior se existir
        if self.dialog:
            try:
                self.dialog.dismiss()
            except Exception:
                pass

        # Cria e abre novo diálogo
        self.dialog = MDDialog(
            MDDialogHeadlineText(text=title),
            MDDialogSupportingText(text=text),
            MDDialogButtonContainer(
                MDButton(
                    MDButtonText(text="Fechar"),
                    style="text",
                    on_release=lambda x: self.dialog.dismiss()
                )
            ),
        )
        self.dialog.open()
    # ------------------ Helpers de texto/endereço ------------------
    def normalizar_prefixo_logradouro(self, texto: str):
        """
        Normaliza prefixos de logradouro no INÍCIO da string para forma abreviada canônica,
        garantindo que sempre fiquem com inicial maiúscula (R., Av.)
        e evitando gerar 'R..' ou 'Av..'.
        """
        if not isinstance(texto, str):
            return texto
        t = texto.strip()

        # Rua -> R.  (variações: rua, r, r:, r., R, R:)
        t = re.sub(r'(?i)^\s*(?:rua|r:|r\.|r(?!\.)|R:|R(?!\.))\s*', 'R. ', t)

        # Avenida -> Av.  (variações: avenida, av, av:, av., AV)
        t = re.sub(r'(?i)^\s*(?:avenida|av:|av\.|av(?!\.)|AVENIDA|AV(?!\.))\s*', 'Av. ', t)

        # Compacta espaços duplos
        t = re.sub(r'\s{2,}', ' ', t)

        # Segurança extra: colapsa 'R..' -> 'R.' e 'Av..' -> 'Av.'
        t = re.sub(r'^(R|Av)\.\.', r'\1.', t)

        return t



    def capitalizar_endereco(self, texto):
        """Capitaliza endereços respeitando preposições comuns em português."""
        preposicoes = {'da', 'de', 'do', 'das', 'dos', 'e', 'em', 'no', 'na', 'nos', 'nas', 'a', 'o', 'as', 'os'}
        if not isinstance(texto, str):
            return texto
        palavras = texto.lower().split()
        if not palavras:
            return texto
        palavras[0] = palavras[0].capitalize()  # primeira palavra sempre capitalizada
        if len(palavras) > 1:
            palavras[-1] = palavras[-1].capitalize()  # última palavra capitalizada
        for i in range(1, len(palavras) - 1):
            if palavras[i] not in preposicoes:
                palavras[i] = palavras[i].capitalize()
        return ' '.join(palavras)

    def create_local_complement(self, row):
        """
        A partir das colunas padrão (Destination Address, City), separa:
          - Local: logradouro + número (+ cidade no formato " - Cidade" ou ", Cidade")
          - Complemento: restante do endereço após a segunda vírgula
        Retorna uma Series [Local, Complemento].
        """
        parts = str(row['Destination Address']).split(',')
        city = str(row.get('City', '')).strip()
        if len(parts) > 2:
            # pega os 2 primeiros fragmentos como "logradouro, número" e o resto vira complemento
            local = ', '.join(p.strip() for p in parts[:2]) + (' - ' + city if city else '')
            complemento = ', '.join(p.strip() for p in parts[2:])
        else:
            # quando não há mais de 2 partes, decide com base se termina em dígito (número)
            address = str(row['Destination Address']).strip()
            if address and address[-1].isdigit():
                local = address + (f' - {city}' if city else '')
            else:
                local = address + (f', {city}' if city else '')
            complemento = ""
        return pd.Series([local.strip(), complemento.strip()])

    def preencher_e_ordenar_sequence_stop(self, df):
        """
        Preenche valores ausentes de 'Sequence' e 'Stop' com sequência crescente
        baseada no maior valor existente e ordena pelo campo 'Sequence'.
        """
        df_copy = df.copy()

        def preencher_coluna(col):
            # Converte para string, usa '-' como placeholder, coleta válidos e completa a sequência
            col_str = [str(x).strip() if not pd.isna(x) else '-' for x in col]
            valores_validos = [int(x) for x in col_str if x.isdigit()]
            ultimo_valor = max(valores_validos) if valores_validos else 0
            for i in range(len(col_str)):
                if col_str[i] == '-':
                    ultimo_valor += 1
                    col_str[i] = str(ultimo_valor)
            return [int(x) for x in col_str]

        if 'Sequence' in df_copy.columns:
            df_copy['Sequence'] = preencher_coluna(df_copy['Sequence'])
        if 'Stop' in df_copy.columns:
            df_copy['Stop'] = preencher_coluna(df_copy['Stop'])

        # define colunas de ordenação (prioriza Sequence)
        order_cols = [c for c in ['Sequence'] if c in df_copy.columns]
        return df_copy.sort_values(by=order_cols, ascending=True).reset_index(drop=True)

    # ------------------ Google Maps helpers ------------------
    def inicializar_googlemaps(self):
        """Cria cliente Google Maps usando variável de ambiente com fallback em GOOGLE_API_KEY."""
        key = os.environ.get('GOOGLE_MAPS_API_KEY', GOOGLE_API_KEY)
        if not key or key == "SUA_CHAVE_API":
            raise ValueError("API Key do Google Maps não configurada.")
        return googlemaps.Client(key=key)

    def escolher_resultado_geocode(self, geocode_result, cidade):
        """
        Dentre os resultados de geocodificação, tenta priorizar aquele que casa com
        a cidade informada (locality ou administrative_area_level_2). Caso não haja
        match, retorna o primeiro resultado.
        """
        cidade = (cidade or "").strip().lower()
        if cidade:
            for result in geocode_result:
                for comp in result.get('address_components', []):
                    tipos = comp.get('types', [])
                    if ('locality' in tipos) or ('administrative_area_level_2' in tipos):
                        if cidade in comp.get('long_name', '').lower():
                            return result
        return geocode_result[0] if geocode_result else None

    def obter_dados_api_completos(self, gmaps, local, cidade):
        """
        Faz geocoding do 'local' e retorna dicionário com:
          Latitude, Longitude, Bairro_Maps, Cep_Maps, Cidade_Maps, Endereco_Formatado
        Usa 'cidade' como pista para escolher o melhor resultado.
        """
        try:
            resultados = gmaps.geocode(local)
            if not resultados:
                return None
            res = self.escolher_resultado_geocode(resultados, cidade)
            if not res:
                return None
            loc = res['geometry']['location']
            latitude = loc.get('lat')
            longitude = loc.get('lng')
            endereco_formatado = res.get('formatted_address', '')

            # Extrai componentes úteis do endereço
            bairro_maps = cep_maps = cidade_maps = None
            for comp in res.get('address_components', []):
                tipos = comp.get('types', [])
                if 'sublocality' in tipos or 'sublocality_level_1' in tipos or 'neighborhood' in tipos:
                    bairro_maps = comp.get('long_name')
                if 'postal_code' in tipos:
                    cep_maps = comp.get('long_name')
                if 'locality' in tipos:
                    cidade_maps = comp.get('long_name')
                if 'administrative_area_level_2' in tipos and not cidade_maps:
                    cidade_maps = comp.get('long_name')

            return {
                'Latitude': latitude,
                'Longitude': longitude,
                'Bairro_Maps': bairro_maps,
                'Cep_Maps': cep_maps,
                'Cidade_Maps': cidade_maps,
                'Endereco_Formatado': endereco_formatado
            }
        except Exception as e:
            print(f"Erro API para {local}: {e}")
            return None

    # ------------------ Utils Base ------------------
    def _ensure_base_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Garante que TODAS as colunas de BASE_COLUMNS existam na base (cria vazias se faltar)."""
        for col in BASE_COLUMNS:
            if col not in df.columns:
                df[col] = None
        return df[BASE_COLUMNS]

    def _sanitize_base(self, base_atual: pd.DataFrame) -> pd.DataFrame:
        """
        Remove duplicatas por 'Local' mantendo o menor 'ID'. Também ordena por ID.
        Útil antes de salvar a base para manter consistência.
        """
        base_atual = self._ensure_base_columns(base_atual)
        # Converte ID p/ numérico para ordenação correta
        base_atual['ID'] = pd.to_numeric(base_atual['ID'], errors='coerce')
        # Ordena por Local,ID para manter menor ID ao dropar duplicatas
        base_atual = base_atual.sort_values(by=['Local', 'ID'], ascending=[True, True])
        base_atual = base_atual.drop_duplicates(subset='Local', keep='first')
        base_atual = base_atual.sort_values(by='ID', ascending=True)
        return base_atual[BASE_COLUMNS]

    def _next_id(self, base_atual: pd.DataFrame) -> int:
        """Retorna próximo ID inteiro com base no máximo atual da base."""
        if 'ID' not in base_atual.columns or base_atual.empty:
            return 1
        max_id = pd.to_numeric(base_atual['ID'], errors='coerce').max()
        return int(max_id) + 1 if pd.notna(max_id) else 1

    def _coalesce(self, *vals):
        """
        Retorna o primeiro valor não nulo/não vazio da lista de valores informada.
        Útil para completar campos com a melhor origem disponível (planilha → API, por ex.).
        """
        for v in vals:
            if v is None:
                continue
            if isinstance(v, float) and pd.isna(v):
                continue
            sv = str(v).strip() if isinstance(v, str) else v
            if sv != '' and sv is not None:
                return sv
        return ''

    # ------------------ Processar Planilha ------------------
    def processar_planilha(self):
        """
        Lê a planilha selecionada, padroniza endereços, separa Local/Complemento,
        ajusta Sequence/Stop e salva versões em XLSX/PKL em bases/xlsx e bases/pkl.
        Também habilita botões de busca (rápida/confiável) na tela.
        """
        if not self.selected_file:
            self.show_dialog("Erro", "Nenhum arquivo selecionado para processamento.")
            return

        try:
            df = pd.read_excel(self.selected_file)  # lê XLSX
            # Normalização robusta do prefixo de logradouro (evita 'R..')
            if 'Destination Address' in df.columns:
                df['Destination Address'] = df['Destination Address'].apply(self.normalizar_prefixo_logradouro)
                df['Destination Address'] = df['Destination Address'].apply(self.capitalizar_endereco)
                df[['Local', 'Complemento']] = df.apply(self.create_local_complement, axis=1)
                # Garante também o prefixo correto em 'Local'
                df['Local'] = df['Local'].apply(self.normalizar_prefixo_logradouro)

            # Preenche Sequence/Stop e ordena
            df = self.preencher_e_ordenar_sequence_stop(df)

            # Define nomes de saída com base no nome do arquivo
            nome_base = os.path.splitext(os.path.basename(self.selected_file))[0]
            final_xlsx_path = os.path.join("bases/xlsx", f"{nome_base}_Final.xlsx")
            final_pkl_path = os.path.join("bases/pkl", f"{nome_base}_Base.pkl")

            # Salva arquivos finais
            df.to_excel(final_xlsx_path, index=False)
            df.to_pickle(final_pkl_path)

            # Habilita botões de busca na UI
            screen = self.root.get_screen("processar_planilha")
            screen.ids.btn_busca_rapida.disabled = False
            screen.ids.btn_busca_confiavel.disabled = False

            # Feedback p/ usuário
            self.show_dialog("Sucesso", f"Processamento concluído!\nArquivos gerados:\n{final_xlsx_path}\n{final_pkl_path}")

        except Exception as e:
            self.show_dialog("Erro", f"Erro no processamento:\n{str(e)}")

    # ------------------ Busca Rápida (não atualiza base) ------------------
    def buscar_latlong_rapida(self):
        """
        Lê a planilha atual, identifica linhas com Lat/Lon vazios/zero e pergunta
        se deseja consultar a API para preencher APENAS esses registros. Não toca
        na Base_Atual.pkl (gera um XLSX separado com os resultados).
        """
        if not self.selected_file:
            self.show_dialog("Erro", "Nenhuma planilha processada para buscar.")
            return

        try:
            df = pd.read_excel(self.selected_file)
            # Normalização robusta do prefixo de logradouro (evita 'R..')
            if 'Destination Address' in df.columns:
                df['Destination Address'] = df['Destination Address'].apply(self.normalizar_prefixo_logradouro)
                df['Destination Address'] = df['Destination Address'].apply(self.capitalizar_endereco)
                df[['Local', 'Complemento']] = df.apply(self.create_local_complement, axis=1)
                # Garante também o prefixo correto em 'Local'
                df['Local'] = df['Local'].apply(self.normalizar_prefixo_logradouro)

            # Conta quantas linhas precisam de API (lat/lon vazios/0)
            faltando_api = 0
            for _, row in df.iterrows():
                lat = row.get('Latitude', 0)
                lng = row.get('Longitude', 0)
                if pd.isna(lat) or pd.isna(lng) or lat == 0 or lng == 0:
                    faltando_api += 1

            if faltando_api > 0:
                # Guarda df temporário e abre diálogo de confirmação
                self.df_temp = df
                self.dialog = self.criar_dialogo_confirmacao_rapida(faltando_api)
                self.dialog.open()
                return

            # Se nada faltando, segue direto
            self.continuar_busca_rapida(df)

        except Exception as e:
            self.show_dialog("Erro", f"Erro na busca rápida:\n{str(e)}")

    def criar_dialogo_confirmacao_rapida(self, faltando):
        """Diálogo de confirmação para busca rápida informando quantos registros faltam."""
        from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogSupportingText, MDDialogButtonContainer
        from kivymd.uix.button import MDButton, MDButtonText

        return MDDialog(
            MDDialogHeadlineText(text="Confirmação"),
            MDDialogSupportingText(
                text=f"{faltando} registros com Latitude/Longitude zerados ou vazios.\nDeseja buscar na API do Google Maps?"
            ),
            MDDialogButtonContainer(
                MDButton(
                    MDButtonText(text="Cancelar"),
                    style="text",
                    on_release=lambda x: self.dialog.dismiss()
                ),
                MDButton(
                    MDButtonText(text="Sim"),
                    style="text",
                    on_release=lambda x: self.continuar_busca_rapida(self.df_temp)
                )
            ),
        )

    def continuar_busca_rapida(self, df):
        """Executa a busca rápida efetivamente e exporta um *_BuscaRapida.xlsx."""
        self.dialog.dismiss()  # fecha confirmação
        total = len(df)
        atualizados_api = 0

        # Acesso aos widgets de progresso na tela de processamento
        screen = self.root.get_screen("processar_planilha")
        progress = screen.ids.progress_bar
        status_label = screen.ids.status_label
        progress.value = 0

        gmaps = self.inicializar_googlemaps()  # cliente Google
        enderecos_api = []  # lista p/ log/relatório

        for idx, row in df.iterrows():
            lat = row.get('Latitude', 0)
            lng = row.get('Longitude', 0)

            if pd.isna(lat) or pd.isna(lng) or lat == 0 or lng == 0:
                # Consulta API para entradas incompletas
                dados = self.obter_dados_api_completos(gmaps, row['Local'], row.get('City', ''))
                if dados:
                    df.at[idx, 'Latitude'] = dados['Latitude']
                    df.at[idx, 'Longitude'] = dados['Longitude']
                    atualizados_api += 1
                    enderecos_api.append(row['Local'])
                    # Registra no histórico de chamadas de API
                    registrar_log_api(row['Local'], dados['Latitude'], dados['Longitude'], "rápida")

            # Atualiza barra de progresso e status
            progress.value = int((idx + 1) / total * 100)
            status_label.text = f"Processando: {idx+1}/{total}"

        # Exporta XLSX final da busca rápida
        nome_base = os.path.splitext(os.path.basename(self.selected_file))[0]
        final_xlsx_path = os.path.join("bases/xlsx", f"{nome_base}_Rap.xlsx")
        df.to_excel(final_xlsx_path, index=False)

        # Gera relatórios auxiliares dos endereços consultados
        with open("bases/xlsx/Enderecos_API.txt", 'w', encoding='utf-8') as f:
            for endereco in enderecos_api:
                f.write(endereco + '\n')
        pd.DataFrame({'Endereço API': enderecos_api}).to_excel("bases/xlsx/Enderecos_API.xlsx", index=False)

        self.show_dialog("Sucesso", f"Busca rápida concluída!\nAtualizados via API: {atualizados_api}\nArquivo salvo: {final_xlsx_path}")

    # ------------------ Busca Confiável (atualiza Base_Atual.pkl) ------------------
    def buscar_latlong_confiavel(self):
        """
        Usa/atualiza a Base_Atual.pkl. Para cada 'Local' da planilha:
          - Se já existir na base com Lat/Lon: usa e complementa campos vazios.
          - Senão: consulta API, cria/atualiza a base e exporta um *_BuscaConfiavel.xlsx.
        """
        if not self.selected_file:
            self.show_dialog("Erro", "Nenhuma planilha processada para buscar.")
            return

        try:
            df = pd.read_excel(self.selected_file)
            # Normalização robusta do prefixo de logradouro (evita 'R..')
            if 'Destination Address' in df.columns:
                df['Destination Address'] = df['Destination Address'].apply(self.normalizar_prefixo_logradouro)
                df['Destination Address'] = df['Destination Address'].apply(self.capitalizar_endereco)
                df[['Local', 'Complemento']] = df.apply(self.create_local_complement, axis=1)
                # Garante também o prefixo correto em 'Local'
                df['Local'] = df['Local'].apply(self.normalizar_prefixo_logradouro)

            # Carrega base atual (ou cria vazia)
            base_path = "bases/pkl/Base_Atual.pkl"
            if os.path.exists(base_path):
                base_atual = pd.read_pickle(base_path)
            else:
                base_atual = pd.DataFrame(columns=BASE_COLUMNS)
            base_atual = self._ensure_base_columns(base_atual)

            # Cria dicionário Local -> {Latitude, Longitude, ID} para consultas rápidas
            base_dict = base_atual.set_index('Local')[['Latitude', 'Longitude', 'ID']].to_dict('index') if not base_atual.empty else {}

            # Conta quantos já estão completos na base
            encontrados = sum(1 for loc in df['Local'] if loc in base_dict and
                              pd.notna(base_dict[loc].get('Latitude')) and pd.notna(base_dict[loc].get('Longitude')))
            faltando = len(df) - encontrados

            if faltando > 0:
                # Pede confirmação para chamar API e atualizar base
                self.df_temp = df
                self.base_dict_temp = base_dict
                self.dialog = self.criar_dialogo_confirmacao(faltando)
                self.dialog.open()
                return

            # Se nada faltando, segue direto sem confirmação
            self.continuar_busca_api(df, base_dict)

        except Exception as e:
            self.show_dialog("Erro", f"Erro na busca confiável:\n{str(e)}")

    def criar_dialogo_confirmacao(self, faltando):
        """Diálogo de confirmação para a busca confiável (atualiza Base_Atual.pkl)."""
        from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogSupportingText, MDDialogButtonContainer
        from kivymd.uix.button import MDButton, MDButtonText

        return MDDialog(
            MDDialogHeadlineText(text="Confirmação"),
            MDDialogSupportingText(
                text=f"{faltando} endereços não estão completos na base.\nDeseja buscar na API do Google Maps e atualizar a Base_Atual.pkl?"
            ),
            MDDialogButtonContainer(
                MDButton(MDButtonText(text="Cancelar"), style="text",
                         on_release=lambda x: self.dialog.dismiss()),
                MDButton(MDButtonText(text="Sim"), style="text",
                         on_release=lambda x: self.continuar_busca_api(self.df_temp, self.base_dict_temp)),
            ),
        )

    def continuar_busca_api(self, df, base_dict):
        """
        Fluxo principal da busca confiável:
          - Usa dados da base quando existir.
          - Consulta API quando necessário e atualiza base.
          - Exporta resultados e relatórios auxiliares.
        """
        self.dialog.dismiss()
        total = len(df)
        atualizados_api = 0
        encontrados_base = 0

        # Widgets de progresso na UI
        screen = self.root.get_screen("processar_planilha")
        progress = screen.ids.progress_bar
        status_label = screen.ids.status_label
        progress.value = 0

        # Carrega base atual (refresca do disco por segurança)
        base_path = "bases/pkl/Base_Atual.pkl"
        if os.path.exists(base_path):
            base_atual = pd.read_pickle(base_path)
        else:
            base_atual = pd.DataFrame(columns=BASE_COLUMNS)
        base_atual = self._ensure_base_columns(base_atual)

        gmaps = self.inicializar_googlemaps()
        enderecos_api = []

        # Itera sobre as linhas da planilha
        for idx, row in df.iterrows():
            local = row['Local']
            cidade = row.get('City', '')

            # Dados vindos da planilha (usados para completar base)
            bairro_plan = row.get('Bairro', '')
            cep_plan = row.get('Zipcode/Postal code', '') or row.get('CEP', '')
            nome_plan = row.get('Nome', '')
            tel_plan = row.get('Telefone', '')

            if local in base_dict and pd.notna(base_dict[local].get('Latitude')) and pd.notna(base_dict[local].get('Longitude')):
                # HIT na base (evita chamada de API)
                lat = base_dict[local]['Latitude']
                lng = base_dict[local]['Longitude']
                encontrados_base += 1

                # Completa campos vazios já existentes na base com dados da planilha
                base_idx = base_atual.index[base_atual['Local'] == local]
                if not base_idx.empty:
                    bi = base_idx[0]
                    base_atual.at[bi, 'Bairro'] = self._coalesce(base_atual.at[bi, 'Bairro'], bairro_plan)
                    base_atual.at[bi, 'CEP'] = self._coalesce(base_atual.at[bi, 'CEP'], cep_plan)
                    base_atual.at[bi, 'Nome'] = self._coalesce(base_atual.at[bi, 'Nome'], nome_plan)
                    base_atual.at[bi, 'Telefone'] = self._coalesce(base_atual.at[bi, 'Telefone'], tel_plan)
            else:
                # MISS na base → consulta API e atualiza
                dados = self.obter_dados_api_completos(gmaps, local, cidade)
                if dados:
                    lat = dados['Latitude']
                    lng = dados['Longitude']
                    atualizados_api += 1
                    enderecos_api.append(local)
                    registrar_log_api(local, lat, lng, "confiável")

                    base_idx = base_atual.index[base_atual['Local'] == local]
                    if not base_idx.empty:
                        # Atualiza linha existente
                        bi = base_idx[0]
                        base_atual.at[bi, 'Latitude'] = lat
                        base_atual.at[bi, 'Longitude'] = lng
                        base_atual.at[bi, 'Bairro'] = self._coalesce(base_atual.at[bi, 'Bairro'], bairro_plan, dados['Bairro_Maps'])
                        base_atual.at[bi, 'CEP'] = self._coalesce(base_atual.at[bi, 'CEP'], cep_plan, dados['Cep_Maps'])
                        base_atual.at[bi, 'Nome'] = self._coalesce(base_atual.at[bi, 'Nome'], nome_plan)
                        base_atual.at[bi, 'Telefone'] = self._coalesce(base_atual.at[bi, 'Telefone'], tel_plan)
                        base_atual.at[bi, 'Bairro_Maps'] = self._coalesce(base_atual.at[bi, 'Bairro_Maps'], dados['Bairro_Maps'])
                        base_atual.at[bi, 'Cep_Maps'] = self._coalesce(base_atual.at[bi, 'Cep_Maps'], dados['Cep_Maps'])
                        base_atual.at[bi, 'Cidade_Maps'] = self._coalesce(base_atual.at[bi, 'Cidade_Maps'], dados['Cidade_Maps'])
                        base_atual.at[bi, 'Endereco_Formatado'] = self._coalesce(base_atual.at[bi, 'Endereco_Formatado'], dados['Endereco_Formatado'])
                    else:
                        # Cria nova linha completa na base
                        novo_id = self._next_id(base_atual)
                        nova = {
                            'ID': novo_id,
                            'Local': local,
                            'Bairro': self._coalesce(bairro_plan, dados['Bairro_Maps']),
                            'CEP': self._coalesce(cep_plan, dados['Cep_Maps']),
                            'Nome': self._coalesce(nome_plan),
                            'Telefone': self._coalesce(tel_plan),
                            'Latitude': lat,
                            'Longitude': lng,
                            'Bairro_Maps': dados['Bairro_Maps'],
                            'Cep_Maps': dados['Cep_Maps'],
                            'Cidade_Maps': dados['Cidade_Maps'],
                            'Endereco_Formatado': dados['Endereco_Formatado']
                        }
                        base_atual = pd.concat([base_atual, pd.DataFrame([nova])], ignore_index=True)
                else:
                    # API sem resultado → deixa NaN
                    lat = None
                    lng = None

            # Atualiza lat/lon também na planilha corrente (resultado da execução)
            df.at[idx, 'Latitude'] = lat
            df.at[idx, 'Longitude'] = lng

            # Progresso UI
            progress.value = int((idx + 1) / total * 100)
            status_label.text = f"Processando: {idx+1}/{total}"

        # Sanitiza e salva base atualizada em múltiplos formatos
        base_atual = self._sanitize_base(base_atual)
        base_atual.to_pickle(base_path)
        base_atual.to_pickle("bases/pkl/Base_Atualizada.pkl")
        base_atual.to_excel("bases/xlsx/Base_Atualizada.xlsx", index=False)

        # Exporta resultado da planilha desta execução
        nome_base = os.path.splitext(os.path.basename(self.selected_file))[0]
        final_xlsx_path = os.path.join("bases/xlsx", f"{nome_base}_Conf.xlsx")
        df.to_excel(final_xlsx_path, index=False)

        # Lista de endereços efetivamente consultados na API nesta execução
        with open("bases/xlsx/Enderecos_API.txt", 'w', encoding='utf-8') as f:
            for endereco in enderecos_api:
                f.write(endereco + '\n')
        pd.DataFrame({'Endereço API': enderecos_api}).to_excel("bases/xlsx/Enderecos_API.xlsx", index=False)

        # Feedback
        self.show_dialog(
            "Sucesso",
            f"Busca confiável concluída!\nEncontrados na base: {encontrados_base}\nBuscados na API: {atualizados_api}\nArquivo salvo: {final_xlsx_path}"
        )

    # ------------------ Info Buscas ------------------
    def mostrar_info_busca(self):
        """Mostra diálogo explicando a diferença entre Busca Rápida e Confiável."""
        from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogSupportingText, MDDialogButtonContainer
        from kivymd.uix.button import MDButton, MDButtonText

        texto_info = (
            "🔎 Busca Rápida → mantém lat/long da planilha e usa API apenas para 0/vazios (não atualiza a base).\n\n"
            "🗂️ Busca Confiável → usa Base_Atual.pkl, consulta API quando necessário, cria/atualiza linhas completas\n"
            "com ID, dados da planilha (Bairro, CEP, Nome, Telefone) e dados da API (Bairro_Maps, Cep_Maps, Cidade_Maps, Endereco_Formatado)."
        )

        dialog = MDDialog(
            MDDialogHeadlineText(text="Informações sobre as buscas"),
            MDDialogSupportingText(text=texto_info),
            MDDialogButtonContainer(
                MDButton(
                    MDButtonText(text="Fechar"),
                    style="text",
                    on_release=lambda x: dialog.dismiss()
                ),
                spacing="8dp",
            ),
        )
        dialog.open()

    # ------------------ Duplicatas ------------------
    def remover_duplicados(self):
        """Abre Base_Atual.pkl, remove duplicatas por 'Local', salva XLSX e PKL sem duplicatas."""
        base_path = "bases/pkl/Base_Atual.pkl"

        if not os.path.exists(base_path):
            self.show_dialog("Erro", "Arquivo Base_Atual.pkl não foi encontrado.")
            return

        try:
            df = pd.read_pickle(base_path)
        except Exception as e:
            self.show_dialog("Erro ao carregar base", str(e))
            return

        total_antes = len(df)
        df = self._sanitize_base(df)  # remove duplicatas e ordena
        total_depois = len(df)
        duplicatas = total_antes - total_depois

        # Persiste mudanças e exporta relatório em Excel
        df.to_pickle(base_path)
        caminho_excel = base_path.replace(".pkl", "_sem_duplicatas.xlsx")
        df.to_excel(caminho_excel, index=False)

        self.show_dialog(
            "Concluído",
            f"Duplicatas removidas ({duplicatas}).\nArquivo Excel salvo:\n{caminho_excel}"
        )

    def ver_registros(self):
        """Exibe contagem total, faltando coordenadas e duplicados por Local na Base_Atual.pkl."""
        from kivymd.uix.dialog import MDDialog, MDDialogHeadlineText, MDDialogSupportingText, MDDialogButtonContainer
        from kivymd.uix.button import MDButton, MDButtonText
        import pandas as pd
        import os

        base_path = "bases/pkl/Base_Atual.pkl"
        if not os.path.exists(base_path):
            self.show_dialog("Erro", "Arquivo Base_Atual.pkl não foi encontrado.")
            return

        try:
            df = pd.read_pickle(base_path)
        except Exception as e:
            self.show_dialog("Erro ao carregar base", str(e))
            return

        total = len(df)
        # métricas úteis (com try/except para robustez)
        try:
            faltando_coords = df[['Latitude', 'Longitude']].isna().any(axis=1).sum()
        except Exception:
            faltando_coords = 0
        try:
            dups_local = df['Local'].duplicated().sum()
        except Exception:
            dups_local = 0

        texto = (
            f"Total de registros: {total}\n"
            f"Sem coordenadas: {faltando_coords}\n"
            f"Duplicados por Local: {dups_local}"
        )

        dialog = MDDialog(
            MDDialogHeadlineText(text=f"Base Atual"),
            MDDialogSupportingText(text=texto),
            MDDialogButtonContainer(
                MDButton(MDButtonText(text="Fechar"), style="text",
                         on_release=lambda x: dialog.dismiss())
            )
        )
        dialog.open()

    # ------------------ Buscar na Base (abre tela de resultados) ------------------
    def buscar_endereco(self):
        """
        Abre diálogo com campo de texto para o usuário digitar parte do endereço.
        Filtra a Base_Atual.pkl por 'Local' contendo o termo (case-insensitive)
        e mostra os resultados na tela ResultadosBuscaScreen.
        """
        from kivymd.uix.dialog import (
            MDDialog, MDDialogHeadlineText, MDDialogContentContainer, MDDialogButtonContainer
        )
        from kivymd.uix.button import MDButton, MDButtonText
        from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
        from kivy.uix.widget import Widget
        from kivy.clock import Clock

        base_path = "bases/pkl/Base_Atual.pkl"
        if not os.path.exists(base_path):
            self.show_dialog("Erro", "Arquivo Base_Atual.pkl não foi encontrado.")
            return
        try:
            df = pd.read_pickle(base_path)
        except Exception as e:
            self.show_dialog("Erro ao carregar base", str(e))
            return

        # Garante colunas necessárias para exibição
        for col in ['ID', 'Local', 'Latitude', 'Longitude', 'Complemento']:
            if col not in df.columns:
                df[col] = None

        # Campo de busca
        campo_busca = MDTextField(mode="outlined", size_hint_x=1)
        campo_busca.add_widget(MDTextFieldHintText(text="Digite parte do endereço"))

        def executar_busca(_):
            termo_busca = (campo_busca.text or "").strip()
            if not termo_busca:
                self.show_dialog("Aviso", "Digite um termo para buscar.")
                return

            # Filtra por Local contendo termo (case-insensitive)
            resultado = df[df['Local'].str.contains(termo_busca, case=False, na=False)][['ID', 'Local', 'Complemento', 'Latitude', 'Longitude']].copy()

            # Normaliza coluna 'ID' -> 'Id' (compatível com a tela de resultados)
            if not resultado.empty:
                resultado.rename(columns={'ID': 'Id'}, inplace=True)
                resultado['Id'] = pd.to_numeric(resultado['Id'], errors='coerce').fillna(0).astype(int)
                resultado = resultado.sort_values(by='Id', ascending=True).reset_index(drop=True)

            # Fecha diálogo e mostra a tela de resultados
            self.dialog.dismiss()
            tela = self.root.get_screen("resultados_busca")
            tela.set_results(resultado)
            self.root.current = "resultados_busca"

        # Monta diálogo
        self.dialog = MDDialog(
            MDDialogHeadlineText(text="Buscar endereço"),
            MDDialogContentContainer(
                campo_busca,
                orientation="vertical",
                spacing="12dp",
                padding=("16dp", "8dp", "16dp", "0dp"),
            ),
            MDDialogButtonContainer(
                Widget(),
                MDButton(MDButtonText(text="Cancelar"), style="text",
                         on_release=lambda x: self.dialog.dismiss()),
                MDButton(MDButtonText(text="Buscar"), style="filled",
                         on_release=executar_busca),
                spacing="8dp",
            ),
        )
        self.dialog.width_offset = "24dp"  # deixa o diálogo um pouco mais estreito
        self.dialog.open()
        # Foca automaticamente o campo de texto quando abrir
        Clock.schedule_once(lambda *_: setattr(campo_busca, "focus", True), 0.1)

    # ------------------ Exportação / Cópia / Correção ------------------
    def exportar_resultados(self):
        """Exporta os resultados exibidos na tela de resultados para um XLSX com timestamp."""
        from datetime import datetime
        tela = self.root.get_screen("resultados_busca")
        df = getattr(tela, 'df_resultados', None)
        if df is None or df.empty:
            self.show_dialog("Aviso", "Não há resultados para exportar.")
            return
        os.makedirs("bases/xlsx", exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho = os.path.join("bases/xlsx", f"Resultados_Busca_{ts}.xlsx")
        try:
            df.to_excel(caminho, index=False)
            self.show_dialog("Sucesso", f"Resultados exportados para:\n{caminho}")
        except Exception as e:
            self.show_dialog("Erro ao exportar", str(e))

    def copiar_ids_resultados(self):
        """Copia todos os IDs (coluna 'Id') exibidos na tela de resultados para a área de transferência."""
        tela = self.root.get_screen("resultados_busca")
        df = getattr(tela, 'df_resultados', None)
        if df is None or df.empty:
            self.show_dialog("Aviso", "Não há IDs para copiar.")
            return
        if 'Id' not in df.columns:
            self.show_dialog("Aviso", "Coluna 'Id' não encontrada nos resultados.")
            return
        ids_text = '\n'.join(str(i) for i in df['Id'].tolist())
        try:
            Clipboard.copy(ids_text)
            self.show_dialog("Copiado", f"{len(df)} ID(s) copiados para a área de transferência.")
        except Exception as e:
            self.show_dialog("Erro", f"Não foi possível copiar: {e}")

    def abrir_corrigir_coordenadas(self):
        """
        Abre diálogo para corrigir Latitude/Longitude por:
          - ID (preferencial) ou
          - Local (igualdade exata)
        Salva a base pós-ajuste (PKL + XLSX auxiliar).
        """
        from kivymd.uix.dialog import (
            MDDialog, MDDialogHeadlineText, MDDialogContentContainer, MDDialogButtonContainer
        )
        from kivymd.uix.button import MDButton, MDButtonText
        from kivymd.uix.textfield import MDTextField, MDTextFieldHintText
        from kivy.uix.widget import Widget

        base_path = "bases/pkl/Base_Atual.pkl"
        if not os.path.exists(base_path):
            self.show_dialog("Erro", "Arquivo Base_Atual.pkl não foi encontrado.")
            return

        # Campos do formulário
        campo_id = MDTextField(mode="outlined")
        campo_id.add_widget(MDTextFieldHintText(text="ID (numérico)"))
        campo_local = MDTextField(mode="outlined")
        campo_local.add_widget(MDTextFieldHintText(text="Local (igualdade exata)"))
        campo_lat = MDTextField(mode="outlined")
        campo_lat.add_widget(MDTextFieldHintText(text="Nova Latitude (ex.: -22.123456)"))
        campo_lon = MDTextField(mode="outlined")
        campo_lon.add_widget(MDTextFieldHintText(text="Nova Longitude (ex.: -47.654321)"))

        def aplicar(_):
            # Lê e normaliza entradas
            id_txt = (campo_id.text or "").strip()
            local_txt = (campo_local.text or "").strip()
            lat_txt = (campo_lat.text or "").strip().replace(",", ".")
            lon_txt = (campo_lon.text or "").strip().replace(",", ".")

            if not lat_txt or not lon_txt:
                self.show_dialog("Aviso", "Informe Latitude e Longitude.")
                return

            # Validação numérica
            try:
                lat_v = float(lat_txt)
                lon_v = float(lon_txt)
            except ValueError:
                self.show_dialog("Erro", "Latitude/Longitude inválidas.")
                return

            # Carrega base
            try:
                dfb = pd.read_pickle(base_path)
            except Exception as e:
                self.show_dialog("Erro", f"Falha ao carregar a base: {e}")
                return

            # Detecta nome real da coluna de ID
            id_col = next((c for c in ["Id", "ID", "id", "iD"] if c in dfb.columns), None)

            atualizados = 0
            if id_txt:
                # Preferência por ajuste via ID
                if id_col is None:
                    self.show_dialog("Erro", "A base não possui coluna de ID.")
                    return
                try:
                    id_num = int(float(id_txt))
                except ValueError:
                    self.show_dialog("Erro", "ID deve ser numérico.")
                    return
                idx = dfb.index[dfb[id_col] == id_num]
                if not idx.empty:
                    dfb.loc[idx, 'Latitude'] = lat_v
                    dfb.loc[idx, 'Longitude'] = lon_v
                    atualizados = len(idx)
                else:
                    self.show_dialog("Aviso", f"ID {id_num} não encontrado.")
                    return
            else:
                # Fallback por Local exato
                if not local_txt:
                    self.show_dialog("Aviso", "Preencha ID ou Local.")
                    return
                if 'Local' not in dfb.columns:
                    self.show_dialog("Erro", "A base não possui coluna 'Local'.")
                    return
                idx = dfb.index[dfb['Local'] == local_txt]
                if not idx.empty:
                    dfb.loc[idx, 'Latitude'] = lat_v
                    dfb.loc[idx, 'Longitude'] = lon_v
                    atualizados = len(idx)
                else:
                    self.show_dialog("Aviso", f"Local '{local_txt}' não encontrado.")
                    return

            try:
                # Sanitiza e salva base + exporta XLSX auxiliar
                dfb = self._sanitize_base(dfb)
                dfb.to_pickle(base_path)
                dfb.to_excel(base_path.replace(".pkl", "_Atualizada.xlsx"), index=False)
            except Exception as e:
                self.show_dialog("Erro", f"Falha ao salvar base: {e}")
                return

            self.dialog.dismiss()
            self.show_dialog("Concluído", f"Atualizado(s) {atualizados} registro(s). Base salva.")

        # Monta diálogo
        self.dialog = MDDialog(
            MDDialogHeadlineText(text="Corrigir Latitude/Longitude"),
            MDDialogContentContainer(
                campo_id,
                campo_local,
                campo_lat,
                campo_lon,
                orientation="vertical",
                spacing="12dp",
                padding=("16dp", "8dp", "16dp", "0dp"),
            ),
            MDDialogButtonContainer(
                Widget(),
                MDButton(MDButtonText(text="Cancelar"), style="text",
                         on_release=lambda x: self.dialog.dismiss()),
                MDButton(MDButtonText(text="Aplicar"), style="filled",
                         on_release=aplicar),
                spacing="8dp",
            ),
        )
        self.dialog.width_offset = "24dp"
        self.dialog.open()

    # Wrappers para manter compatibilidade com os .kv
    def buscar_rapida(self):
        """Chama a busca rápida (compat c/ on_release no .kv)."""
        return self.buscar_latlong_rapida()

    def buscar_confiavel(self):
        """Chama a busca confiável (compat c/ on_release no .kv)."""
        return self.buscar_latlong_confiavel()

    def abrir_main_rotas(self):
        """Abre o script main_rotas.py em um processo separado usando o mesmo Python."""
        try:
            import os
            base_dir = os.path.dirname(os.path.abspath(__file__))  # pasta deste arquivo
            script_path = os.path.join(base_dir, "main_rotas.py")  # caminho esperado
            if not os.path.exists(script_path):
                raise FileNotFoundError(f"Arquivo não encontrado: {script_path}")
            subprocess.Popen([sys.executable, script_path])  # inicia processo
            print("[OK] main_rotas.py iniciado em um novo processo.")
        except Exception as e:
            print(f"[ERRO] Não foi possível abrir main_rotas.py: {e}")

# ------------------------------ Chamando Whatsa envio ------------------------------

            # main_whatsapp.py

    def abrir_main_whatsapp(self):
        """Abre o script main_whatsapp.py em um processo separado usando o mesmo Python."""
        try:
            import os
            base_dir = os.path.dirname(os.path.abspath(__file__))  # pasta deste arquivo
            script_path = os.path.join(base_dir, "main_whatsapp.py")  # caminho esperado
            if not os.path.exists(script_path):
                raise FileNotFoundError(f"Arquivo não encontrado: {script_path}")
            subprocess.Popen([sys.executable, script_path])  # inicia processo
            print("[OK] main_whatsapp.py iniciado em um novo processo.")
        except Exception as e:
            print(f"[ERRO] Não foi possível abrir main_whatsapp.py: {e}")


# ------------------------------ Logs de chamadas à API ------------------------------
def registrar_log_api(endereco, latitude, longitude, tipo_busca):
    """
    Acrescenta (ou cria) o arquivo logs/Enderecos_API_Historico.xlsx com colunas:
      Id (incremental), Data/Hora, Endereço, Latitude, Longitude, Tipo ("rápida"/"confiável").

    Correção: tolera arquivo inexistente, cabeçalho diferente (ID/id) e calcula próximo Id com segurança.
    """
    from datetime import datetime

    caminho_log = "logs/Enderecos_API_Historico.xlsx"
    colunas_padrao = ["Id", "Data/Hora", "Endereço", "Latitude", "Longitude", "Tipo"]
    os.makedirs("logs", exist_ok=True)

    # Lê o log, se existir; caso contrário, cria DF vazio com colunas padrão
    if os.path.exists(caminho_log) and os.path.getsize(caminho_log) > 0:
        try:
            df_log = pd.read_excel(caminho_log)
        except Exception:
            df_log = pd.DataFrame(columns=colunas_padrao)
    else:
        df_log = pd.DataFrame(columns=colunas_padrao)

    # Normaliza a coluna Id (aceita variações 'ID', 'id', etc.)
    if "Id" not in df_log.columns:
        id_alt = next((c for c in ["ID", "id", "iD"] if c in df_log.columns), None)
        if id_alt:
            df_log.rename(columns={id_alt: "Id"}, inplace=True)
        else:
            df_log["Id"] = pd.Series(dtype="Int64")

    # Garante todas as colunas padrão presentes
    for c in colunas_padrao:
        if c not in df_log.columns:
            df_log[c] = None
    df_log = df_log[colunas_padrao]

    # Calcula próximo Id de forma robusta
    try:
        mx = pd.to_numeric(df_log["Id"], errors="coerce").max()
    except Exception:
        mx = None
    proximo_id = (int(mx) + 1) if (mx is not None and pd.notna(mx)) else 1

    novo = {
        "Id": proximo_id,
        "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Endereço": endereco,
        "Latitude": latitude,
        "Longitude": longitude,
        "Tipo": tipo_busca  # "rápida" ou "confiável"
    }

    df_log = pd.concat([df_log, pd.DataFrame([novo])], ignore_index=True)
    try:
        df_log.to_excel(caminho_log, index=False)
    except Exception as e:
        print(f"[LOG_API] Falha ao salvar log: {e}")


# ------------------------------ Main guard ------------------------------
if __name__ == "__main__":
    # Inicia o aplicativo KivyMD
    AppMaps().run()