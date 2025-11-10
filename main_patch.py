
# -*- coding: utf-8 -*-
"""
main_patch.py
Aplica métodos em AppMaps sem precisar reescrever seu main.py.
Uso no final do main.py (após a definição da classe AppMaps):
    from main_patch import apply as _apply_patch
    _apply_patch(AppMaps)
"""

import re
import pandas as pd

def apply(AppMapsClass):
    # --- Helpers ---
    def normalizar_prefixo_logradouro(self, texto: str):
        if not isinstance(texto, str):
            return texto
        t = texto.strip()
        t = re.sub(r'(?i)^\s*(?:rua|r:|r\.|r(?!\.)|R:|R(?!\.))\s*', 'R. ', t)
        t = re.sub(r'(?i)^\s*(?:avenida|av:|av\.|av(?!\.)|AVENIDA|AV)\s*', 'Av. ', t)
        t = re.sub(r'\s{2,}', ' ', t)
        t = re.sub(r'^(R|Av)\.\.', r'\1.', t)
        return t

    def normalizar_telefone(self, valor):
        if valor is None:
            return pd.NA
        try:
            if pd.isna(valor):
                return pd.NA
        except Exception:
            pass
        s = str(valor).strip()
        if s == '' or s.lower() in {'nan', 'none'}:
            return pd.NA
        s = re.sub(r'\D+', '', s)
        return pd.NA if s == '' else s

    # --- Processamento ---
    def processar_planilha(self):
        if not getattr(self, "selected_file", None):
            try:
                self.show_dialog("Erro", "Nenhuma planilha selecionada. Clique em 'Importar base...' primeiro.")
            except Exception:
                print("[processar_planilha] Nenhuma planilha selecionada.")
            return

        try:
            df = pd.read_excel(self.selected_file)

            # Telefones
            for col in df.columns:
                lc = str(col).strip().lower()
                if lc in {'telefone','telefone 1','telefone1','phone','phone 1','phone1','celular','celular 1','celular1'}:
                    df[col] = df[col].apply(self.normalizar_telefone)
                    try:
                        df[col] = df[col].astype('string')
                    except Exception:
                        pass

            # Endereços
            if 'Destination Address' in df.columns:
                df['Destination Address'] = df['Destination Address'].apply(self.normalizar_prefixo_logradouro)
                if hasattr(self, 'capitalizar_endereco'):
                    df['Destination Address'] = df['Destination Address'].apply(self.capitalizar_endereco)
                if hasattr(self, 'create_local_complement'):
                    df[['Local', 'Complemento']] = df.apply(self.create_local_complement, axis=1)
                if 'Local' in df.columns:
                    df['Local'] = df['Local'].apply(self.normalizar_prefixo_logradouro)

            # Ordem/sequence se existir utilitário
            if hasattr(self, 'preencher_e_ordenar_sequence_stop'):
                try:
                    df = self.preencher_e_ordenar_sequence_stop(df)
                except Exception as e:
                    print("[processar_planilha] aviso preencher_e_ordenar_sequence_stop:", e)

            self.df_temp = df

            # UI
            try:
                screen = self.root.get_screen("processar_planilha")
                if hasattr(screen, 'ids') and 'status_label' in screen.ids:
                    screen.ids.status_label.text = f"Planilha processada • {len(df)} linhas"
                if hasattr(screen, 'ids') and 'btn_busca_rapida' in screen.ids:
                    screen.ids.btn_busca_rapida.disabled = False
                if hasattr(screen, 'ids') and 'btn_busca_confiavel' in screen.ids:
                    screen.ids.btn_busca_confiavel.disabled = False
            except Exception as e:
                print("[processar_planilha] aviso UI:", e)

            try:
                self.show_dialog("Pronto", f"Processamento concluído. Linhas: {len(df)}")
            except Exception:
                print("[processar_planilha] concluído:", len(df))

        except Exception as e:
            try:
                self.show_dialog("Erro", f"Falha ao processar planilha:\n{e}")
            except Exception:
                print("[processar_planilha] ERRO:", e)
            return

    def processar_planilha_btn(self, *args, **kwargs):
        return self.processar_planilha()

    # Bind
    AppMapsClass.normalizar_prefixo_logradouro = normalizar_prefixo_logradouro
    AppMapsClass.normalizar_telefone = normalizar_telefone
    AppMapsClass.processar_planilha = processar_planilha
    AppMapsClass.processar_planilha_btn = processar_planilha_btn

