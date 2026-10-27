import sys
import os
import pandas as pd
import locale
import json
import re
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QStackedWidget, QLabel, 
                               QFrame, QLineEdit, QFileDialog, QTableWidget, QHeaderView,
                               QTableWidgetItem, QMessageBox, QTextEdit, QComboBox)
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt

DIR_GLOBAL = "GLOBAL"
DIR_IMPORT = "IMPORT"
DIR_EXPORT = "EXPORT"
ARQUIVO_HISTORICO = os.path.join(DIR_GLOBAL, "historico_pendencias.csv")
ARQUIVO_RESUMO = os.path.join(DIR_GLOBAL, "dados_app.json")

for d in [DIR_GLOBAL, DIR_IMPORT, DIR_EXPORT]:
    os.makedirs(d, exist_ok=True)

# Helper function to reliably get resource path for PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class CodureApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CODURE - Controle de Duplicidade de Registros")
        self.showMaximized()
        
        # Seta o ícone na janela principal
        icone_path = resource_path("icone.png")
        if os.path.exists(icone_path):
            self.setWindowIcon(QIcon(icone_path))
        
        try:
            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_ALL, '')
            except:
                pass
        
        self.dados_app = {
            'competencia_atual': '',
            'historicos': {}
        }
        
        self.dados_bpa_limpo = []
        self.dados_excluidos = []
        self.dados_conferencia = []
        self.cabecalho_bpa = []
        
        self.configurar_estilos()
        
        widget_central = QWidget()
        layout_principal = QHBoxLayout(widget_central)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(0)
        self.setCentralWidget(widget_central)

        sidebar = QFrame()
        sidebar.setFixedWidth(250)
        sidebar.setStyleSheet("background-color: #1A1A1A; border-right: 1px solid #333333;")
        layout_sidebar = QVBoxLayout(sidebar)
        layout_sidebar.setContentsMargins(0, 20, 0, 0)
        layout_sidebar.setSpacing(10)

        # Container para o título e ícone na sidebar
        h_titulo = QHBoxLayout()
        h_titulo.setContentsMargins(15, 0, 15, 0)
        
        if os.path.exists(icone_path):
            lbl_icone = QLabel()
            pixmap = QPixmap(icone_path).scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            lbl_icone.setPixmap(pixmap)
            h_titulo.addWidget(lbl_icone)
            
        titulo = QLabel("CODURE")
        titulo.setStyleSheet("font-size: 24px; font-weight: bold; color: #4CAF50;")
        h_titulo.addWidget(titulo)
        h_titulo.addStretch()
        
        layout_sidebar.addLayout(h_titulo)
        layout_sidebar.addSpacing(20)

        self.btn_etapa1 = QPushButton("1. Configuração Global")
        self.btn_etapa2 = QPushButton("2. Triagem do Mês")
        self.btn_etapa3 = QPushButton("3. Gestão e Repescagem")
        self.btn_etapa4 = QPushButton("4. Resumo Consolidado")
        
        for btn in [self.btn_etapa1, self.btn_etapa2, self.btn_etapa3, self.btn_etapa4]:
            btn.setObjectName("btn_sidebar")
            layout_sidebar.addWidget(btn)
        
        layout_sidebar.addStretch()

        self.paginas = QStackedWidget()
        self.paginas.addWidget(self.criar_pagina_configuracao())
        self.paginas.addWidget(self.criar_pagina_processamento())
        self.paginas.addWidget(self.criar_pagina_gestao())
        self.paginas.addWidget(self.criar_pagina_resumo())

        self.btn_etapa1.clicked.connect(lambda: self.paginas.setCurrentIndex(0))
        self.btn_etapa2.clicked.connect(lambda: self.paginas.setCurrentIndex(1))
        self.btn_etapa3.clicked.connect(lambda: self.paginas.setCurrentIndex(2))
        self.btn_etapa4.clicked.connect(lambda: self.paginas.setCurrentIndex(3))

        layout_principal.addWidget(sidebar)
        layout_principal.addWidget(self.paginas)
        
        self.carregar_dados_app()
        self.carregar_historico_tabela()

    def configurar_estilos(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #121212; }
            QLabel { color: #E0E0E0; font-size: 14px; }
            h2 { color: #4CAF50; }
            QPushButton { background-color: #1F1F1F; color: #FFFFFF; border: 1px solid #333333; padding: 10px; border-radius: 4px; font-weight: bold; }
            QPushButton:hover { background-color: #333333; }
            QPushButton:disabled { background-color: #0F0F0F; color: #555555; border: 1px solid #222222; }
            QPushButton#btn_sidebar { background-color: transparent; border: none; text-align: left; padding-left: 15px; font-size: 14px; }
            QPushButton#btn_sidebar:hover { background-color: #2C2C2C; border-left: 4px solid #4CAF50; }
            QPushButton#btn_action { background-color: #2196F3; font-size: 14px; padding: 15px; margin-top: 5px; }
            QPushButton#btn_action:hover { background-color: #1976D2; }
            QPushButton#btn_export { background-color: #4CAF50; font-size: 13px; padding: 10px; }
            QPushButton#btn_export:hover { background-color: #388E3C; }
            QPushButton#btn_warning { background-color: #FF9800; font-size: 14px; padding: 15px; margin-top: 10px; }
            QPushButton#btn_warning:hover { background-color: #F57C00; }
            QLineEdit, QComboBox { background-color: #1E1E1E; color: #FFFFFF; border: 1px solid #444444; padding: 8px; border-radius: 4px; font-size: 14px; }
            QComboBox::drop-down { border: none; }
            QComboBox QAbstractItemView { background-color: #1E1E1E; color: #FFFFFF; selection-background-color: #4CAF50; }
            QTableWidget { background-color: #1E1E1E; color: #E0E0E0; gridline-color: #333333; border: 1px solid #333333; }
            QHeaderView::section { background-color: #2C2C2C; color: #FFFFFF; padding: 4px; border: 1px solid #333333; }
            QTextEdit { background-color: #1E1E1E; color: #00FF00; font-family: Consolas, monospace; border: 1px solid #333333; padding: 10px; }
        """)

    def salvar_dados_app(self):
        try:
            with open(ARQUIVO_RESUMO, 'w', encoding='utf-8') as f:
                json.dump(self.dados_app, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Erro ao salvar dados: {e}")

    def carregar_dados_app(self):
        if os.path.exists(ARQUIVO_RESUMO):
            try:
                with open(ARQUIVO_RESUMO, 'r', encoding='utf-8') as f:
                    self.dados_app.update(json.load(f))
                self.input_competencia.setText(self.dados_app.get('competencia_atual', ''))
                self.atualizar_combo_competencias()
            except Exception as e:
                print(f"Erro ao carregar dados: {e}")

    def sanitizar_cpf(self, valor):
        if pd.isna(valor): return ""
        v_str = str(valor)
        numeros = re.sub(r'\D', '', v_str)
        if not numeros: return ""
        return numeros.zfill(11)
        
    def sanitizar_codigo(self, valor, tamanho):
        if pd.isna(valor): return ""
        v_str = str(valor).strip()
        if v_str.endswith('.0'): v_str = v_str[:-2]
        return v_str.zfill(tamanho)

    def formatar_moeda(self, valor):
        try:
            return locale.currency(valor, symbol=False, grouping=True)
        except Exception:
            return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def calcular_competencia_limite(self, competencia_origem):
        ano = int(competencia_origem[:4])
        mes = int(competencia_origem[4:6])
        mes += 3
        if mes > 12:
            mes -= 12
            ano += 1
        return f"{ano}{mes:02d}"

    def carregar_tabela_sigtap(self):
        filepath = self.input_sigtap.text()
        if not filepath:
            raise Exception("Tabela SIGTAP não configurada na Etapa 1.")
        df_sigtap = pd.read_excel(filepath)
        df_sigtap.columns = [c.split(',')[0].strip() for c in df_sigtap.columns]
        
        pa_id = df_sigtap['PA_ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(9)
        pa_dv = df_sigtap['PA_DV'].astype(str).str.replace(r'\.0$', '', regex=True)
        codigos = pa_id + pa_dv
        valores = pd.to_numeric(df_sigtap['PA_SA'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
        return dict(zip(codigos, valores))

    def salvar_config_global(self):
        comp = self.input_competencia.text().strip()
        if not comp or len(comp) != 6 or not comp.isdigit():
            QMessageBox.warning(self, "Aviso", "Digite uma competência válida no formato AAAAMM.")
            return
        self.dados_app['competencia_atual'] = comp
        self.salvar_dados_app()
        QMessageBox.information(self, "Sucesso", "Configurações Globais salvas com sucesso!")

    def atualizar_combo_competencias(self):
        self.combo_comp.clear()
        comps = list(self.dados_app.get('historicos', {}).keys())
        comps.sort(reverse=True)
        if comps:
            self.combo_comp.addItems(comps)
            comp_atual = self.dados_app.get('competencia_atual')
            if comp_atual in comps:
                self.combo_comp.setCurrentText(comp_atual)
        self.atualizar_painel_resumo()

    def atualizar_painel_resumo(self, comp_selecionada=None):
        comp = self.combo_comp.currentText()
        if not comp or comp not in self.dados_app.get('historicos', {}):
            self.lbl_apresentado.setText("Total Apresentado (Prestador): 0 proc. | R$ 0,00")
            self.lbl_retido.setText("Possíveis OCI (Retidos): 0 proc. | R$ 0,00")
            self.lbl_baixado.setText("OCIs Concluídas (Baixados): 0 proc. | R$ 0,00")
            self.lbl_liberado.setText("Total Liberado (BPA Limpo): 0 proc. | R$ 0,00")
            return
            
        dados = self.dados_app['historicos'][comp]
        
        def mk_texto(prefix, cat):
            qtd = dados[cat]['Qtd']
            valor = self.formatar_moeda(dados[cat]['Valor'])
            return f"{prefix} {qtd} proc. | R$ {valor}"
            
        self.lbl_apresentado.setText(mk_texto("Total Apresentado (Prestador):", "Apresentado"))
        self.lbl_retido.setText(mk_texto("Possíveis OCI (Retidos):", "Retido"))
        self.lbl_baixado.setText(mk_texto("OCIs Concluídas (Baixados):", "Baixado"))
        self.lbl_liberado.setText(mk_texto("Total Liberado (BPA Limpo):", "Liberado"))

    def avancar_competencia(self):
        comp = self.input_competencia.text().strip()
        if not comp or len(comp) != 6 or not comp.isdigit():
            QMessageBox.warning(self, "Aviso", "A competência na Etapa 1 está inválida.")
            return
        
        ano = int(comp[:4])
        mes = int(comp[4:6])
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
        
        nova_comp = f"{ano}{mes:02d}"
        self.input_competencia.setText(nova_comp)
        self.dados_app['competencia_atual'] = nova_comp
        self.salvar_dados_app()
        QMessageBox.information(self, "Avanço Automático", f"Competência global avançada para o próximo mês:\n{nova_comp}")

    def extrair_malha_planilha(self, df_sec):
        malha = {}
        chaves_concluidas = set()
        
        col_status = 'STATUS OCI' if 'STATUS OCI' in df_sec.columns else 'STATUS'
        col_exames = 'EXAMES REALIZADOS (PRESTADOR)' if 'EXAMES REALIZADOS (PRESTADOR)' in df_sec.columns else 'PROCEDIMENTO'
        
        if 'CPF' not in df_sec.columns or col_status not in df_sec.columns or col_exames not in df_sec.columns:
            raise Exception("A planilha não possui as colunas necessárias (CPF, STATUS OCI/STATUS, EXAMES REALIZADOS/PROCEDIMENTO).")

        for _, row in df_sec.iterrows():
            cpf = self.sanitizar_cpf(row['CPF'])
            if not cpf: continue
            
            status_raw = str(row[col_status]).strip().upper().replace('Í', 'I').replace(' ', '_')
            exames_brutos = str(row[col_exames])
            codigos_encontrados = re.findall(r'\b\d{10}\b', exames_brutos)
            
            for cod in codigos_encontrados:
                malha[(cpf, cod)] = status_raw
                # FIX: Aceita CONCLUIDO ou OCI_CONCLUIDA
                if status_raw.startswith('CONCLUID') or status_raw.startswith('OCI_CONCLUID'):
                    chaves_concluidas.add((cpf, cod))
                    
        return malha, chaves_concluidas

    def executar_triagem(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        if not comp_atual:
            QMessageBox.warning(self, "Aviso", "A competência não foi salva. Volte na Etapa 1 e clique em 'Salvar Configurações'.")
            return

        arq_prestador = self.input_txt_bpa.text()
        arq_sec = self.input_planilha_sec_mes.text()

        if not all([arq_prestador, arq_sec]):
            QMessageBox.warning(self, "Aviso", "Selecione o TXT do Prestador e a Planilha da Secretaria.")
            return

        self.console_triagem.append("Iniciando triagem em memória...")
        QApplication.processEvents()

        try:
            tabela_precos = self.carregar_tabela_sigtap()
            df_sec = pd.read_excel(arq_sec)
            
            self.console_triagem.append("Mapeando exames via Regex na Malha OCI...")
            malha_oci, _ = self.extrair_malha_planilha(df_sec)

            self.dados_app['historicos'][comp_atual] = {
                'Apresentado': {'Qtd': 0, 'Valor': 0.0},
                'Liberado': {'Qtd': 0, 'Valor': 0.0},
                'Retido': {'Qtd': 0, 'Valor': 0.0},
                'Baixado': {'Qtd': 0, 'Valor': 0.0},
                'arquivo_origem': os.path.basename(arq_prestador)
            }
            resumo = self.dados_app['historicos'][comp_atual]

            self.dados_bpa_limpo = []
            self.dados_excluidos = []
            self.dados_conferencia = []
            self.cabecalho_bpa = []
            registros_quarentena = []

            with open(arq_prestador, 'r', encoding='latin-1') as txt_in:
                for linha in txt_in:
                    if linha.startswith("01"):
                        nova_linha = linha[:7] + comp_atual + linha[13:]
                        self.cabecalho_bpa.append(nova_linha)
                        continue
                    elif linha.startswith("02"):
                        self.cabecalho_bpa.append(linha)
                        continue
                    
                    if linha.startswith("03"):
                        comp_realizacao = linha[9:15].strip()
                        procedimento = linha[49:59].strip()
                        qtd_str = linha[88:94].strip()
                        quantidade = int(qtd_str) if qtd_str.isdigit() else 1
                        cpf = linha[338:349].strip()
                        nome = linha[112:142].strip()
                        
                        valor = tabela_precos.get(procedimento, 0.0)
                        resumo['Apresentado']['Qtd'] += 1
                        resumo['Apresentado']['Valor'] += valor
                        
                        chave = (cpf, procedimento)
                        status_final_registro = "BPA-I"
                        
                        if chave in malha_oci:
                            status_malha = malha_oci[chave]
                            status_final_registro = status_malha
                            
                            self.dados_excluidos.append({
                                'Competencia_Realizacao': comp_realizacao,
                                'CPF': cpf, 'Nome': nome, 'Procedimento': procedimento,
                                'Valor': valor, 'Status_Origem': status_malha
                            })

                            # FIX: Match flexivel
                            if status_malha.startswith('POSSIVEL'):
                                resumo['Retido']['Qtd'] += 1
                                resumo['Retido']['Valor'] += valor
                                registros_quarentena.append({
                                    'CPF': cpf, 'Procedimento': procedimento,
                                    'Competencia_Origem': comp_realizacao,
                                    'Competencia_Limite': self.calcular_competencia_limite(comp_realizacao),
                                    'Status': 'AGUARDANDO_OCI', 'Linha_Raw': linha.strip('\n')
                                })
                            elif status_malha.startswith('CONCLUID') or status_malha.startswith('OCI_CONCLUID'):
                                resumo['Baixado']['Qtd'] += 1
                                resumo['Baixado']['Valor'] += valor
                        else:
                            resumo['Liberado']['Qtd'] += 1
                            resumo['Liberado']['Valor'] += valor
                            self.dados_bpa_limpo.append(linha)
                        
                        self.dados_conferencia.append({
                            'Competencia_Realizacao': comp_realizacao,
                            'CPF': cpf, 'Nome': nome, 'Procedimento': procedimento,
                            'Quantidade': quantidade, 'Status': status_final_registro
                        })

            if registros_quarentena:
                df_novos = pd.DataFrame(registros_quarentena)
                if not os.path.exists(ARQUIVO_HISTORICO):
                    df_novos.to_csv(ARQUIVO_HISTORICO, index=False, sep=';', encoding='utf-8')
                else:
                    df_novos.to_csv(ARQUIVO_HISTORICO, mode='a', index=False, header=False, sep=';', encoding='utf-8')
                self.carregar_historico_tabela()

            self.salvar_dados_app()
            self.atualizar_combo_competencias()
            
            self.btn_export_bpa.setEnabled(True)
            self.btn_export_excluidos.setEnabled(True)
            self.btn_export_conf.setEnabled(True)

            self.console_triagem.append("Triagem finalizada em memória! Arquivo histórico (CSV) foi atualizado.\n-> Utilize os botões abaixo para gerar os relatórios na pasta EXPORT.")
            QMessageBox.information(self, "Triagem Finalizada", "Processamento em memória concluído! Você já pode gerar os arquivos.")

        except Exception as e:
            self.console_triagem.append(f"ERRO: {str(e)}")
            QMessageBox.critical(self, "Erro", f"Erro no processamento: {str(e)}")

    def gerar_bpa_limpo(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        arq_saida = os.path.join(DIR_EXPORT, f"BPA_LIMPO_{comp_atual}.TXT")
        with open(arq_saida, 'w', encoding='latin-1') as f:
            f.writelines(self.cabecalho_bpa)
            f.writelines(self.dados_bpa_limpo)
        self.console_triagem.append(f"[GERADO] -> {os.path.basename(arq_saida)}")
        QMessageBox.information(self, "Sucesso", f"BPA Limpo salvo na pasta EXPORT:\n{os.path.basename(arq_saida)}")

    def gerar_relatorio_excluidos(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        if not self.dados_excluidos:
            QMessageBox.information(self, "Aviso", "Nenhum registro foi excluído nesta triagem.")
            return
        arq_saida = os.path.join(DIR_EXPORT, f"RELATORIO_EXCLUIDOS_{comp_atual}.csv")
        pd.DataFrame(self.dados_excluidos).to_csv(arq_saida, index=False, sep=';', encoding='utf-8-sig')
        self.console_triagem.append(f"[GERADO] -> {os.path.basename(arq_saida)}")
        QMessageBox.information(self, "Sucesso", f"Relatório de Excluídos salvo na pasta EXPORT:\n{os.path.basename(arq_saida)}")

    def gerar_relatorio_conferencia(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        arq_saida = os.path.join(DIR_EXPORT, f"CONFERENCIA_MANUAL_{comp_atual}.csv")
        pd.DataFrame(self.dados_conferencia).to_csv(arq_saida, index=False, sep=';', encoding='utf-8-sig')
        self.console_triagem.append(f"[GERADO] -> {os.path.basename(arq_saida)}")
        QMessageBox.information(self, "Sucesso", f"Relatório de Conferência salvo na pasta EXPORT:\n{os.path.basename(arq_saida)}")

    def carregar_historico_tabela(self):
        if os.path.exists(ARQUIVO_HISTORICO):
            try:
                df = pd.read_csv(ARQUIVO_HISTORICO, sep=';', dtype=str).fillna('')
                self.tabela_historico.setRowCount(len(df))
                for i, row in df.iterrows():
                    self.tabela_historico.setItem(i, 0, QTableWidgetItem(str(row.get('CPF', ''))))
                    self.tabela_historico.setItem(i, 1, QTableWidgetItem(str(row.get('Procedimento', ''))))
                    self.tabela_historico.setItem(i, 2, QTableWidgetItem(str(row.get('Competencia_Origem', ''))))
                    self.tabela_historico.setItem(i, 3, QTableWidgetItem(str(row.get('Competencia_Limite', ''))))
                    self.tabela_historico.setItem(i, 4, QTableWidgetItem(str(row.get('Status', ''))))
                
                if hasattr(self, 'input_filtro'):
                    self.filtrar_tabela_historico(self.input_filtro.text())
            except Exception as e:
                print(f"Erro ao carregar histórico: {e}")

    def filtrar_tabela_historico(self, texto):
        texto = texto.lower().strip()
        for row in range(self.tabela_historico.rowCount()):
            match = False
            if not texto:
                match = True
            else:
                for col in range(self.tabela_historico.columnCount()):
                    item = self.tabela_historico.item(row, col)
                    if item and texto in item.text().lower():
                        match = True
                        break
            self.tabela_historico.setRowHidden(row, not match)

    def atualizar_historico_via_planilha(self):
        arq_sec_extra = self.input_planilha_sec_extra.text()
        if not arq_sec_extra:
            QMessageBox.warning(self, "Aviso", "Selecione a Planilha de Atualização OCI.")
            return
            
        if not os.path.exists(ARQUIVO_HISTORICO):
            QMessageBox.information(self, "Aviso", "Não há registros retidos no histórico para atualizar.")
            return
            
        try:
            df_sec = pd.read_excel(arq_sec_extra)
            _, chaves_concluidas = self.extrair_malha_planilha(df_sec)
            
            if not chaves_concluidas:
                QMessageBox.information(self, "Aviso", "A planilha não possui registros de OCI_CONCLUIDA válidos.")
                return

            df_hist = pd.read_csv(ARQUIVO_HISTORICO, sep=';', dtype=str)
            atualizados = 0
            
            for index, row in df_hist.iterrows():
                if row['Status'] == 'AGUARDANDO_OCI':
                    chave_hist = (self.sanitizar_cpf(row['CPF']), self.sanitizar_codigo(row['Procedimento'], 10))
                    if chave_hist in chaves_concluidas:
                        df_hist.at[index, 'Status'] = 'OCI_CONCLUIDA'
                        atualizados += 1
            
            if atualizados > 0:
                df_hist.to_csv(ARQUIVO_HISTORICO, index=False, sep=';', encoding='utf-8')
                self.carregar_historico_tabela()
                QMessageBox.information(self, "Sucesso", f"O histórico foi atualizado!\n{atualizados} registros marcados como OCI_CONCLUIDA.")
            else:
                QMessageBox.information(self, "Informação", "Nenhum registro pendente no CSV foi afetado por esta planilha.")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar via planilha: {str(e)}")

    def executar_repescagem(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        arq_prestador = self.input_txt_bpa.text()

        if not comp_atual or not arq_prestador:
            QMessageBox.warning(self, "Aviso", "O 'BPA Bruto' e a Competência Atual precisam estar preenchidos (Etapas 1 e 2) para base do cabeçalho.")
            return
            
        if not os.path.exists(ARQUIVO_HISTORICO):
            QMessageBox.information(self, "Aviso", "Nenhum histórico encontrado para repescagem.")
            return

        try:
            df_hist = pd.read_csv(ARQUIVO_HISTORICO, sep=';', dtype=str)
            mascara = (df_hist['Status'] == 'AGUARDANDO_OCI') & (df_hist['Competencia_Limite'] <= comp_atual)
            df_vencidos = df_hist[mascara]

            if df_vencidos.empty:
                QMessageBox.information(self, "Info", "Nenhum procedimento atrasado apto para repescagem neste mês.")
                return
            
            tabela_precos = self.carregar_tabela_sigtap()
            qtd_repescado = 0
            valor_total_repescado = 0.0

            cabecalho = []
            with open(arq_prestador, 'r', encoding='latin-1') as txt_in:
                for linha in txt_in:
                    if linha.startswith("01"):
                        nova_linha = linha[:7] + comp_atual + linha[13:]
                        cabecalho.append(nova_linha)
                    elif linha.startswith("02"):
                        cabecalho.append(linha)
                    else:
                        break
            
            arq_saida = os.path.join(DIR_EXPORT, f"BPA_REPESCAGEM_{comp_atual}.TXT")
            with open(arq_saida, 'w', encoding='latin-1') as txt_out:
                txt_out.writelines(cabecalho)
                for index, row in df_vencidos.iterrows():
                    txt_out.write(str(row['Linha_Raw']) + '\n')
                    df_hist.at[index, 'Status'] = 'BPA_REPESCADO'
                    
                    procedimento = self.sanitizar_codigo(row['Procedimento'], 10)
                    valor_total_repescado += tabela_precos.get(procedimento, 0.0)
                    qtd_repescado += 1
            
            df_hist.to_csv(ARQUIVO_HISTORICO, index=False, sep=';', encoding='utf-8')
            self.carregar_historico_tabela()
            
            val_format = self.formatar_moeda(valor_total_repescado)
            mensagem = f"Repescagem concluída na pasta EXPORT!\n\nFísico: {qtd_repescado} procedimentos\nFinanceiro: R$ {val_format}"
            QMessageBox.information(self, "Sucesso", mensagem)

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro na repescagem: {str(e)}")

    def baixar_extrato_txt(self):
        comp = self.combo_comp.currentText()
        if not comp or comp not in self.dados_app.get('historicos', {}):
            QMessageBox.warning(self, "Aviso", "Nenhum processamento encontrado para gerar extrato.")
            return
            
        dados = self.dados_app['historicos'][comp]
        arquivo_base = dados.get('arquivo_origem', 'N/A')
        data_geracao = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        
        nome_extrato = os.path.join(DIR_EXPORT, f"EXTRATO_BPA_OCI_{comp}.TXT")
        
        conteudo = f"=== CODURE - RESUMO DE FATURAMENTO ===\n"
        conteudo += f"Data da Geração: {data_geracao}\n"
        conteudo += f"Competência Visualizada: {comp}\n"
        conteudo += f"Arquivo Base Importado: {arquivo_base}\n"
        conteudo += "="*65 + "\n\n"
        
        for k in ['Apresentado', 'Retido', 'Baixado', 'Liberado']:
            v = dados[k]
            label = k.ljust(15)
            qtd = str(v['Qtd']).zfill(4)
            valor = self.formatar_moeda(v['Valor']).rjust(12)
            conteudo += f"| {label} | Físico: {qtd} un. | Financeiro: R$ {valor} |\n"
                
        conteudo += "\n" + "="*65 + "\n"
        
        with open(nome_extrato, "w", encoding="utf-8") as f:
            f.write(conteudo)
            
        QMessageBox.information(self, "Sucesso", f"Extrato da competência {comp} salvo na pasta EXPORT:\n{os.path.basename(nome_extrato)}")

    def criar_pagina_configuracao(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        
        layout.addWidget(QLabel("<h2>Etapa 1: Parâmetros Globais</h2>"))
        layout.addSpacing(20)
        
        h_comp = QHBoxLayout()
        lbl_comp = QLabel("Competência Atual de Submissão (AAAAMM):")
        lbl_comp.setFixedWidth(280)
        h_comp.addWidget(lbl_comp)
        
        self.input_competencia = QLineEdit()
        self.input_competencia.setText(self.dados_app.get('competencia_atual', ''))
        self.input_competencia.setPlaceholderText("Ex: 202610")
        self.input_competencia.setFixedWidth(150)
        h_comp.addWidget(self.input_competencia)
        h_comp.addStretch()
        layout.addLayout(h_comp)
        
        layout.addSpacing(10)
        self.input_sigtap = self.criar_linha_arquivo(layout, "Tabela de Valores (SIGTAP Ex-DBF):", DIR_GLOBAL)
        
        layout.addSpacing(20)
        btn_salvar = QPushButton("SALVAR CONFIGURAÇÕES")
        btn_salvar.setObjectName("btn_action")
        btn_salvar.setFixedWidth(250)
        btn_salvar.clicked.connect(self.salvar_config_global)
        layout.addWidget(btn_salvar)
        
        layout.addStretch()
        return page

    def criar_pagina_processamento(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        
        layout.addWidget(QLabel("<h2>Etapa 2: Triagem e Geração de Arquivos</h2>"))
        layout.addWidget(QLabel("Triagem em memória e geração modular de relatórios (Salvos na pasta EXPORT)"))
        layout.addSpacing(10)
        
        self.input_txt_bpa = self.criar_linha_arquivo(layout, "BPA Bruto do Mês (TXT):", DIR_IMPORT)
        self.input_planilha_sec_mes = self.criar_linha_arquivo(layout, "Malha OCI da Secretaria (Excel):", DIR_IMPORT)
        
        h_botoes = QHBoxLayout()
        
        self.btn_rodar_triagem = QPushButton("1. EXECUTAR TRIAGEM (Memória)")
        self.btn_rodar_triagem.setObjectName("btn_action")
        self.btn_rodar_triagem.clicked.connect(self.executar_triagem)
        h_botoes.addWidget(self.btn_rodar_triagem)
        
        self.btn_export_bpa = QPushButton("2. GERAR BPA LIMPO")
        self.btn_export_bpa.setObjectName("btn_export")
        self.btn_export_bpa.setEnabled(False)
        self.btn_export_bpa.clicked.connect(self.gerar_bpa_limpo)
        h_botoes.addWidget(self.btn_export_bpa)
        
        self.btn_export_excluidos = QPushButton("3. GERAR REL. EXCLUÍDOS")
        self.btn_export_excluidos.setObjectName("btn_export")
        self.btn_export_excluidos.setEnabled(False)
        self.btn_export_excluidos.clicked.connect(self.gerar_relatorio_excluidos)
        h_botoes.addWidget(self.btn_export_excluidos)
        
        self.btn_export_conf = QPushButton("4. GERAR CONFERÊNCIA MANUAL")
        self.btn_export_conf.setObjectName("btn_export")
        self.btn_export_conf.setEnabled(False)
        self.btn_export_conf.clicked.connect(self.gerar_relatorio_conferencia)
        h_botoes.addWidget(self.btn_export_conf)
        
        layout.addLayout(h_botoes)
        
        layout.addWidget(QLabel("Console de Processamento:"))
        self.console_triagem = QTextEdit()
        self.console_triagem.setReadOnly(True)
        layout.addWidget(self.console_triagem)
        return page

    def criar_pagina_gestao(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        
        h_title = QHBoxLayout()
        h_title.addWidget(QLabel("<h2>Etapa 3: Gestão de Banco de Dados e Atrasados</h2>"))
        btn_atualizar = QPushButton("Atualizar Tabela CSV (Pós-Edição Manual)")
        btn_atualizar.setFixedWidth(280)
        btn_atualizar.clicked.connect(self.carregar_historico_tabela)
        h_title.addWidget(btn_atualizar)
        layout.addLayout(h_title)
        
        layout.addWidget(QLabel("A) Atualização Manual de OCI (Caso a secretaria mande planilhas de OCIs concluídas em lote)"))
        self.input_planilha_sec_extra = self.criar_linha_arquivo(layout, "Planilha Avulsa OCI (Excel):", DIR_IMPORT)
        btn_atualizar_oci = QPushButton("Atualizar Status (Malha Avulsa)")
        btn_atualizar_oci.setObjectName("btn_action")
        btn_atualizar_oci.setFixedWidth(280)
        btn_atualizar_oci.clicked.connect(self.atualizar_historico_via_planilha)
        layout.addWidget(btn_atualizar_oci)
        
        layout.addSpacing(20)
        layout.addWidget(QLabel("B) Histórico de Procedimentos Retidos no CSV:"))
        
        self.input_filtro = QLineEdit()
        self.input_filtro.setPlaceholderText("🔎 Filtrar tabela por CPF, Procedimento, Competência ou Status (Ex: AGUARDANDO_OCI)...")
        self.input_filtro.textChanged.connect(self.filtrar_tabela_historico)
        layout.addWidget(self.input_filtro)
        
        self.tabela_historico = QTableWidget(0, 5)
        self.tabela_historico.setHorizontalHeaderLabels(["CPF", "Procedimento", "Origem", "Limite", "Status"])
        self.tabela_historico.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.tabela_historico)
        
        self.btn_gerar_repescagem = QPushButton("GERAR BPA DE REPESCAGEM NA PASTA EXPORT (Prazos Estourados)")
        self.btn_gerar_repescagem.setObjectName("btn_warning")
        self.btn_gerar_repescagem.clicked.connect(self.executar_repescagem)
        layout.addWidget(self.btn_gerar_repescagem)
        
        return page

    def criar_pagina_resumo(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        
        h_filtro = QHBoxLayout()
        h_filtro.addWidget(QLabel("<h2>Etapa 4: Extrato Físico-Financeiro</h2>"))
        h_filtro.addStretch()
        h_filtro.addWidget(QLabel("Visualizar Competência:"))
        
        self.combo_comp = QComboBox()
        self.combo_comp.setFixedWidth(120)
        self.combo_comp.currentTextChanged.connect(self.atualizar_painel_resumo)
        h_filtro.addWidget(self.combo_comp)
        layout.addLayout(h_filtro)
        
        self.lbl_apresentado = QLabel("")
        self.lbl_retido = QLabel("")
        self.lbl_baixado = QLabel("")
        self.lbl_liberado = QLabel("")
        
        for lbl in [self.lbl_apresentado, self.lbl_retido, self.lbl_baixado, self.lbl_liberado]:
            lbl.setStyleSheet("font-size: 16px; padding: 15px; background-color: #2C2C2C; border-radius: 4px; margin-bottom: 10px; font-family: monospace;")
            layout.addWidget(lbl)
            
        layout.addSpacing(20)
        
        h_botoes = QHBoxLayout()
        h_botoes.setSpacing(20) # Add some space between buttons
        
        btn_baixar = QPushButton("DOWNLOAD DO EXTRATO (TXT)")
        btn_baixar.setObjectName("btn_action")
        # Ajuste de tamanho forçado via estilo (igualando visualmente na proporcao 50/50 com o da direita)
        btn_baixar.setStyleSheet("QPushButton { background-color: #2196F3; font-size: 14px; padding: 15px; } QPushButton:hover { background-color: #1976D2; }")
        btn_baixar.clicked.connect(self.baixar_extrato_txt)
        h_botoes.addWidget(btn_baixar, 1) # weight 1
        
        btn_proxima = QPushButton("AVANÇAR PARA A PRÓXIMA COMPETÊNCIA (MÊS SEGUINTE)")
        btn_proxima.setObjectName("btn_warning")
        btn_proxima.setStyleSheet("QPushButton { background-color: #FF9800; font-size: 14px; padding: 15px; } QPushButton:hover { background-color: #F57C00; }")
        btn_proxima.clicked.connect(self.avancar_competencia)
        h_botoes.addWidget(btn_proxima, 1) # weight 1
        
        layout.addLayout(h_botoes)
        layout.addStretch()
        return page

    def criar_linha_arquivo(self, layout_pai, texto_label, pasta_inicial=""):
        h_layout = QHBoxLayout()
        lbl = QLabel(texto_label)
        lbl.setFixedWidth(280)
        h_layout.addWidget(lbl)
        
        input_file = QLineEdit()
        input_file.setReadOnly(True)
        btn_procurar = QPushButton("Localizar")
        btn_procurar.setFixedWidth(100)
        
        default_dir = os.path.join(os.getcwd(), pasta_inicial)
        btn_procurar.clicked.connect(lambda: input_file.setText(QFileDialog.getOpenFileName(self, "Selecionar Arquivo", default_dir)[0]))
        
        h_layout.addWidget(input_file)
        h_layout.addWidget(btn_procurar)
        layout_pai.addLayout(h_layout)
        return input_file

if __name__ == "__main__":
    import ctypes
    try:
        myappid = 'franca.saude.codure.v10' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = QApplication(sys.argv)
    janela = CodureApp()
    janela.show()
    sys.exit(app.exec())
