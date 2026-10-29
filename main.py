import sys
import os
import pandas as pd
import openpyxl
import locale
import json
import re
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QStackedWidget, QLabel, 
                               QFrame, QLineEdit, QFileDialog, QTableWidget, QHeaderView,
                               QTableWidgetItem, QMessageBox, QTextEdit, QComboBox, QListWidget, QListWidgetItem)
from PySide6.QtGui import QIcon, QPixmap, QColor
from PySide6.QtCore import Qt

DIR_GLOBAL = "GLOBAL"
DIR_IMPORT = "IMPORT"
DIR_EXPORT = "EXPORT"
ARQUIVO_HISTORICO = os.path.join(DIR_GLOBAL, "historico_pendencias.csv")
ARQUIVO_RESUMO = os.path.join(DIR_GLOBAL, "dados_app.json")
CODIGO_DENSITOMETRIA = "0204060028"

for d in [DIR_GLOBAL, DIR_IMPORT, DIR_EXPORT]:
    os.makedirs(d, exist_ok=True)

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
        self.dados_conferencia_rx = []
        self.dados_conferencia_denso = []
        self.cabecalho_bpa = []
        
        self.buffer_pacientes = {}
        self.conflitos_pendentes = {}
        self.status_conflitos = {}
        self.linhas_processadas = []
        self.registros_quarentena = []
        
        self.configurar_estilos()
        
        widget_central = QWidget()
        layout_principal = QHBoxLayout(widget_central)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(0)
        self.setCentralWidget(widget_central)

        sidebar = QFrame()
        sidebar.setFixedWidth(280)
        sidebar.setStyleSheet("background-color: #1A1A1A; border-right: 1px solid #333333;")
        layout_sidebar = QVBoxLayout(sidebar)
        layout_sidebar.setContentsMargins(0, 20, 0, 0)
        layout_sidebar.setSpacing(10)

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
        self.btn_etapa3 = QPushButton("3. Resolução de Conflitos")
        self.btn_etapa4 = QPushButton("4. Gestão e Repescagem")
        self.btn_etapa5 = QPushButton("5. Resumo Consolidado")
        
        for btn in [self.btn_etapa1, self.btn_etapa2, self.btn_etapa3, self.btn_etapa4, self.btn_etapa5]:
            btn.setObjectName("btn_sidebar")
            layout_sidebar.addWidget(btn)
        
        layout_sidebar.addStretch()

        self.paginas = QStackedWidget()
        self.paginas.addWidget(self.criar_pagina_configuracao())
        self.paginas.addWidget(self.criar_pagina_processamento())
        self.paginas.addWidget(self.criar_pagina_resolucao())
        self.paginas.addWidget(self.criar_pagina_gestao())
        self.paginas.addWidget(self.criar_pagina_resumo())

        self.btn_etapa1.clicked.connect(lambda: self.paginas.setCurrentIndex(0))
        self.btn_etapa2.clicked.connect(lambda: self.paginas.setCurrentIndex(1))
        self.btn_etapa3.clicked.connect(lambda: self.paginas.setCurrentIndex(2))
        self.btn_etapa4.clicked.connect(lambda: self.paginas.setCurrentIndex(3))
        self.btn_etapa5.clicked.connect(lambda: self.paginas.setCurrentIndex(4))

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
            QTableWidget, QListWidget { background-color: #1E1E1E; color: #E0E0E0; gridline-color: #333333; border: 1px solid #333333; }
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
            for lbl in [self.lbl_apresentado, self.lbl_retido, self.lbl_baixado, self.lbl_liberado]:
                lbl.setText("Nenhum dado encontrado.")
            return
            
        dados = self.dados_app['historicos'][comp]
        
        def mk_texto(cat, titulo, show_denso):
            t = dados[cat]['Total']
            r = dados[cat]['RX']
            d = dados[cat]['DENSO']
            
            base = f"{titulo}\nTOTAL: {t['Qtd']} proc. | R$ {self.formatar_moeda(t['Valor'])}\n"
            if show_denso:
                base += f" ├─ RX: {r['Qtd']} un. | R$ {self.formatar_moeda(r['Valor'])}\n"
                base += f" └─ DENSO: {d['Qtd']} un. | R$ {self.formatar_moeda(d['Valor'])}"
            else:
                base += f" └─ RX: {r['Qtd']} un. | R$ {self.formatar_moeda(r['Valor'])}"
            return base
            
        self.lbl_apresentado.setText(mk_texto("Apresentado", "Apresentado (Prestador)", True))
        self.lbl_retido.setText(mk_texto("Retido", "Possíveis OCI (Retidos)", False))
        self.lbl_baixado.setText(mk_texto("Baixado", "OCIs Concluídas (Baixados)", False))
        self.lbl_liberado.setText(mk_texto("Liberado", "Total Liberado (BPA Limpo)", True))

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
        return malha

    def executar_triagem(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        if not comp_atual:
            QMessageBox.warning(self, "Aviso", "A competência não foi salva na Etapa 1.")
            return

        arq_prestador = self.input_txt_bpa.text()
        arq_sec = self.input_planilha_sec_mes.text()

        if not all([arq_prestador, arq_sec]):
            QMessageBox.warning(self, "Aviso", "Selecione o TXT do Prestador e a Planilha da Secretaria.")
            return

        self.console_triagem.append("Iniciando triagem em memória e agrupando CPFs...")
        QApplication.processEvents()

        try:
            self.tabela_precos = self.carregar_tabela_sigtap()
            df_sec = pd.read_excel(arq_sec)
            malha_oci = self.extrair_malha_planilha(df_sec)

            base_dict = lambda: {'Total': {'Qtd': 0, 'Valor': 0.0}, 'RX': {'Qtd': 0, 'Valor': 0.0}, 'DENSO': {'Qtd': 0, 'Valor': 0.0}}
            self.dados_app['historicos'][comp_atual] = {
                'Apresentado': base_dict(),
                'Liberado': base_dict(),
                'Retido': base_dict(),
                'Baixado': base_dict(),
                'arquivo_origem': os.path.basename(arq_prestador)
            }

            self.buffer_pacientes = {}
            self.cabecalho_bpa = []
            
            with open(arq_prestador, 'r', encoding='latin-1') as txt_in:
                for linha in txt_in:
                    if linha.startswith("01"):
                        self.cabecalho_bpa.append(linha[:7] + comp_atual + linha[13:])
                        continue
                    elif linha.startswith("02"):
                        self.cabecalho_bpa.append(linha)
                        continue
                    elif linha.startswith("03"):
                        cpf = linha[338:349].strip()
                        comp_realizacao = linha[9:15].strip()
                        data_raw = linha[36:44].strip()
                        
                        if len(data_raw) == 8 and data_raw.isdigit():
                            data_sort = f"{data_raw[4:8]}{data_raw[2:4]}{data_raw[0:2]}"
                            data_fmt = f"{data_raw[6:8]}/{data_raw[4:6]}/{data_raw[0:4]}"
                        else:
                            data_sort = comp_realizacao + "00"
                            data_fmt = comp_realizacao
                            
                        registro = {
                            'linha_raw': linha,
                            'cpf': cpf,
                            'proc': linha[49:59].strip(),
                            'nome': linha[112:142].strip(),
                            'comp_realizacao': comp_realizacao,
                            'data_sort': data_sort,
                            'data_fmt': data_fmt,
                            'valor': self.tabela_precos.get(linha[49:59].strip(), 0.0)
                        }
                        
                        if cpf not in self.buffer_pacientes:
                            self.buffer_pacientes[cpf] = []
                        self.buffer_pacientes[cpf].append(registro)

            self.conflitos_pendentes = {}
            self.linhas_processadas = []
            self.status_conflitos = {}

            for cpf, registros in self.buffer_pacientes.items():
                candidatos_oci = []
                for reg in registros:
                    chave = (cpf, reg['proc'])
                    if reg['proc'] == CODIGO_DENSITOMETRIA:
                        reg['status_malha'] = 'BPA-I'
                    else:
                        if chave in malha_oci:
                            status_malha = malha_oci[chave]
                            reg['status_malha'] = status_malha
                            if status_malha.startswith('CONCLUID') or status_malha.startswith('OCI_CONCLUID'):
                                candidatos_oci.append(reg)
                        else:
                            reg['status_malha'] = 'BPA-I'
                        
                if len(candidatos_oci) > 1:
                    self.conflitos_pendentes[cpf] = registros
                    self.status_conflitos[cpf] = 'pendente'
                else:
                    self.linhas_processadas.extend(registros)

            if self.conflitos_pendentes:
                self.popular_lista_conflitos()
                self.btn_finalizar_triagem.setEnabled(False)
                QMessageBox.warning(self, "Conflitos de OCI", f"Atenção: {len(self.conflitos_pendentes)} paciente(s) possuem mais de um exame marcado como OCI Concluída.\n\nA regra determina que apenas 1 vire OCI. Redirecionando para a Etapa 3 para resolução manual.")
                self.paginas.setCurrentIndex(2)
            else:
                self.finalizar_triagem_geral()

        except Exception as e:
            self.console_triagem.append(f"ERRO: {str(e)}")
            QMessageBox.critical(self, "Erro", f"Erro no processamento: {str(e)}")

    def popular_lista_conflitos(self):
        self.lista_cpfs.clear()
        for cpf, registros in self.conflitos_pendentes.items():
            nome = registros[0]['nome']
            item = QListWidgetItem(f"{cpf} - {nome}")
            
            if self.status_conflitos[cpf] == 'pendente':
                item.setBackground(QColor('#FFEB3B')) 
                item.setForeground(QColor('#000000'))
            else:
                item.setBackground(QColor('#4CAF50')) 
                item.setForeground(QColor('#FFFFFF'))
                
            self.lista_cpfs.addItem(item)
            
    def carregar_conflito_selecionado(self, item):
        cpf = item.text().split(' - ')[0]
        registros = self.conflitos_pendentes[cpf]
        
        candidatos = [r for r in registros if r['status_malha'].startswith('CONCLUID') or r['status_malha'].startswith('OCI_CONCLUID') or r.get('decisao')]
        
        self.tabela_conflitos.setRowCount(len(candidatos))
        for i, cand in enumerate(candidatos):
            self.tabela_conflitos.setItem(i, 0, QTableWidgetItem(cand['proc']))
            self.tabela_conflitos.setItem(i, 1, QTableWidgetItem(cand['data_fmt']))
            self.tabela_conflitos.setItem(i, 2, QTableWidgetItem(self.formatar_moeda(cand['valor'])))
            
            combo = QComboBox()
            combo.addItems(["", "OCI_CONCLUIDA", "BPA-I"])
            if cand.get('decisao'):
                combo.setCurrentText(cand['decisao'])
            self.tabela_conflitos.setCellWidget(i, 3, combo)
            
    def salvar_decisao_paciente(self):
        item = self.lista_cpfs.currentItem()
        if not item: return
        cpf = item.text().split(' - ')[0]
        
        registros = self.conflitos_pendentes[cpf]
        candidatos = [r for r in registros if r['status_malha'].startswith('CONCLUID') or r['status_malha'].startswith('OCI_CONCLUID') or r.get('decisao')]
        
        decisoes = []
        for i in range(len(candidatos)):
            combo = self.tabela_conflitos.cellWidget(i, 3)
            decisao = combo.currentText()
            decisoes.append(decisao)
            
        if "" in decisoes:
            QMessageBox.warning(self, "Aviso", "Preencha a decisão para todos os procedimentos da lista antes de salvar.")
            return
            
        if decisoes.count("OCI_CONCLUIDA") != 1:
            QMessageBox.warning(self, "Aviso", "Deve haver exatamente UMA decisão 'OCI_CONCLUIDA' salva por paciente.\nOs demais deverão ser 'BPA-I'.")
            return
            
        for i, cand in enumerate(candidatos):
            cand['decisao'] = decisoes[i]
            cand['status_malha'] = decisoes[i] 
            
        self.status_conflitos[cpf] = 'resolvido'
        
        item.setBackground(QColor('#4CAF50'))
        item.setForeground(QColor('#FFFFFF'))
        
        if all(s == 'resolvido' for s in self.status_conflitos.values()):
            self.btn_finalizar_triagem.setEnabled(True)
            QMessageBox.information(self, "Quase Lá", "Todas as decisões foram tomadas! Clique no botão de FINALIZAR TRIAGEM logo abaixo da lista.")
            
    def concluir_resolucao_conflitos(self):
        for cpf, registros in self.conflitos_pendentes.items():
            self.linhas_processadas.extend(registros)
        self.conflitos_pendentes.clear()
        self.status_conflitos.clear()
        self.finalizar_triagem_geral()

    def finalizar_triagem_geral(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        self.dados_bpa_limpo = []
        self.dados_excluidos = []
        self.dados_conferencia_rx = []
        self.dados_conferencia_denso = []
        self.registros_quarentena = []
        
        resumo = self.dados_app['historicos'][comp_atual]

        for reg in self.linhas_processadas:
            tipo = 'DENSO' if reg['proc'] == CODIGO_DENSITOMETRIA else 'RX'
            valor = reg['valor']
            
            # Normalização de Nomenclatura para evitar o problema de "CONCLUIDO" no relatorio
            status_original = reg['status_malha']
            if status_original.startswith('POSSIVEL'):
                status_padrao = 'POSSIVEL_OCI'
            elif status_original.startswith('CONCLUID') or status_original.startswith('OCI_CONCLUID'):
                status_padrao = 'OCI_CONCLUIDA'
            else:
                status_padrao = status_original
            
            resumo['Apresentado']['Total']['Qtd'] += 1
            resumo['Apresentado']['Total']['Valor'] += valor
            resumo['Apresentado'][tipo]['Qtd'] += 1
            resumo['Apresentado'][tipo]['Valor'] += valor
            
            cat_destino = None
            
            if status_padrao == 'POSSIVEL_OCI':
                cat_destino = 'Retido'
                self.dados_excluidos.append({
                    'Competencia_Realizacao': reg['comp_realizacao'],
                    'CPF': reg['cpf'], 'Nome': reg['nome'], 'Procedimento': reg['proc'],
                    'Valor': valor, 'Status_Origem': status_padrao
                })
                self.registros_quarentena.append({
                    'CPF': reg['cpf'], 'Procedimento': reg['proc'],
                    'Competencia_Origem': reg['comp_realizacao'],
                    'Competencia_Limite': self.calcular_competencia_limite(reg['comp_realizacao']),
                    'Status': 'AGUARDANDO_OCI', 'Linha_Raw': reg['linha_raw'].strip('\n')
                })
            elif status_padrao == 'OCI_CONCLUIDA':
                cat_destino = 'Baixado'
                self.dados_excluidos.append({
                    'Competencia_Realizacao': reg['comp_realizacao'],
                    'CPF': reg['cpf'], 'Nome': reg['nome'], 'Procedimento': reg['proc'],
                    'Valor': valor, 'Status_Origem': status_padrao
                })
            else:
                cat_destino = 'Liberado'
                self.dados_bpa_limpo.append(reg['linha_raw'])
                
            resumo[cat_destino]['Total']['Qtd'] += 1
            resumo[cat_destino]['Total']['Valor'] += valor
            resumo[cat_destino][tipo]['Qtd'] += 1
            resumo[cat_destino][tipo]['Valor'] += valor
                
            if tipo == 'DENSO':
                self.dados_conferencia_denso.append({
                    'CPF': reg['cpf'],
                    'Nome': reg['nome'],
                    'Procedimento': reg['proc']
                })
            else:
                self.dados_conferencia_rx.append({
                    'Data_Sort': reg['data_sort'],
                    'Data_Atendimento': reg['data_fmt'],
                    'CPF': reg['cpf'],
                    'Nome': reg['nome'],
                    'Procedimento': reg['proc'],
                    'Status': status_padrao
                })

        if self.registros_quarentena:
            df_novos = pd.DataFrame(self.registros_quarentena)
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

        self.paginas.setCurrentIndex(1)
        self.console_triagem.append("Triagem FINALIZADA com sucesso! Arquivo histórico (CSV) foi atualizado.\n-> Utilize os botões abaixo para gerar os relatórios na pasta EXPORT.")
        QMessageBox.information(self, "Triagem Finalizada", "Processamento completo! Você já pode gerar os arquivos de exportação.")

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
        QMessageBox.information(self, "Sucesso", f"Relatório de Excluídos salvo:\n{os.path.basename(arq_saida)}")

    def gerar_relatorio_conferencia(self):
        comp_atual = self.dados_app.get('competencia_atual', '')
        gerados = []
        
        df_rx = pd.DataFrame(self.dados_conferencia_rx)
        if not df_rx.empty:
            df_rx = df_rx.sort_values(by=['Data_Sort', 'Nome'])
            df_rx = df_rx[['Data_Atendimento', 'CPF', 'Nome', 'Procedimento', 'Status']]
            arq_rx = os.path.join(DIR_EXPORT, f"CONFERENCIA_RX_{comp_atual}.xlsx")
            df_rx.to_excel(arq_rx, index=False)
            gerados.append(os.path.basename(arq_rx))
            
        df_denso = pd.DataFrame(self.dados_conferencia_denso)
        if not df_denso.empty:
            df_denso = df_denso.sort_values(by=['Nome'])
            df_denso = df_denso[['CPF', 'Nome', 'Procedimento']]
            arq_denso = os.path.join(DIR_EXPORT, f"CONFERENCIA_DENSITOMETRIA_{comp_atual}.xlsx")
            df_denso.to_excel(arq_denso, index=False)
            gerados.append(os.path.basename(arq_denso))
            
        if gerados:
            msg = "\n".join(gerados)
            self.console_triagem.append(f"[GERADOS] -> \n{msg}")
            QMessageBox.information(self, "Sucesso", f"Planilhas de Conferência (XLSX) salvos na pasta EXPORT:\n\n{msg}")
        else:
            QMessageBox.information(self, "Aviso", "Nenhum dado de conferência para gerar.")

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
            malha = self.extrair_malha_planilha(df_sec)
            
            chaves_concluidas = {k for k, v in malha.items() if v.startswith('CONCLUID') or v.startswith('OCI_CONCLUID')}
            
            if not chaves_concluidas:
                QMessageBox.information(self, "Aviso", "A planilha não possui registros de OCI Concluída válidos.")
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
            QMessageBox.warning(self, "Aviso", "O 'BPA Bruto' e a Competência Atual precisam estar preenchidos para base do cabeçalho.")
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
                        cabecalho.append(linha[:7] + comp_atual + linha[13:])
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
            QMessageBox.information(self, "Sucesso", f"Repescagem concluída!\n\nFísico: {qtd_repescado} procedimentos\nFinanceiro: R$ {val_format}")

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
            t = dados[k]['Total']
            r = dados[k]['RX']
            d = dados[k]['DENSO']
            
            label = k.ljust(15)
            qtd_tot = str(t['Qtd']).zfill(4)
            val_tot = self.formatar_moeda(t['Valor']).rjust(12)
            conteudo += f"| {label} | TOTAL Físico: {qtd_tot} un. | Financeiro: R$ {val_tot} |\n"
            
            qtd_rx = str(r['Qtd']).zfill(4)
            val_rx = self.formatar_moeda(r['Valor']).rjust(12)
            
            show_denso = (k in ['Apresentado', 'Liberado'])
            if show_denso:
                qtd_denso = str(d['Qtd']).zfill(4)
                val_denso = self.formatar_moeda(d['Valor']).rjust(12)
                conteudo += f"|                 |  ├─ RX: {qtd_rx} un.       | Financeiro: R$ {val_rx.rjust(12)} |\n"
                conteudo += f"|                 |  └─ DENSO: {qtd_denso} un.    | Financeiro: R$ {val_denso.rjust(12)} |\n"
            else:
                conteudo += f"|                 |  └─ RX: {qtd_rx} un.       | Financeiro: R$ {val_rx.rjust(12)} |\n"
                
            conteudo += "-"*65 + "\n"
                
        with open(nome_extrato, "w", encoding="utf-8") as f:
            f.write(conteudo)
            
        QMessageBox.information(self, "Sucesso", f"Extrato salvo:\n{os.path.basename(nome_extrato)}")

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
        self.btn_rodar_triagem = QPushButton("1. EXECUTAR TRIAGEM")
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

    def criar_pagina_resolucao(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        
        layout.addWidget(QLabel("<h2>Etapa 3: Resolução de Conflitos OCI</h2>"))
        layout.addWidget(QLabel("Pacientes com múltiplos exames identificados como 'OCI Concluída' na malha.\nPara CADA paciente, escolha EXATAMENTE UMA linha para ser a OCI, e as demais devem ser BPA-I."))
        layout.addSpacing(10)
        
        h_split = QHBoxLayout()
        
        v_lista = QVBoxLayout()
        v_lista.addWidget(QLabel("Pacientes Pendentes (Amarelo = Pendente | Verde = Salvo):"))
        self.lista_cpfs = QListWidget()
        self.lista_cpfs.setFixedWidth(300)
        self.lista_cpfs.itemClicked.connect(self.carregar_conflito_selecionado)
        v_lista.addWidget(self.lista_cpfs)
        
        self.btn_finalizar_triagem = QPushButton("FINALIZAR TRIAGEM E GERAR RESULTADOS")
        self.btn_finalizar_triagem.setObjectName("btn_action")
        self.btn_finalizar_triagem.setEnabled(False)
        self.btn_finalizar_triagem.clicked.connect(self.concluir_resolucao_conflitos)
        v_lista.addWidget(self.btn_finalizar_triagem)
        
        h_split.addLayout(v_lista)
        
        v_tabela = QVBoxLayout()
        v_tabela.addWidget(QLabel("Defina o destino (Decisão) para CADA um dos exames abaixo e Salve:"))
        self.tabela_conflitos = QTableWidget(0, 4)
        self.tabela_conflitos.setHorizontalHeaderLabels(["Procedimento", "Data", "Valor", "Decisão Final"])
        self.tabela_conflitos.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        v_tabela.addWidget(self.tabela_conflitos)
        
        btn_resolver = QPushButton("SALVAR DECISÃO DESTE PACIENTE")
        btn_resolver.setObjectName("btn_warning")
        btn_resolver.clicked.connect(self.salvar_decisao_paciente)
        v_tabela.addWidget(btn_resolver)
        
        h_split.addLayout(v_tabela)
        layout.addLayout(h_split)
        return page

    def criar_pagina_gestao(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        h_title = QHBoxLayout()
        h_title.addWidget(QLabel("<h2>Etapa 4: Gestão de Banco de Dados e Atrasados</h2>"))
        btn_atualizar = QPushButton("Atualizar Tabela CSV")
        btn_atualizar.setFixedWidth(200)
        btn_atualizar.clicked.connect(self.carregar_historico_tabela)
        h_title.addWidget(btn_atualizar)
        layout.addLayout(h_title)
        
        layout.addWidget(QLabel("A) Atualização Manual OCI (Lote avulso)"))
        self.input_planilha_sec_extra = self.criar_linha_arquivo(layout, "Planilha Avulsa (Excel):", DIR_IMPORT)
        btn_atualizar_oci = QPushButton("Atualizar Status (Malha Avulsa)")
        btn_atualizar_oci.setObjectName("btn_action")
        btn_atualizar_oci.setFixedWidth(280)
        btn_atualizar_oci.clicked.connect(self.atualizar_historico_via_planilha)
        layout.addWidget(btn_atualizar_oci)
        
        layout.addSpacing(20)
        layout.addWidget(QLabel("B) Histórico (AGUARDANDO_OCI)"))
        self.input_filtro = QLineEdit()
        self.input_filtro.setPlaceholderText("🔎 Filtrar tabela...")
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
        h_filtro.addWidget(QLabel("<h2>Etapa 5: Extrato Físico-Financeiro</h2>"))
        h_filtro.addStretch()
        h_filtro.addWidget(QLabel("Competência:"))
        
        self.combo_comp = QComboBox()
        self.combo_comp.setFixedWidth(120)
        self.combo_comp.currentTextChanged.connect(self.atualizar_painel_resumo)
        h_filtro.addWidget(self.combo_comp)
        layout.addLayout(h_filtro)
        
        self.lbl_apresentado = QLabel()
        self.lbl_retido = QLabel()
        self.lbl_baixado = QLabel()
        self.lbl_liberado = QLabel()
        
        for lbl in [self.lbl_apresentado, self.lbl_retido, self.lbl_baixado, self.lbl_liberado]:
            lbl.setStyleSheet("font-size: 15px; padding: 15px; background-color: #2C2C2C; border-radius: 4px; margin-bottom: 5px; font-family: monospace;")
            layout.addWidget(lbl)
            
        layout.addSpacing(20)
        
        h_botoes = QHBoxLayout()
        h_botoes.setSpacing(20)
        
        btn_baixar = QPushButton("DOWNLOAD DO EXTRATO (TXT)")
        btn_baixar.setStyleSheet("QPushButton { background-color: #2196F3; font-size: 14px; padding: 15px; } QPushButton:hover { background-color: #1976D2; }")
        btn_baixar.clicked.connect(self.baixar_extrato_txt)
        h_botoes.addWidget(btn_baixar, 1)
        
        btn_proxima = QPushButton("AVANÇAR COMPETÊNCIA")
        btn_proxima.setStyleSheet("QPushButton { background-color: #FF9800; font-size: 14px; padding: 15px; } QPushButton:hover { background-color: #F57C00; }")
        btn_proxima.clicked.connect(self.avancar_competencia)
        h_botoes.addWidget(btn_proxima, 1)
        
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
        myappid = 'franca.saude.codure.v13' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = QApplication(sys.argv)
    janela = CodureApp()
    janela.show()
    sys.exit(app.exec())
