import sys
import os
import json
from datetime import date
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QHBoxLayout, QFormLayout, QLabel, QLineEdit,
                           QPushButton, QMessageBox, QComboBox, QInputDialog,
                           QMenu)
from PyQt6.QtGui import QIcon

import xlwings as xw
from preencher import MAPPING_00001, MAPPING_00002, MAPPING_00003, OUTPUT_DIR, get_next_proposal_number

def resource_path(relative_path):
    # se estiver rodando empacotado, usa o _MEIPASS
    base = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    return os.path.join(base, relative_path)

def normalize_price(valor_str: str) -> str:
    # troca vírgula por ponto
    v = valor_str.replace(',', '.')
    partes = v.split('.')
    if len(partes) <= 1:
        return partes[0]
    # junta tudo exceto o último segmento e coloca ponto antes do último
    return ''.join(partes[:-1]) + '.' + partes[-1]

class PropostaWindow(QMainWindow):
    def carregar_estruturas(self):
        try:
            with open(resource_path('json/estruturas.json'), 'r', encoding='utf-8') as f:
                data = json.load(f)
                estruturas = data.get('estruturas', ["TELHADO METÁLICO", "TELHADO CERÂMICO", "SOLO"])
                estruturas.sort()
                return estruturas
        except FileNotFoundError:
            return ["TELHADO METÁLICO", "TELHADO CERÂMICO", "SOLO"]

    def salvar_estruturas(self):
        with open(resource_path('json/estruturas.json'), 'w', encoding='utf-8') as f:
            json.dump({'estruturas': self.estruturas}, f, ensure_ascii=False, indent=4)
            
    def atualizar_combos_estruturas(self):
        """Atualiza todos os comboboxes de estrutura com a lista atualizada"""
        # Atualiza combo da proposta principal
        combo = self.campos["Estrutura Para"]
        current_text = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("Adicionar Estrutura")
        combo.addItems(self.estruturas)
        if current_text in self.estruturas:
            combo.setCurrentText(current_text)
        combo.blockSignals(False)
        
        # Atualiza combo da proposta 2 se existir
        if hasattr(self, 'campos_proposta2') and "Estrutura Para" in self.campos_proposta2:
            combo2 = self.campos_proposta2["Estrutura Para"]
            current_text2 = combo2.currentText()
            combo2.blockSignals(True)
            combo2.clear()
            combo2.addItem("Adicionar Estrutura")
            combo2.addItems(self.estruturas)
            if current_text2 in self.estruturas:
                combo2.setCurrentText(current_text2)
            combo2.blockSignals(False)

    def mostrar_menu_estrutura(self, pos):
        combo = self.campos["Estrutura Para"]
        estrutura_atual = combo.currentText()

        if estrutura_atual != "Adicionar Estrutura":
            menu = QMenu()
            editar_acao = menu.addAction("Editar Estrutura")
            excluir_acao = menu.addAction("Excluir Estrutura")
            acao = menu.exec(combo.mapToGlobal(pos))

            if acao == editar_acao:
                novo_nome, ok = QInputDialog.getText(
                    self, "Editar Estrutura",
                    "Nova estrutura:",
                    text=estrutura_atual
                )

                if ok and novo_nome.strip():
                    novo_nome = novo_nome.strip().upper()
                    if novo_nome != estrutura_atual:
                        self.estruturas.remove(estrutura_atual)
                        self.estruturas.append(novo_nome)
                        self.salvar_estruturas()
                        self.atualizar_combos_estruturas()

                        combo.blockSignals(True)
                        combo.clear()
                        combo.addItem("Adicionar Estrutura")
                        combo.addItems(self.estruturas)
                        combo.setCurrentText(novo_nome)
                        combo.blockSignals(False)
            elif acao == excluir_acao:
                resposta = QMessageBox.question(
                    self,
                    "Confirmar Exclusão",
                    f"Tem certeza que deseja excluir a estrutura {estrutura_atual}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if resposta == QMessageBox.StandardButton.Yes:
                    self.estruturas.remove(estrutura_atual)
                    self.salvar_estruturas()
                    self.atualizar_combos_estruturas()

                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Estrutura")
                    combo.addItems(self.estruturas)
                    combo.setCurrentText(self.estruturas[0] if self.estruturas else "Adicionar Estrutura")
                    combo.blockSignals(False)

    def estrutura_changed(self, text):
        if text == "Adicionar Estrutura":
            nova_estrutura, ok = QInputDialog.getText(self, "Nova Estrutura", "Tipo de estrutura:")
            if ok and nova_estrutura.strip():
                nova_estrutura = nova_estrutura.strip().upper()
                if nova_estrutura not in self.estruturas:
                    self.estruturas.append(nova_estrutura)
                    self.salvar_estruturas()
                    self.atualizar_combos_estruturas()
                    combo = self.campos["Estrutura Para"]
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Estrutura")
                    combo.addItems(self.estruturas)
                    combo.setCurrentText(nova_estrutura)
                    combo.blockSignals(False)

    def carregar_tema(self):
        try:
            with open(resource_path('json/tema.json'), 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('tema', 'sistema')
        except FileNotFoundError:
            return 'sistema'
    
    def salvar_tema(self, tema):
        with open(resource_path('json/tema.json'), 'w', encoding='utf-8') as f:
            json.dump({'tema': tema}, f, ensure_ascii=False, indent=4)
    
    def aplicar_tema(self, tema):
        if tema == 'claro':
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #ffffff; color: #000000; }
                QLineEdit, QComboBox { background-color: #f0f0f0; border: 1px solid #cccccc; padding: 5px; }
                QPushButton { background-color: #e0e0e0; border: 1px solid #cccccc; padding: 5px 10px; }
                QPushButton:hover { background-color: #d0d0d0; }
            """)
            self.btn_tema.setText('○')
        elif tema == 'escuro':
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #2b2b2b; color: #ffffff; }
                QLineEdit, QComboBox { background-color: #3b3b3b; border: 1px solid #505050; padding: 5px; color: #ffffff; }
                QPushButton { background-color: #404040; border: 1px solid #505050; padding: 5px 10px; color: #ffffff; }
                QPushButton:hover { background-color: #505050; }
            """)
            self.btn_tema.setText('○')
        else:  # sistema
            self.setStyleSheet('')
            self.btn_tema.setText('◐')
    
    def alternar_tema(self):
        temas = ['sistema', 'claro', 'escuro']
        tema_atual = self.carregar_tema()
        novo_tema = temas[(temas.index(tema_atual) + 1) % len(temas)]
        self.salvar_tema(novo_tema)
        self.aplicar_tema(novo_tema)
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Propostas")
        self.setMinimumWidth(600)
        
        # Definir o ícone da aplicação
        icon = QIcon("ícone.ico")
        self.setWindowIcon(icon)
        
        # Widget e layout principal
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Adicionar botão de tema no canto superior direito
        header_layout = QHBoxLayout()
        header_layout.addStretch()
        self.btn_tema = QPushButton('◐')
        self.btn_tema.setFixedSize(25, 25)
        self.btn_tema.clicked.connect(self.alternar_tema)
        header_layout.addWidget(self.btn_tema)
        layout.addLayout(header_layout)
        
        # Aplicar tema inicial
        tema_inicial = self.carregar_tema()
        self.aplicar_tema(tema_inicial)
        
        # Form layout para os campos
        form_layout = QFormLayout()
        self.form_layout = form_layout
        self.campos = {}
        self.campos_inversores = []
        self.campos_proposta3 = {}
        self.label_preco = None # Adicionado para guardar a label do Preço
        self.labels_proposta3 = {} # Adicionado para guardar as labels da proposta 3
        
        # após criar form_layout…
        self.subtitle_proposta1 = QLabel("<b>DADOS DA PROPOSTA</b>")

        # Conectar o evento de mudança do tipo de proposta
        self.tipo_proposta_anterior = "1- Proposta Simples"
        
        # Carregar consultores, logradouros e estados
        self.consultores = self.carregar_consultores()
        self.logradouros = self.carregar_logradouros()
        self.estados = self.carregar_estados()
        self.estruturas = self.carregar_estruturas()
        
        # Valores padrão
        today = date.today().strftime("%d/%m/%Y")
        next_number = get_next_proposal_number()
        
        # Criar campos do formulário
        # Adicionar Tipo de Proposta antes dos dados do cliente
        combo_tipo_proposta = QComboBox()
        combo_tipo_proposta.setEditable(False)
        combo_tipo_proposta.addItems([
            "1- Proposta Simples",
            "2- Proposta Dupla",
            "3- Proposta com Mão de Obra"
        ])
        combo_tipo_proposta.currentTextChanged.connect(self.tipo_proposta_changed)
        self.campos["Tipo de Proposta"] = combo_tipo_proposta
        form_layout.addRow("Tipo de Proposta", combo_tipo_proposta)

        # Selecionar o mapeamento correto com base no tipo de proposta
        if combo_tipo_proposta.currentText() == "1- Proposta Simples":
            self.mapping = MAPPING_00001
        elif combo_tipo_proposta.currentText() == "2- Proposta Dupla":
            self.mapping = MAPPING_00002
        elif combo_tipo_proposta.currentText() == "3- Proposta com Mão de Obra":
            self.mapping = MAPPING_00003

        for label in self.mapping.keys():
            if label == "Nome do Cliente":
                form_layout.addRow(QLabel())
                form_layout.addRow(QLabel("<b>DADOS DO CLIENTE</b>"), QLabel())
            elif label == "Quantidade de Painéis":
                form_layout.addRow(QLabel())
                form_layout.addRow(self.subtitle_proposta1, QLabel())

            if label == "Data":
                self.campos[label] = QLineEdit(today)
                self.campos[label].textChanged.connect(self.format_date)
                self.campos[label].setMaxLength(10)
            elif label == "N° da Proposta":
                self.campos[label] = QLineEdit(next_number)
            elif label == "Logradouro":
                combo = QComboBox()
                combo.setEditable(False)
                combo.addItem("Adicionar Logradouro")
                combo.addItems(self.carregar_logradouros())
                combo.setCurrentText("RUA")
                combo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                combo.customContextMenuRequested.connect(self.mostrar_menu_logradouro)
                combo.currentTextChanged.connect(self.logradouro_changed)
                self.campos[label] = combo
            elif label == "Estado":
                combo = QComboBox()
                combo.setEditable(False)
                combo.addItem("Adicionar Estado")
                combo.addItems(self.carregar_estados())
                combo.setCurrentText("MG")
                combo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                combo.customContextMenuRequested.connect(self.mostrar_menu_estado)
                combo.currentTextChanged.connect(self.estado_changed)
                self.campos[label] = combo
            elif label == "Cidade":
                self.campos[label] = QLineEdit("NOVA SERRANA")
            elif label == "Quantidade de Inversores":
                self.campos[label] = QLineEdit()
                self.campos[label].textChanged.connect(self.atualizar_campos_inversores)
            elif label == "Potência Inversor 1 (W)":
                continue
            elif label == "Estrutura Para":
                combo = QComboBox()
                combo.setEditable(False)
                combo.addItem("Adicionar Estrutura")
                combo.addItems(self.carregar_estruturas())
                combo.setCurrentText("TELHADO METÁLICO")
                combo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                combo.customContextMenuRequested.connect(self.mostrar_menu_estrutura)
                combo.currentTextChanged.connect(self.estrutura_changed)
                self.campos[label] = combo
            elif label == "Telefone":
                self.campos[label] = QLineEdit()
                self.campos[label].textChanged.connect(self.format_phone)
                self.campos[label].setMaxLength(14)
            elif label == "Consultor":
                combo = QComboBox()
                combo.setEditable(False)
                combo.addItem("Adicionar Consultor")
                combo.addItems(self.consultores)
                # Carregar último consultor usado
                try:
                    with open(resource_path('json/ultimo_consultor.json'), 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        ultimo_consultor = data.get('ultimo_consultor')
                        if ultimo_consultor and ultimo_consultor in self.consultores:
                            combo.setCurrentText(ultimo_consultor)
                        else:
                            combo.setCurrentText(self.consultores[0] if self.consultores else "Adicionar Consultor")
                except:
                    combo.setCurrentText(self.consultores[0] if self.consultores else "Adicionar Consultor")
                combo.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                combo.customContextMenuRequested.connect(self.mostrar_menu_consultor)
                combo.currentTextChanged.connect(self.consultor_changed)
                self.campos[label] = combo
            # Não adicionar campos de preço aqui diretamente no loop
            elif label in ["Preço", "Preço dos Equipamentos", "Preço da Mão de Obra", "Preço Total"]:
                continue # Pula a criação automática desses campos
            else:
                self.campos[label] = QLineEdit()
            
            form_layout.addRow(label, self.campos[label])
            
            # Adiciona campo de número após Endereço
            if label == "Endereço":
                self.numero_end = QLineEdit()
                form_layout.addRow("Número", self.numero_end)

        # --- Adicionar campos de preço manualmente após o loop ---
        # Campo Preço (para propostas 1 e 2)
        self.campos["Preço"] = QLineEdit()
        self.label_preco = QLabel("Preço")
        form_layout.addRow(self.label_preco, self.campos["Preço"])

        # Campos para Proposta 3
        self.campos_proposta3["Preço dos Equipamentos"] = QLineEdit()
        self.campos_proposta3["Preço da Mão de Obra"] = QLineEdit()
        self.campos_proposta3["Preço Total"] = QLineEdit()
        self.campos_proposta3["Preço Total"].setReadOnly(False)

        # Conectar eventos para atualização automática do total
        self.campos_proposta3["Preço dos Equipamentos"].textChanged.connect(self.atualizar_preco_total)
        self.campos_proposta3["Preço da Mão de Obra"].textChanged.connect(self.atualizar_preco_total)

        # Adicionar campos da proposta 3 ao layout e guardar labels
        label_equip = QLabel("Preço dos Equipamentos")
        form_layout.addRow(label_equip, self.campos_proposta3["Preço dos Equipamentos"])
        self.labels_proposta3["Preço dos Equipamentos"] = label_equip

        label_obra = QLabel("Preço da Mão de Obra")
        form_layout.addRow(label_obra, self.campos_proposta3["Preço da Mão de Obra"])
        self.labels_proposta3["Preço da Mão de Obra"] = label_obra

        label_total = QLabel("Preço Total")
        form_layout.addRow(label_total, self.campos_proposta3["Preço Total"])
        self.labels_proposta3["Preço Total"] = label_total

        # Ocultar campos da proposta 3 inicialmente (já que o padrão é tipo 1)
        for label_widget in self.labels_proposta3.values():
            label_widget.setVisible(False)
        for campo_widget in self.campos_proposta3.values():
            campo_widget.setVisible(False)
        # --- Fim da adição manual dos campos de preço ---

        # Subtítulo Proposta 2 (inicialmente oculto)
        self.espaco_proposta2 = QLabel()
        form_layout.addRow(self.espaco_proposta2, QLabel())
        self.espaco_proposta2.setVisible(False)
        self.subtitle_proposta2 = QLabel("<b>DADOS DA PROPOSTA 2</b>")
        form_layout.addRow(self.subtitle_proposta2, QLabel())
        self.subtitle_proposta2.setVisible(False)

        self.campos_proposta2 = {
            "Quantidade de Painéis": QLineEdit(),
            "Potência dos Painéis (W)": QLineEdit(),
            "Quantidade de Inversores": QLineEdit(),
            "Estrutura Para": QComboBox(),
            "Produção Média Mensal": QLineEdit(),
            "Preço": QLineEdit(),
        }
        
        # Configurar campos especiais da proposta 2
        self.campos_proposta2["Quantidade de Inversores"].textChanged.connect(self.atualizar_campos_inversores_proposta2)
        self.campos_inversores_proposta2 = []
        
        # Configurar combo de estrutura
        combo_estrutura = self.campos_proposta2["Estrutura Para"]
        combo_estrutura.setEditable(False)
        combo_estrutura.addItem("Adicionar Estrutura")
        combo_estrutura.addItems(self.estruturas)
        combo_estrutura.setCurrentText("TELHADO METÁLICO")
        combo_estrutura.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        combo_estrutura.customContextMenuRequested.connect(self.mostrar_menu_estrutura_proposta2)
        combo_estrutura.currentTextChanged.connect(self.estrutura_changed_proposta2)
        self.labels_proposta2 = {}
        for label, widget in self.campos_proposta2.items():
            lbl = QLabel(label)
            form_layout.addRow(lbl, widget)
            self.labels_proposta2[label] = lbl
            widget.setVisible(False)
            lbl.setVisible(False)

        # Ocultar campos da proposta 3 inicialmente (já que o padrão é tipo 1)
        for label_widget in self.labels_proposta3.values():
            label_widget.setVisible(False)
        for campo_widget in self.campos_proposta3.values():
            campo_widget.setVisible(False)
        # --- Fim da adição manual dos campos de preço ---

        from PyQt6.QtWidgets import QScrollArea

        # … dentro de __init__:
        container = QWidget()
        container.setLayout(form_layout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)

        layout.addWidget(scroll)

        
        # Botão de gerar proposta
        btn_gerar = QPushButton("Gerar Proposta")
        btn_gerar.clicked.connect(self.gerar_proposta)
        layout.addWidget(btn_gerar)
    
    def format_date(self, text):
        campo_data = self.campos["Data"]
        texto = "".join(filter(str.isdigit, text))
        
        if not texto:
            campo_data.setText("")
            return
        
        if len(texto) > 8:
            texto = texto[:8]
        
        formatted = ""
        cursor_pos = 0
        
        if len(texto) > 0:
            formatted += texto[:2]
            cursor_pos = min(2, len(texto))
        if len(texto) > 2:
            formatted += "/" + texto[2:4]
            cursor_pos = min(5, len(texto) + 1)
        if len(texto) > 4:
            formatted += "/" + texto[4:]
            cursor_pos = min(10, len(texto) + 2)
            
        if formatted != text:
            campo_data.setText(formatted)
            campo_data.setCursorPosition(cursor_pos)
    
    def format_phone(self, text):
        campo_phone = self.campos["Telefone"]
        texto = "".join(filter(str.isdigit, text))
        
        if not texto:
            campo_phone.setText("")
            return
        
        # Limita o número de dígitos a 11
        if len(texto) > 11:
            texto = texto[:11]
        
        formatted = ""
        
        # Formata conforme o número de dígitos
        if len(texto) <= 8:  # XXXX-XXXX
            if len(texto) > 4:
                formatted = f"{texto[:4]}-{texto[4:]}"
            else:
                formatted = texto
        elif len(texto) == 9:  # XXXXX-XXXX
            formatted = f"{texto[:5]}-{texto[5:]}"
        else:  # XX XXXXX-XXXX
            if len(texto) > 7:
                formatted = f"{texto[:2]} {texto[2:7]}-{texto[7:]}"
            else:
                formatted = f"{texto[:2]} {texto[2:]}"
        
        if formatted != text:
            campo_phone.setText(formatted)
            campo_phone.setCursorPosition(len(formatted))
    
    def atualizar_campos_inversores(self, text):
        # Remover campos antigos
        for campo in self.campos_inversores:
            campo.deleteLater()
        self.campos_inversores.clear()
        
        # Adicionar novos campos
        try:
            quantidade = int(text) if text else 0
            form_layout = self.form_layout
            
            # Encontrar o índice após "Quantidade de Inversores"
            idx = 0
            for i in range(form_layout.rowCount()):
                if form_layout.itemAt(i, QFormLayout.ItemRole.LabelRole) and \
                   form_layout.itemAt(i, QFormLayout.ItemRole.LabelRole).widget().text() == "Quantidade de Inversores":
                    idx = i + 1
                    break
            
            # Adicionar campos de potência
            for i in range(quantidade):
                label = QLabel(f"Potência do Inversor {i+1} (W)")
                campo = QLineEdit()
                form_layout.insertRow(idx + i, label, campo)
                self.campos_inversores.append(label)
                self.campos_inversores.append(campo)
        except ValueError:
            pass

    def tipo_proposta_changed(self, novo_tipo):
        is_proposta3 = (novo_tipo == "3- Proposta com Mão de Obra")

        # Alternar visibilidade do campo Preço (Propostas 1 e 2)
        if "Preço" in self.campos and self.label_preco:
            self.campos["Preço"].setVisible(not is_proposta3)
            self.label_preco.setVisible(not is_proposta3)

        # Alternar visibilidade dos campos da Proposta 3
        for label_key, label_widget in self.labels_proposta3.items():
            label_widget.setVisible(is_proposta3)
        for campo_key, campo_widget in self.campos_proposta3.items():
            campo_widget.setVisible(is_proposta3)

        self.tipo_proposta_anterior = novo_tipo

        # renomeia o subtítulo de acordo com o tipo 2
        if novo_tipo == "2- Proposta Dupla":
            self.subtitle_proposta1.setText("<b>DADOS DA PROPOSTA 1</b>")
        else:
            # volta ao nome padrão nos demais casos
            self.subtitle_proposta1.setText("<b>DADOS DA PROPOSTA</b>")

        is_proposta2 = (novo_tipo == "2- Proposta Dupla")
        self.espaco_proposta2.setVisible(is_proposta2)
        self.subtitle_proposta2.setVisible(is_proposta2)
        for label, lbl in self.labels_proposta2.items():
            lbl.setVisible(is_proposta2)
            self.campos_proposta2[label].setVisible(is_proposta2)

        # Remover campos antigos de inversores da proposta 2
        for campo in self.campos_inversores_proposta2:
            campo.deleteLater()
        self.campos_inversores_proposta2.clear()

        # Adicionar novos campos de inversores da proposta 2 se for tipo 2
        if is_proposta2:
            try:
                quantidade = int(self.campos_proposta2["Quantidade de Inversores"].text()) if self.campos_proposta2["Quantidade de Inversores"].text() else 0
                form_layout = self.form_layout
                idx = form_layout.rowCount() - 1
                for i in range(quantidade):
                    label = QLabel(f"Potência do Inversor {i+1} (W)")
                    campo = QLineEdit()
                    form_layout.insertRow(idx + i, label, campo)
                    self.campos_inversores_proposta2.append(label)
                    self.campos_inversores_proposta2.append(campo)
            except ValueError:
                pass

        self.tipo_proposta_anterior = novo_tipo

    def atualizar_preco_total(self):
        try:
            equip = normalize_price(self.campos_proposta3["Preço dos Equipamentos"].text() or "0")
            obra = normalize_price(self.campos_proposta3["Preço da Mão de Obra"].text() or "0")
            total = float(equip) + float(obra)
            self.campos_proposta3["Preço Total"].setText(f"{total:.2f}")
        except ValueError:
            pass

    def atualizar_campos_inversores_proposta2(self, text):
        # Remover campos antigos
        for campo in self.campos_inversores_proposta2:
            campo.deleteLater()
        self.campos_inversores_proposta2.clear()
        
        # Adicionar novos campos
        try:
            quantidade = int(text) if text else 0
            form_layout = self.form_layout
            
            # Encontrar o índice após "Quantidade de Inversores" da proposta 2
            idx = 0
            for i in range(form_layout.rowCount()):
                if form_layout.itemAt(i, QFormLayout.ItemRole.LabelRole) and \
                   form_layout.itemAt(i, QFormLayout.ItemRole.LabelRole).widget().text() == "Quantidade de Inversores" and \
                   form_layout.itemAt(i, QFormLayout.ItemRole.FieldRole).widget() == self.campos_proposta2["Quantidade de Inversores"]:
                    idx = i + 1
                    break
            
            # Adicionar campos de potência
            for i in range(quantidade):
                label = QLabel(f"Potência do Inversor {i+1} (W)")
                campo = QLineEdit()
                form_layout.insertRow(idx + i, label, campo)
                self.campos_inversores_proposta2.append(label)
                self.campos_inversores_proposta2.append(campo)
        except ValueError:
            pass

    def mostrar_menu_estrutura_proposta2(self, pos):
        combo = self.campos_proposta2["Estrutura Para"]
        estrutura_atual = combo.currentText()

        if estrutura_atual != "Adicionar Estrutura":
            menu = QMenu()
            editar_acao = menu.addAction("Editar Estrutura")
            excluir_acao = menu.addAction("Excluir Estrutura")
            acao = menu.exec(combo.mapToGlobal(pos))

            if acao == editar_acao:
                novo_nome, ok = QInputDialog.getText(
                    self, "Editar Estrutura",
                    "Nova estrutura:",
                    text=estrutura_atual
                )

                if ok and novo_nome.strip():
                    novo_nome = novo_nome.strip().upper()
                    if novo_nome != estrutura_atual:
                        self.estruturas.remove(estrutura_atual)
                        self.estruturas.append(novo_nome)
                        self.salvar_estruturas()
                        self.atualizar_combos_estruturas()

                        combo.blockSignals(True)
                        combo.clear()
                        combo.addItem("Adicionar Estrutura")
                        combo.addItems(self.estruturas)
                        combo.setCurrentText(novo_nome)
                        combo.blockSignals(False)
            elif acao == excluir_acao:
                resposta = QMessageBox.question(
                    self,
                    "Confirmar Exclusão",
                    f"Tem certeza que deseja excluir a estrutura {estrutura_atual}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if resposta == QMessageBox.StandardButton.Yes:
                    self.estruturas.remove(estrutura_atual)
                    self.salvar_estruturas()
                    self.atualizar_combos_estruturas()

                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Estrutura")
                    combo.addItems(self.estruturas)
                    combo.setCurrentText(self.estruturas[0] if self.estruturas else "Adicionar Estrutura")
                    combo.blockSignals(False)

    def estrutura_changed_proposta2(self, text):
        if text == "Adicionar Estrutura":
            nova_estrutura, ok = QInputDialog.getText(self, "Nova Estrutura", "Tipo de estrutura:")
            if ok and nova_estrutura.strip():
                nova_estrutura = nova_estrutura.strip().upper()
                if nova_estrutura not in self.estruturas:
                    self.estruturas.append(nova_estrutura)
                    self.salvar_estruturas()
                    self.atualizar_combos_estruturas()
                    combo = self.campos_proposta2["Estrutura Para"]
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Estrutura")
                    combo.addItems(self.estruturas)
                    combo.setCurrentText(nova_estrutura)
                    combo.blockSignals(False)

    def gerar_proposta(self):
        try:
            dados = {}
            
            for label, campo in self.campos.items():
                if isinstance(campo, QComboBox):
                    valor = campo.currentText()
                else:
                    raw = campo.text().strip()
                    if label == "Preço":
                        valor = normalize_price(raw)
                    else:
                        valor = raw.upper()
                        if label == "Data" and not valor:
                            valor = date.today().strftime("%d/%m/%Y")
                dados[label] = valor
            
            # Adicionar campos específicos da proposta 3 se necessário
            if dados["Tipo de Proposta"] == "3- Proposta com Mão de Obra":
                for label, campo in self.campos_proposta3.items():
                    # Normaliza apenas os campos de entrada
                    dados[label] = normalize_price(campo.text().strip() or "0")
            else: # Se não for proposta 3, pega o valor do campo Preço
                 if "Preço" in self.campos:
                     dados["Preço"] = normalize_price(self.campos["Preço"].text().strip() or "0")
            
            # ── AQUI: popula dados da Proposta 2 ──
            if dados["Tipo de Proposta"] == "2- Proposta Dupla":
                # campos fixos da proposta 2
                for label, widget in self.campos_proposta2.items():
                    valor = (widget.currentText() 
                             if isinstance(widget, QComboBox) 
                             else normalize_price(widget.text().strip()))
                    dados[f"{label} 2"] = valor.upper()
                # inversores da proposta 2 (mesma lógica do grupo 1)
                vals2 = [
                    self.campos_inversores_proposta2[i+1].text().strip().replace(',', '.')
                    for i in range(0, len(self.campos_inversores_proposta2), 2)
                    if self.campos_inversores_proposta2[i+1].text().strip()
                ]
                # agrupa e formata “X de Y LV”
                from collections import Counter
                cnt2 = Counter(vals2)
                parts2 = []
                for val, qtd in sorted(cnt2.items(), key=lambda x: float(x[0])):
                    parts2.append(f"{qtd} de {val} LV" if qtd>1 else f"{val} LV")
                if parts2:
                    texto2 = ", ".join(parts2[:-1]) + " E " + parts2[-1] if len(parts2)>1 else parts2[0]
                else:
                    texto2 = ""
                dados["Potência Inversor 1 (W) 2"] = texto2

            # Concatenar potências dos inversores
            form_layout = self.centralWidget().layout().itemAt(0).layout()
            from collections import Counter

            # 1) Captura só os valores preenchidos
            vals = [
                self.campos_inversores[i+1].text().strip().replace(',', '.')
                for i in range(0, len(self.campos_inversores), 2)
                if self.campos_inversores[i+1].text().strip()
            ]

            # 2) Conta repetições e ordena numericamente
            cnt = Counter(vals)
            itens = sorted(cnt.items(), key=lambda x: float(x[0]))

            # 3) Monta cada parte: "X de Y LV" ou "Y LV"
            parts = []
            for val, qtd in itens:
                if qtd > 1:
                    parts.append(f"{qtd} de {val} LV")
                else:
                    parts.append(f"{val} LV")

            # 4) Junta com ", " e " E " antes do último
            if not parts:
                resultado = ""
            elif len(parts) == 1:
                resultado = parts[0]
            else:
                resultado = ", ".join(parts[:-1]) + " E " + parts[-1]

            dados["Potência Inversor 1 (W)"] = resultado
            
            # Seleciona o template baseado no tipo de proposta
            if dados["Tipo de Proposta"] ==  "1- Proposta Simples":
                tipo_proposta = "00001 - FAZER PROPOSTA PC"
                mapping = MAPPING_00001
            elif dados["Tipo de Proposta"] ==  "2- Proposta Dupla":
                tipo_proposta = "00002 - FAZER DUPLA PROPOSTA PC"
                mapping = MAPPING_00002
            elif dados["Tipo de Proposta"] ==  "3- Proposta com Mão de Obra":
                tipo_proposta = "00003 - FAZER PROPOSTA MAO DE OBRA E EQUIPAMENTOS PC"
                mapping = MAPPING_00003
            template = f"templates/{tipo_proposta}.xlsx"
            
            app = xw.App(visible=False)
            wb = app.books.open(resource_path(template))
            sht = wb.sheets[0]
            
            # Preenche as células (tratando as chaves “2”)
            for label, cell in mapping.items():
                sht.range(cell).value = dados.get(label, "")
            
            # Preenche número do endereço
            numero = self.numero_end.text().strip().upper()
            if numero:
                sht.range("I13").value = "Nº"
                sht.range("J13").value = numero
            
            # Gera nome do PDF
            numero_proposta = dados["N° da Proposta"]
            nome_cliente = dados["Nome do Cliente"]
            filename = f"{numero_proposta}PROPOSTA {nome_cliente}.pdf"
            output_path = os.path.join(OUTPUT_DIR, filename)
            
            # Exporta para PDF
            wb.to_pdf(output_path)
            
            # Finaliza
            wb.close()
            app.quit()
            
            # Criar caixa de mensagem personalizada
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Sucesso")
            msg_box.setText(f"PDF gerado com sucesso!\n{output_path}")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
            btn_visualizar = msg_box.addButton("Visualizar PDF", QMessageBox.ButtonRole.ActionRole)
            btn_imprimir  = msg_box.addButton("Imprimir PDF",  QMessageBox.ButtonRole.ActionRole)
            msg_box.exec()

            clicado = msg_box.clickedButton()
            if clicado == btn_visualizar:
                os.startfile(output_path)
            elif clicado == btn_imprimir:
                # no Windows, "print" manda para a impressora padrão
                os.startfile(output_path, "print")
            
            # Salvar último consultor usado
            with open(resource_path('json/ultimo_consultor.json'), 'w', encoding='utf-8') as f:
                json.dump({'ultimo_consultor': dados['Consultor']}, f, ensure_ascii=False, indent=4)
            
            # Atualiza número da proposta
            next_num = str(int(numero_proposta) + 1)
            self.campos["N° da Proposta"].setText(next_num)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar proposta:\n{str(e)}")

    def carregar_logradouros(self):
        try:
            with open(resource_path('json/logradouros.json'), 'r', encoding='utf-8') as f:
                data = json.load(f)
                logradouros = data.get('logradouros', [])
                logradouros.sort()
                return logradouros
        except FileNotFoundError:
            return ["RUA", "AVENIDA", "TRAVESSA", "ALAMEDA", "PRAÇA", "RODOVIA"]
    
    def salvar_logradouros(self):
        with open(resource_path('json/logradouros.json'), 'w', encoding='utf-8') as f:
            json.dump({'logradouros': self.logradouros}, f, ensure_ascii=False, indent=4)
    
    def logradouro_changed(self, text):
        if text == "Adicionar Logradouro":
            novo_logradouro, ok = QInputDialog.getText(self, "Novo Logradouro", "Tipo de logradouro:")
            if ok and novo_logradouro.strip():
                novo_logradouro = novo_logradouro.strip().upper()
                if novo_logradouro not in self.logradouros:
                    self.logradouros.append(novo_logradouro)
                    self.salvar_logradouros()
                    combo = self.campos["Logradouro"]
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Logradouro")
                    combo.addItems(self.logradouros)
                    combo.setCurrentText(novo_logradouro)
                    combo.blockSignals(False)

    def mostrar_menu_logradouro(self, pos):
        combo = self.campos["Logradouro"]
        logradouro_atual = combo.currentText()
        
        if logradouro_atual != "Adicionar Logradouro":
            menu = QMenu()
            editar_acao = menu.addAction("Editar Logradouro")
            excluir_acao = menu.addAction("Excluir Logradouro")
            acao = menu.exec(combo.mapToGlobal(pos))
            
            if acao == editar_acao:
                novo_nome, ok = QInputDialog.getText(
                    self, "Editar Logradouro",
                    "Novo tipo de logradouro:",
                    text=logradouro_atual
                )
                
                if ok and novo_nome.strip():
                    novo_nome = novo_nome.strip().upper()
                    if novo_nome != logradouro_atual:
                        self.logradouros.remove(logradouro_atual)
                        self.logradouros.append(novo_nome)
                        self.salvar_logradouros()
                        
                        combo.blockSignals(True)
                        combo.clear()
                        combo.addItem("Adicionar Logradouro")
                        combo.addItems(self.logradouros)
                        combo.setCurrentText(novo_nome)
                        combo.blockSignals(False)
            elif acao == excluir_acao:
                resposta = QMessageBox.question(
                    self,
                    "Confirmar Exclusão",
                    f"Tem certeza que deseja excluir o logradouro {logradouro_atual}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                
                if resposta == QMessageBox.StandardButton.Yes:
                    self.logradouros.remove(logradouro_atual)
                    self.salvar_logradouros()
                    
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Logradouro")
                    combo.addItems(self.logradouros)
                    combo.setCurrentText(self.logradouros[0] if self.logradouros else "Adicionar Logradouro")
                    combo.blockSignals(False)

    def carregar_estados(self):
        try:
            with open(resource_path('json/estados.json'), 'r', encoding='utf-8') as f:
                data = json.load(f)
                estados = data.get('estados', ["MG", "SP", "RJ", "ES"])
                estados.sort()
                return estados
        except FileNotFoundError:
            return ["MG", "SP", "RJ", "ES"]
    
    def salvar_estados(self):
        with open(resource_path('json/estados.json'), 'w', encoding='utf-8') as f:
            json.dump({'estados': self.estados}, f, ensure_ascii=False, indent=4)
            
    def estado_changed(self, text):
        if text == "Adicionar Estado":
            novo_estado, ok = QInputDialog.getText(self, "Novo Estado", "Sigla do estado:")
            if ok and novo_estado.strip():
                novo_estado = novo_estado.strip().upper()
                if novo_estado not in self.estados:
                    self.estados.append(novo_estado)
                    self.salvar_estados()
                    combo = self.campos["Estado"]
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Estado")
                    combo.addItems(self.estados)
                    combo.setCurrentText(novo_estado)
                    combo.blockSignals(False)

    def mostrar_menu_estado(self, pos):
        combo = self.campos["Estado"]
        estado_atual = combo.currentText()
        
        if estado_atual != "Adicionar Estado":
            menu = QMenu()
            editar_acao = menu.addAction("Editar Estado")
            excluir_acao = menu.addAction("Excluir Estado")
            acao = menu.exec(combo.mapToGlobal(pos))
            
            if acao == editar_acao:
                novo_nome, ok = QInputDialog.getText(
                    self, "Editar Estado",
                    "Nova sigla do estado:",
                    text=estado_atual
                )
                
                if ok and novo_nome.strip():
                    novo_nome = novo_nome.strip().upper()
                    if novo_nome != estado_atual:
                        self.estados.remove(estado_atual)
                        self.estados.append(novo_nome)
                        self.salvar_estados()
                        
                        combo.blockSignals(True)
                        combo.clear()
                        combo.addItem("Adicionar Estado")
                        combo.addItems(self.estados)
                        combo.setCurrentText(novo_nome)
                        combo.blockSignals(False)
            elif acao == excluir_acao:
                resposta = QMessageBox.question(
                    self,
                    "Confirmar Exclusão",
                    f"Tem certeza que deseja excluir o estado {estado_atual}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                
                if resposta == QMessageBox.StandardButton.Yes:
                    self.estados.remove(estado_atual)
                    self.salvar_estados()
                    
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Estado")
                    combo.addItems(self.estados)
                    combo.setCurrentText(self.estados[0] if self.estados else "Adicionar Estado")
                    combo.blockSignals(False)
            
    def carregar_consultores(self):
        try:
            with open(resource_path('json/consultores.json'), 'r', encoding='utf-8') as f:
                data = json.load(f)
                consultores = data.get('consultores', [])
                consultores.sort()
                return consultores
        except FileNotFoundError:
            return ["KLEYTON DE PÁDUA"]
    
    def salvar_consultores(self):
        with open(resource_path('json/consultores.json'), 'w', encoding='utf-8') as f:
            json.dump({'consultores': self.consultores}, f, ensure_ascii=False, indent=4)
    
    def consultor_changed(self, text):
        if text == "Adicionar Consultor":
            novo_consultor, ok = QInputDialog.getText(self, "Novo Consultor", "Nome do consultor:")
            if ok and novo_consultor.strip():
                novo_consultor = novo_consultor.strip().upper()
                if novo_consultor not in self.consultores:
                    self.consultores.append(novo_consultor)
                    self.salvar_consultores()
                    combo = self.campos["Consultor"]
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Consultor")
                    combo.addItems(self.consultores)
                    combo.setCurrentText(novo_consultor)
                    combo.blockSignals(False)

    def mostrar_menu_consultor(self, pos):
        combo = self.campos["Consultor"]
        consultor_atual = combo.currentText()
        
        if consultor_atual != "Adicionar Consultor":
            menu = QMenu()
            editar_acao = menu.addAction("Editar Consultor")
            excluir_acao = menu.addAction("Excluir Consultor")
            acao = menu.exec(combo.mapToGlobal(pos))
            
            if acao == editar_acao:
                novo_nome, ok = QInputDialog.getText(
                    self, "Editar Consultor",
                    "Novo nome do consultor:",
                    text=consultor_atual
                )
                
                if ok and novo_nome.strip():
                    novo_nome = novo_nome.strip().upper()
                    if novo_nome != consultor_atual:
                        # Remove o nome antigo e adiciona o novo
                        self.consultores.remove(consultor_atual)
                        self.consultores.append(novo_nome)
                        self.salvar_consultores()
                        
                        # Atualiza o combo
                        combo.blockSignals(True)
                        combo.clear()
                        combo.addItem("Adicionar Consultor")
                        combo.addItems(self.consultores)
                        combo.setCurrentText(novo_nome)
                        combo.blockSignals(False)
            elif acao == excluir_acao:
                resposta = QMessageBox.question(
                    self,
                    "Confirmar Exclusão",
                    f"Tem certeza que deseja excluir o consultor {consultor_atual}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                
                if resposta == QMessageBox.StandardButton.Yes:
                    self.consultores.remove(consultor_atual)
                    self.salvar_consultores()
                    
                    # Atualiza o combo
                    combo.blockSignals(True)
                    combo.clear()
                    combo.addItem("Adicionar Consultor")
                    combo.addItems(self.consultores)
                    combo.setCurrentText(self.consultores[0] if self.consultores else "Adicionar Consultor")
                    combo.blockSignals(False)

def main():
    app = QApplication(sys.argv)
    window = PropostaWindow()
    window.showMaximized()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()