import sys
import sqlite3
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QTableWidget, QTableWidgetItem, QComboBox, QDialog, QFormLayout, QAbstractItemView, QDateEdit, QFileDialog
from PyQt5.QtGui import QColor, QBrush
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5 import QtGui
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# Conexão com o banco de dados
conn = sqlite3.connect('almoxarifado.db')
c = conn.cursor()

# Criação da tabela de produtos no banco de dados
c.execute('''CREATE TABLE IF NOT EXISTS produtos (
                nome TEXT NOT NULL,
                categoria TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                quantidade_minima INTEGER NOT NULL,
                vencimento TEXT NOT NULL,
                data_registro TEXT NOT NULL,
                retirada TEXT
            )''')

# Criação da tabela de retiradas no banco de dados
c.execute('''CREATE TABLE IF NOT EXISTS retiradas (
                produto TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                data TEXT NOT NULL,
                responsavel TEXT NOT NULL
            )''')
            
# Criação da tabela de entradas no banco de dados
c.execute('''CREATE TABLE IF NOT EXISTS entradas (
                produto TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                data TEXT NOT NULL,
                responsavel TEXT NOT NULL
            )''')
            
                        
class EditDialog(QDialog):
    def __init__(self, produto, main_window, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Editar Item")
        self.layout = QVBoxLayout(self)
        self.produto = produto
        self.main_window = main_window

        self.form_layout = QFormLayout()
        self.layout.addLayout(self.form_layout)

        self.name_input = QLineEdit()
        self.name_input.setText(produto[0])
        self.form_layout.addRow("Nome:", self.name_input)

        self.category_combo = QComboBox()
        self.category_combo.addItem("Itens de Escritório")
        self.category_combo.addItem("Itens Médicos")
        self.form_layout.addRow("Categoria:", self.category_combo)

        self.quantity_input = QLineEdit()
        self.quantity_input.setText(str(produto[2]))
        self.form_layout.addRow("Quantidade:", self.quantity_input)

        self.min_quantity_input = QLineEdit()
        self.min_quantity_input.setText(str(produto[3]))
        self.form_layout.addRow("Quantidade Mínima:", self.min_quantity_input)

        self.expiry_date_layout = QHBoxLayout()
        self.form_layout.addRow("Data de Vencimento:", self.expiry_date_layout)

        self.day_label = QLabel("Dia (opcional):")
        self.day_input = QLineEdit()
        self.expiry_date_layout.addWidget(self.day_label)
        self.expiry_date_layout.addWidget(self.day_input)

        self.month_label = QLabel("Mês:")
        self.month_input = QLineEdit()
        self.expiry_date_layout.addWidget(self.month_label)
        self.expiry_date_layout.addWidget(self.month_input)

        self.year_label = QLabel("Ano:")
        self.year_input = QLineEdit()
        self.expiry_date_layout.addWidget(self.year_label)
        self.expiry_date_layout.addWidget(self.year_input)

        self.register_date_layout = QHBoxLayout()
        self.form_layout.addRow("Data de Registro:", self.register_date_layout)

        self.register_day_label = QLabel("Dia:")
        self.register_day_input = QLineEdit()
        self.register_date_layout.addWidget(self.register_day_label)
        self.register_date_layout.addWidget(self.register_day_input)

        self.register_month_label = QLabel("Mês:")
        self.register_month_input = QLineEdit()
        self.register_date_layout.addWidget(self.register_month_label)
        self.register_date_layout.addWidget(self.register_month_input)

        self.register_year_label = QLabel("Ano:")
        self.register_year_input = QLineEdit()
        self.register_date_layout.addWidget(self.register_year_label)
        self.register_date_layout.addWidget(self.register_year_input)

        self.button_layout = QHBoxLayout()
        self.layout.addLayout(self.button_layout)

        self.save_button = QPushButton("Salvar")
        self.save_button.clicked.connect(self.save_changes)
        self.button_layout.addWidget(self.save_button)

        self.delete_button = QPushButton("Excluir")
        self.delete_button.clicked.connect(self.delete_product)
        self.button_layout.addWidget(self.delete_button)

        # Preencher campos de dia, mês e ano corretamente
        vencimento = produto[4].split("/")
        if len(vencimento) > 0:
            self.day_input.setText(vencimento[0])
        if len(vencimento) > 1:
            self.month_input.setText(vencimento[1])
        if len(vencimento) > 2:
            self.year_input.setText(vencimento[2])

        data_registro = produto[6].split("/")
        if len(data_registro) > 0:
            self.register_day_input.setText(data_registro[0])
        if len(data_registro) > 1:
            self.register_month_input.setText(data_registro[1])
        if len(data_registro) > 2:
            self.register_year_input.setText(data_registro[2])

    def save_changes(self):
        nome = self.name_input.text()
        categoria = self.category_combo.currentText()
        quantidade = self.quantity_input.text()
        min_quantidade = self.min_quantity_input.text()
        dia_vencimento = self.day_input.text()
        mes_vencimento = self.month_input.text()
        ano_vencimento = self.year_input.text()
        dia_registro = self.register_day_input.text()
        mes_registro = self.register_month_input.text()
        ano_registro = self.register_year_input.text()

        if nome and categoria and quantidade and min_quantidade and mes_vencimento and ano_vencimento and mes_registro and ano_registro:
            try:
                quantidade = int(quantidade)
                min_quantidade = int(min_quantidade)
                vencimento = ""
                if not dia_vencimento:
                    dia_vencimento = "01"
                vencimento += f"{dia_vencimento}/"
                vencimento += f"{mes_vencimento}/{ano_vencimento}"
                data_registro = ""
                if not dia_registro:
                    dia_registro = "01"
                data_registro += f"{dia_registro}/"
                data_registro += f"{mes_registro}/{ano_registro}"
                c.execute("UPDATE produtos SET nome=?, categoria=?, quantidade=?, quantidade_minima=?, vencimento=?, data_registro=? WHERE nome=?",
                          (nome, categoria, quantidade, min_quantidade, vencimento, data_registro, self.produto[0]))
                conn.commit()
                QMessageBox.information(self, "Sucesso", "Alterações salvas com sucesso!")
                self.accept()
            except ValueError:
                QMessageBox.warning(self, "Erro", "Quantidade inválida. Insira um valor numérico.")
        else:
            QMessageBox.warning(self, "Erro", "Preencha todos os campos obrigatórios.")

    def delete_product(self):
        reply = QMessageBox.question(self, "Excluir Item", "Tem certeza que deseja excluir este item?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            c.execute("DELETE FROM produtos WHERE nome=?", (self.produto[0],))
            c.execute("DELETE FROM retiradas WHERE produto=?", (self.produto[0],))
            conn.commit()
            QMessageBox.information(self, "Sucesso", "Item excluído com sucesso!")
            self.accept()


class WithdrawalDialog(QDialog):
    def __init__(self, quantidade_maxima, produto, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Registrar Retirada")
        self.setFixedSize(655, 500)  # Define o tamanho fixo da janela
        self.layout = QVBoxLayout(self)

        self.quantity_label = QLabel("Quantidade:")
        self.layout.addWidget(self.quantity_label)

        self.quantity_input = QLineEdit()
        self.layout.addWidget(self.quantity_input)

        self.date_label = QLabel("Data:")
        self.layout.addWidget(self.date_label)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.layout.addWidget(self.date_edit)

        self.responsavel_label = QLabel("Responsável:")
        self.layout.addWidget(self.responsavel_label)

        self.responsavel_input = QLineEdit()
        self.layout.addWidget(self.responsavel_input)

        self.button_layout = QHBoxLayout()
        self.layout.addLayout(self.button_layout)

        self.save_button = QPushButton("Salvar")
        self.save_button.clicked.connect(self.accept)
        self.button_layout.addWidget(self.save_button)

        self.quantity_maxima = quantidade_maxima
        self.quantity_input.setValidator(QtGui.QIntValidator(1, quantidade_maxima, self))

    def accept(self):
        quantidade_retirada = self.quantity_input.text()
        if quantidade_retirada.isdigit() and int(quantidade_retirada) > 0:
            super().accept()
        else:
            QMessageBox.warning(self, "Erro", "A quantidade de retirada deve ser um valor inteiro maior que zero.")
            
class EntryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Registrar Entrada")
        self.setFixedSize(655, 500)  # Define o tamanho fixo da janela
        self.layout = QVBoxLayout(self)

        self.quantity_label = QLabel("Quantidade:")
        self.layout.addWidget(self.quantity_label)

        self.quantity_input = QLineEdit()
        self.layout.addWidget(self.quantity_input)

        self.date_label = QLabel("Data:")
        self.layout.addWidget(self.date_label)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.layout.addWidget(self.date_edit)

        self.responsavel_label = QLabel("Responsável:")
        self.layout.addWidget(self.responsavel_label)

        self.responsavel_input = QLineEdit()
        self.layout.addWidget(self.responsavel_input)

        self.button_layout = QHBoxLayout()
        self.layout.addLayout(self.button_layout)

        self.save_button = QPushButton("Salvar")
        self.save_button.clicked.connect(self.accept)
        self.button_layout.addWidget(self.save_button)

        self.quantity_input.setValidator(QtGui.QIntValidator(1, 99999, self))

    def accept(self):
        quantidade_entrada = self.quantity_input.text()
        if quantidade_entrada.isdigit() and int(quantidade_entrada) > 0:
            super().accept()
        else:
            QMessageBox.warning(self, "Erro", "A quantidade de entrada deve ser um valor inteiro maior que zero.")
          

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Sistema de Gerenciamento de Almoxarifado")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.label = QLabel("Nome do item para adicionar:")
        self.layout.addWidget(self.label)

        self.name_input = QLineEdit()
        self.layout.addWidget(self.name_input)

        self.category_label = QLabel("Categoria:")
        self.layout.addWidget(self.category_label)

        self.category_combo = QComboBox()
        self.category_combo.addItem("Itens de Escritório")
        self.category_combo.addItem("Itens Médicos")
        self.layout.addWidget(self.category_combo)

        self.quantity_label = QLabel("Quantidade:")
        self.layout.addWidget(self.quantity_label)

        self.quantity_input = QLineEdit()
        self.layout.addWidget(self.quantity_input)

        self.min_quantity_label = QLabel("Quantidade Mínima:")
        self.layout.addWidget(self.min_quantity_label)

        self.min_quantity_input = QLineEdit()
        self.layout.addWidget(self.min_quantity_input)

        self.expiry_date_layout = QHBoxLayout()
        self.layout.addLayout(self.expiry_date_layout)

        self.expiry_date_label = QLabel("Data de Vencimento:")
        self.expiry_date_layout.addWidget(self.expiry_date_label)

        self.day_label = QLabel("Dia (opcional):")
        self.day_input = QLineEdit()
        self.expiry_date_layout.addWidget(self.day_label)
        self.expiry_date_layout.addWidget(self.day_input)

        self.month_label = QLabel("Mês:")
        self.month_input = QLineEdit()
        self.expiry_date_layout.addWidget(self.month_label)
        self.expiry_date_layout.addWidget(self.month_input)

        self.year_label = QLabel("Ano:")
        self.year_input = QLineEdit()
        self.expiry_date_layout.addWidget(self.year_label)
        self.expiry_date_layout.addWidget(self.year_input)

        self.register_date_layout = QHBoxLayout()
        self.layout.addLayout(self.register_date_layout)

        self.register_date_label = QLabel("Data de Registro:")
        self.register_date_layout.addWidget(self.register_date_label)

        self.register_day_label = QLabel("Dia:")
        self.register_day_input = QLineEdit()
        self.register_date_layout.addWidget(self.register_day_label)
        self.register_date_layout.addWidget(self.register_day_input)

        self.register_month_label = QLabel("Mês:")
        self.register_month_input = QLineEdit()
        self.register_date_layout.addWidget(self.register_month_label)
        self.register_date_layout.addWidget(self.register_month_input)

        self.register_year_label = QLabel("Ano:")
        self.register_year_input = QLineEdit()
        self.register_date_layout.addWidget(self.register_year_label)
        self.register_date_layout.addWidget(self.register_year_input)

        self.add_button = QPushButton("Adicionar Item")
        self.add_button.clicked.connect(self.add_product)
        self.layout.addWidget(self.add_button)

        self.edit_button = QPushButton("Editar Item")
        self.edit_button.clicked.connect(self.edit_product)
        self.layout.addWidget(self.edit_button)

        self.delete_button = QPushButton("Excluir Item")
        self.delete_button.clicked.connect(self.delete_product)
        self.layout.addWidget(self.delete_button)

        self.withdrawal_button = QPushButton("Registrar Retirada")
        self.withdrawal_button.clicked.connect(self.open_withdrawal_dialog)
        self.layout.addWidget(self.withdrawal_button)

        self.entry_button = QPushButton("Registrar Entrada")
        self.entry_button.clicked.connect(self.open_entry_dialog)
        self.layout.addWidget(self.entry_button)
        
        self.report_button = QPushButton("Gerar Relatório")
        self.report_button.clicked.connect(self.generate_report)
        self.layout.addWidget(self.report_button)
  
  
        self.category_filter_combo = QComboBox()
        self.category_filter_combo.addItem("Todas as Categorias")
        self.category_filter_combo.addItem("Itens de Escritório")
        self.category_filter_combo.addItem("Itens Médicos")
        self.category_filter_combo.currentIndexChanged.connect(self.load_products)
        self.layout.addWidget(self.category_filter_combo)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(
            ["Nome", "Categoria", "Quantidade", "Quantidade Mínima", "Data de Vencimento", "Entrada", "Retirada"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Tabela somente leitura
        self.layout.addWidget(self.table)

        self.legend_label = QLabel("Legenda:")
        self.layout.addWidget(self.legend_label)

        self.legend_layout = QHBoxLayout()
        self.layout.addLayout(self.legend_layout)

        self.legend_expired_label = QLabel(" Vencidos")
        self.legend_expired_label.setStyleSheet("background-color: red; color: white; padding: 5px;")
        self.legend_layout.addWidget(self.legend_expired_label)

        self.legend_near_expiry_label = QLabel("Próx. do Venc.")
        self.legend_near_expiry_label.setStyleSheet("background-color: yellow; padding: 5px;")
        self.legend_layout.addWidget(self.legend_near_expiry_label)

        self.legend_valid_label = QLabel(" Válidos")
        self.legend_valid_label.setStyleSheet("background-color: green; color: white; padding: 5px;")
        self.legend_layout.addWidget(self.legend_valid_label)

        self.legend_low_quantity_label = QLabel("Item Acabando")
        self.legend_low_quantity_label.setStyleSheet("background-color: blue; color: white; padding: 5px;")
        self.legend_layout.addWidget(self.legend_low_quantity_label)

        self.legend_zero_quantity_label = QLabel(" Item zerado")
        self.legend_zero_quantity_label.setStyleSheet("background-color: purple; color: white; padding: 5px;")
        self.legend_layout.addWidget(self.legend_zero_quantity_label)

        self.timer = QTimer()
        self.timer.timeout.connect(self.toggle_blink)
        self.timer.start(500)  # Intervalo de 500 ms para alternar o estilo

        self.load_products()
        self.timer_row = -1  # Índice da linha que está piscando
        self.timer_color = QColor("purple")  # Cor roxa para piscar

    def load_products(self):
        self.table.setRowCount(0)
        selected_category = self.category_filter_combo.currentText()

        if selected_category == "Todas as Categorias":
            c.execute("SELECT nome, categoria, quantidade, quantidade_minima, vencimento, retirada, data_registro FROM produtos")
        else:
            c.execute("SELECT nome, categoria, quantidade, quantidade_minima, vencimento, retirada, data_registro FROM produtos WHERE categoria=?", (selected_category,))

        products = c.fetchall()

        for row_number, row_data in enumerate(products):
            self.table.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                item = QTableWidgetItem(str(data))
                self.table.setItem(row_number, column_number, item)

                vencimento = row_data[4]
                vencimento_date = QDate.fromString(vencimento, "d/M/yyyy")
                if not vencimento_date.isValid():
                    vencimento_date = QDate.fromString(vencimento, "M/yyyy")
                retirada = row_data[5]
                if retirada:
                    retirada_date = QDate.fromString(retirada, "dd/MM/yyyy")
                today = QDate.currentDate()
                if vencimento_date < today:
                    item.setBackground(QColor("red"))
                    item.setForeground(QColor("white"))
                elif vencimento_date.addMonths(-2) <= today:
                    item.setBackground(QColor("yellow"))
                else:
                    item.setBackground(QColor("green"))
                    item.setForeground(QColor("white"))

                quantidade = row_data[2]
                quantidade_minima = row_data[3]
                if quantidade == 0:
                    item.setBackground(QColor("purple"))
                    item.setForeground(QColor("white"))
                elif quantidade <= quantidade_minima:
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)

                    if self.timer.isActive():
                        item.setBackground(QColor("blue"))
                        item.setForeground(QColor("white"))
                    else:
                        item.setBackground(QColor("white"))
                        item.setForeground(QColor("blue"))

            button = QPushButton("Detalhes")
            button.clicked.connect(lambda _, p=row_data[0]: self.open_withdrawal_details(p))
            self.table.setCellWidget(row_number, 5, button)

            button = QPushButton("Detalhes")
            button.clicked.connect(lambda _, p=row_data[0]: self.open_entry_details(p))
            self.table.setCellWidget(row_number, 6, button)

        self.table.resizeColumnsToContents()

    def toggle_blink(self):
        row = -1
        for row in range(self.table.rowCount()):
            quantidade = int(self.table.item(row, 2).text())
            quantidade_minima = int(self.table.item(row, 3).text())
            if quantidade == 0:
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item.background().color() == QColor("white"):
                        item.setBackground(QColor("purple"))
                        item.setForeground(QColor("white"))
                    else:
                        item.setBackground(QColor("white"))
                        item.setForeground(QColor("purple"))
            elif quantidade <= quantidade_minima:
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item.background().color() == QColor("white"):
                        item.setBackground(QColor("blue"))
                        item.setForeground(QColor("white"))
                    else:
                        item.setBackground(QColor("white"))
                        item.setForeground(QColor("blue"))

        if row == self.timer_row:
            self.timer_row = -1
        else:
            self.timer_row = row

        if self.table.selectedItems():
            self.table.itemSelectionChanged.emit()

    def add_product(self):
        nome = self.name_input.text()
        categoria = self.category_combo.currentText()
        quantidade = self.quantity_input.text()
        min_quantidade = self.min_quantity_input.text()
        dia_vencimento = self.day_input.text()
        mes_vencimento = self.month_input.text()
        ano_vencimento = self.year_input.text()
        dia_registro = self.register_day_input.text()
        mes_registro = self.register_month_input.text()
        ano_registro = self.register_year_input.text()

        if nome and categoria and quantidade and min_quantidade and mes_vencimento and ano_vencimento and dia_registro and mes_registro and ano_registro:
            try:
                quantidade = int(quantidade)
                min_quantidade = int(min_quantidade)
                vencimento = ""
                if not dia_vencimento:
                    dia_vencimento = "01"
                vencimento += f"{dia_vencimento}/"
                vencimento += f"{mes_vencimento}/{ano_vencimento}"
                data_registro = ""
                if not dia_registro:
                    dia_registro = "01"
                data_registro += f"{dia_registro}/"
                data_registro += f"{mes_registro}/{ano_registro}"
                c.execute("INSERT INTO produtos VALUES (?, ?, ?, ?, ?, ?, ?)",
                          (nome, categoria, quantidade, min_quantidade, vencimento, data_registro, None))                                                   
                if quantidade <= 0:
                      self.timer_row = self.table.rowCount() - 1

                conn.commit()
                QMessageBox.information(self, "Sucesso", "Item adicionado com sucesso!")
                self.clear_input_fields()
                self.load_products()
            except ValueError:
                QMessageBox.warning(self, "Erro", "Quantidade inválida. Insira um valor numérico.")
        else:
            QMessageBox.warning(self, "Erro", "Preencha todos os campos obrigatórios.")

    def edit_product(self):
        selected_items = self.table.selectedItems()
        if len(selected_items) == 0:
            QMessageBox.warning(self, "Erro", "Selecione um item para editar.")
            return

        selected_row = selected_items[0].row()
        nome = self.table.item(selected_row, 0).text()
        categoria = self.table.item(selected_row, 1).text()
        quantidade = self.table.item(selected_row, 2).text()
        min_quantidade = self.table.item(selected_row, 3).text()
        vencimento = self.table.item(selected_row, 4).text()
        data_registro = self.table.item(selected_row, 6).text()

        produto = (nome, categoria, quantidade, min_quantidade, vencimento, None, data_registro)

        dialog = EditDialog(produto, self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_products()

    def delete_product(self):
        selected_items = self.table.selectedItems()
        if len(selected_items) == 0:
            QMessageBox.warning(self, "Erro", "Selecione um item para excluir.")
            return

        selected_row = selected_items[0].row()
        nome = self.table.item(selected_row, 0).text()

        reply = QMessageBox.question(self, "Excluir Item", "Tem certeza que deseja excluir este item?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            c.execute("DELETE FROM produtos WHERE nome=?", (nome,))
            c.execute("DELETE FROM retiradas WHERE produto=?", (nome,))
            conn.commit()
            QMessageBox.information(self, "Sucesso", "Item excluído com sucesso!")
            self.load_products()

    def open_withdrawal_dialog(self):
        selected_items = self.table.selectedItems()
        if len(selected_items) == 0:
            QMessageBox.warning(self, "Erro", "Selecione um item para registrar a retirada.")
            return

        selected_row = selected_items[0].row()
        quantidade_atual = int(self.table.item(selected_row, 2).text())
        produto = self.table.item(selected_row, 0).text()

        dialog = WithdrawalDialog(quantidade_atual, produto, self)
        if dialog.exec_() == QDialog.Accepted:
            quantidade_retirada = int(dialog.quantity_input.text())
            if quantidade_retirada <= quantidade_atual:  # Verifica se a quantidade de retirada é menor ou igual à quantidade atual em estoque
               data = dialog.date_edit.date().toString("dd/MM/yyyy")
               responsavel = dialog.responsavel_input.text()

               c.execute("UPDATE produtos SET quantidade=quantidade-? WHERE nome=?", (quantidade_retirada, produto))
               c.execute("INSERT INTO retiradas VALUES (?, ?, ?, ?)", (produto, quantidade_retirada, data, responsavel))
               conn.commit()

               QMessageBox.information(self, "Sucesso", "Retirada registrada com sucesso!")
               self.load_products()
            else:
               QMessageBox.warning(self, "Erro", "A quantidade de retirada é maior do que a quantidade disponível em estoque.")
               

    def open_withdrawal_details(self, produto):
        c.execute("SELECT quantidade, data, responsavel FROM retiradas WHERE produto=?", (produto,))
        retiradas = c.fetchall()

        details_dialog = QDialog(self)
        details_dialog.setWindowTitle("Detalhes da Retirada")
        details_dialog.setFixedSize(655, 500)

        details_layout = QVBoxLayout(details_dialog)

        table = QTableWidget(details_dialog)
        table.setRowCount(len(retiradas))
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Quantidade", "Data", "Responsável"])
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for row_number, row_data in enumerate(retiradas):
            for column_number, data in enumerate(row_data):
                item = QTableWidgetItem(str(data))
                table.setItem(row_number, column_number, item)

        details_layout.addWidget(table)

        details_dialog.exec_()
        
    def open_entry_dialog(self):
        selected_items = self.table.selectedItems()
        if len(selected_items) == 0:
            QMessageBox.warning(self, "Erro", "Selecione um item para registrar a entrada.")
            return

        selected_row = selected_items[0].row()
        quantidade_maxima = int(self.table.item(selected_row, 2).text())
        produto = self.table.item(selected_row, 0).text()

        dialog = EntryDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            quantidade_entrada = int(dialog.quantity_input.text())
            data = dialog.date_edit.date().toString("dd/MM/yyyy")
            responsavel = dialog.responsavel_input.text()

            c.execute("UPDATE produtos SET quantidade=quantidade+? WHERE nome=?", (quantidade_entrada, produto))
            c.execute("INSERT INTO entradas VALUES (?, ?, ?, ?)", (produto, quantidade_entrada, data, responsavel))
            conn.commit()

            QMessageBox.information(self, "Sucesso", "Entrada registrada com sucesso!")
            self.load_products()
        
    def open_entry_details(self, produto):
        c.execute("SELECT quantidade, data, responsavel FROM entradas WHERE produto=?", (produto,))
        entradas = c.fetchall()

        details_dialog = QDialog(self)
        details_dialog.setWindowTitle("Detalhes da entrada")
        details_dialog.setFixedSize(655, 500)

        details_layout = QVBoxLayout(details_dialog)

        table = QTableWidget(details_dialog)
        table.setRowCount(len(entradas))
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Quantidade", "Data", "Responsável"])
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for row_number, row_data in enumerate(entradas):
            for column_number, data in enumerate(row_data):
                item = QTableWidgetItem(str(data))
                table.setItem(row_number, column_number, item)

        details_layout.addWidget(table)

        details_dialog.exec_() 

    def generate_report(self):
        dialog = QFileDialog()
        file_path, _ = dialog.getSaveFileName(self, "Salvar Relatório", "", "Arquivos Excel (*.xlsx)")

        if file_path:
           if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

            workbook = openpyxl.Workbook()

            # Planilha de Produtos
            products_sheet = workbook.active
            products_sheet.title = "Itens"

            # Definir estilo para o cabeçalho
            header_style = openpyxl.styles.NamedStyle(name="header_style")
            header_style.alignment = Alignment(horizontal="center", vertical="center")

            # Preencher cabeçalho
            products_sheet.merge_cells("A1:G1")
            header_cell = products_sheet.cell(row=1, column=1, value="Almoxarifado Caps Paulo da Portela - Relatório de Itens - Prefeitura do Rio de Janeiro")
            header_cell.style = header_style

            # Preencher nomes das colunas
            product_columns = ["Nome", "Categoria", "Quantidade", "Quantidade Mínima", "Vencimento", "Data do Primeiro Registro"]
            for col, column_name in enumerate(product_columns, start=1):
                cell = products_sheet.cell(row=3, column=col, value=column_name)
                cell.style = header_style

            # Preencher dados dos produtos
            c.execute("SELECT * FROM produtos")
            products = c.fetchall()
            for row, product in enumerate(products, start=4):
                for col, data in enumerate(product, start=1):
                    cell = products_sheet.cell(row=row, column=col, value=data)

            # Ajustar largura das colunas
            for column in range(1, len(product_columns) + 1):
                column_letter = get_column_letter(column)
                products_sheet.column_dimensions[column_letter].width = 15

            # Planilha de Retiradas
            withdrawals_sheet = workbook.create_sheet(title="Retiradas")

            # Preencher cabeçalho
            withdrawals_sheet.merge_cells("A1:D1")
            header_cell = withdrawals_sheet.cell(row=1, column=1, value="Almoxarifado Caps Paulo da Portela - Relatório de Retiradas - Prefeitura do Rio de Janeiro")
            header_cell.style = header_style

            # Preencher nomes das colunas
            withdrawal_columns = ["Produto", "Quantidade", "Data", "Responsável"]
            for col, column_name in enumerate(withdrawal_columns, start=1):
                cell = withdrawals_sheet.cell(row=3, column=col, value=column_name)
                cell.style = header_style

            # Preencher dados das retiradas
            c.execute("SELECT * FROM retiradas")
            withdrawals = c.fetchall()
            for row, withdrawal in enumerate(withdrawals, start=4):
                for col, data in enumerate(withdrawal, start=1):
                    cell = withdrawals_sheet.cell(row=row, column=col, value=data)

            # Ajustar largura das colunas
            for column in range(1, len(withdrawal_columns) + 1):
                column_letter = get_column_letter(column)
                withdrawals_sheet.column_dimensions[column_letter].width = 15

            # Planilha de Entradas
            entries_sheet = workbook.create_sheet(title="Entradas")

            # Preencher cabeçalho
            entries_sheet.merge_cells("A1:D1")
            header_cell = entries_sheet.cell(row=1, column=1, value="Almoxarifado Caps Paulo da Portela - Relatório de Entradas - Prefeitura do Rio de Janeiro")
            header_cell.style = header_style

            # Preencher nomes das colunas
            entry_columns = ["Produto", "Quantidade", "Data", "Responsável"]
            for col, column_name in enumerate(entry_columns, start=1):
                cell = entries_sheet.cell(row=3, column=col, value=column_name)
                cell.style = header_style

            # Preencher dados das entradas
            c.execute("SELECT * FROM entradas")
            entries = c.fetchall()
            for row, entry in enumerate(entries, start=4):
                for col, data in enumerate(entry, start=1):
                    cell = entries_sheet.cell(row=row, column=col, value=data)

            # Ajustar largura das colunas
            for column in range(1, len(entry_columns) + 1):
                column_letter = get_column_letter(column)
                entries_sheet.column_dimensions[column_letter].width = 15

            # Planilha de Edições
            edits_sheet = workbook.create_sheet(title="Edições")

            # Preencher cabeçalho
            edits_sheet.merge_cells("A1:E1")
            header_cell = edits_sheet.cell(row=1, column=1, value="Almoxarifado Caps Paulo da Portela - Relatório de Edições - Prefeitura do Rio de Janeiro")
            header_cell.style = header_style

            # Preencher nomes das colunas
            edit_columns = ["Produto", "Quantidade Anterior", "Quantidade Atual", "Data", "Responsável"]
            for col, column_name in enumerate(edit_columns, start=1):
                cell = edits_sheet.cell(row=3, column=col, value=column_name)
                cell.style = header_style

            # Salvar o arquivo
            workbook.save(file_path)
            QMessageBox.information(self, "Relatório Gerado", "O relatório foi gerado com sucesso!")
     

    def clear_input_fields(self):
        self.name_input.clear()
        self.quantity_input.clear()
        self.min_quantity_input.clear()
        self.day_input.clear()
        self.month_input.clear()
        self.year_input.clear()
        self.register_day_input.clear()
        self.register_month_input.clear()
        self.register_year_input.clear()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())
