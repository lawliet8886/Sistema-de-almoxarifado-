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
         self.category_combo.addItem("Itens de Limpeza")   
         self.category_combo.addItem("Medicações") 
         self.category_combo.addItem("Itens Importados") 
  
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
         self.month_input.textChanged.connect(self.check_month_input) 
         self.expiry_date_layout.addWidget(self.month_label) 
         self.expiry_date_layout.addWidget(self.month_input)         
         # Conectar o sinal `returnPressed` com o slot `focusNextChild`. 
         self.name_input.returnPressed.connect(self.focusNextChild) 
         self.quantity_input.returnPressed.connect(self.focusNextChild) 
         self.min_quantity_input.returnPressed.connect(self.focusNextChild) 
         self.day_input.returnPressed.connect(self.focusNextChild) 
         self.month_input.returnPressed.connect(self.focusNextChild)         
  
         # Configurar a ordem de tabulação. 
         self.setTabOrder(self.name_input, self.quantity_input) 
         self.setTabOrder(self.quantity_input, self.min_quantity_input) 
         self.setTabOrder(self.min_quantity_input, self.day_input)    
         self.setTabOrder(self.day_input, self.month_input) 
  
         self.year_label = QLabel("Ano:") 
         self.year_input = QLineEdit() 
         self.year_input.textChanged.connect(self.check_year_input) 
         self.expiry_date_layout.addWidget(self.year_label) 
         self.expiry_date_layout.addWidget(self.year_input) 
  
         self.register_date_layout = QHBoxLayout() 
         self.layout.addLayout(self.register_date_layout) 
  
         self.register_date_label = QLabel("Data de Registro:") 
         self.register_date_layout.addWidget(self.register_date_label) 
  
         self.unit_value_input = QLineEdit(self) 
         self.unit_value_input.   setPlaceholderText("Valor Unitário") 
         self.unit_value_input.   setValidator(QIntValidator(0, 999999999, self)) 
         self.layout.addWidget(self.unit_value_input) 
  
         self.register_day_label = QLabel("Dia:") 
         self.register_day_input = QLineEdit() 
         self.register_day_input.textChanged.connect(self.check_register_day_input) 
         self.register_date_layout.addWidget(self.register_day_label) 
         self.register_date_layout.addWidget(self.register_day_input) 
  
         self.register_month_label = QLabel("Mês:") 
         self.register_month_input = QLineEdit() 
         self.register_month_input.textChanged.connect(self.check_register_month_input) 
         self.register_date_layout.addWidget(self.register_month_label) 
         self.register_date_layout.addWidget(self.register_month_input) 
  
         self.register_year_label = QLabel("Ano:") 
         self.register_year_input = QLineEdit() 
         self.register_year_input.textChanged.connect(self.check_register_year_input) 
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
  
         self.import_button = QPushButton("Importar") 
         self.import_button.clicked.connect(self.import_from_excel) 
         self.layout.addWidget(self.import_button) 
         self.import_input = QLineEdit() 
  
         self.search_button =    QPushButton("Buscar") 
         self.search_button.clicked.connect(self.search_products) 
         self.layout.addWidget(self.search_button) 
         self.search_input = QLineEdit() 
         self.layout.addWidget(self.search_input)   
  
         self.total_value_label = QLabel() 
         self.layout.addWidget(self.total_value_label) 
  
  
         self.category_filter_combo = QComboBox() 
         self.category_filter_combo.addItem("Todas as Categorias") 
         self.category_filter_combo.addItem("Itens de Escritório") 
         self.category_filter_combo.addItem("Itens Médicos") 
         self.category_filter_combo.addItem("Itens de Limpeza") 
         self.category_filter_combo.addItem("Medicações") 
         self.category_filter_combo.addItem("Itens Importados") 
  
         self.category_filter_combo.currentIndexChanged.connect(self.load_products) 
         self.layout.addWidget(self.category_filter_combo) 
         self.table = QTableWidget() 
         self.table.setColumnCount(9) 
         self.table.setHorizontalHeaderLabels( 
             ["Nome", "Categoria", "Quantidade", "Quantidade Mínima", "Data de Vencimento", "Entrada", "Retirada", "Valor Unitário", "Valor Total"]) 
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
     def check_month_input(self, text): 
         if len(text) == 2: 
             self.year_input.setFocus() 
  
     def check_year_input(self, text): 
         if len(text) == 4: 
             self.register_day_input.setFocus() 
  
     def check_register_day_input(self, text): 
         if len(text) == 2: 
             self.register_month_input.setFocus() 
  
     def check_register_month_input(self, text): 
         if len(text) == 2: 
             self.register_year_input.setFocus() 
  
     def check_register_year_input(self, text): 
         if len(text) == 4: 
             self.unit_value_input.setFocus()                     
  
     def keyPressEvent(self, event): 
         if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter: 
             focused_widget = self.focusWidget() 
             if isinstance(focused_widget, QPushButton): 
                 focused_widget.click() 
  
     def import_from_excel(self): 
       try: 
     # Exibir caixa de diálogo com informações sobre o formato do arquivo Excel 
         message = ("O arquivo Excel deve ter os seguintes nomes de colunas na terceira linha( não precisam estar nessa ordem; Não precisa estarem todas preenchidas; Não precisa que o arquivo tenha todas essas colunas, pois os valores faltantes serão preenchidos com uma informção padrão do código):\n" 
                "Nome, Categoria, Quantidade, Quantidade Mínima, Vencimento, Data do Primeiro Registro, Valor Unitário\n" 
                "Deseja continuar com a importação?") 
         reply = QMessageBox.question(self, 'Formato do Arquivo Excel', message,  QMessageBox.Yes, QMessageBox.No) 
         if reply == QMessageBox.Yes: 
         # Continuar com a importação 
             options = QFileDialog.Options() 
             options |= QFileDialog.ReadOnly 
             file_name, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", 
                                                    "Excel Files (*.xlsx);;All Files (*)", options=options) 
             if file_name: 
                 workbook = openpyxl.load_workbook(file_name) 
                 sheet = workbook.active 
             else: 
                     return 
     # Obter os nomes das colunas da primeira linha 
             column_names = [(cell.value if cell.value is not None else '') for cell in sheet[3]] 
  
             for row in sheet.iter_rows(min_row=4, values_only=True):  # Ignorar a primeira linha (cabeçalho) 
         # Criar um dicionário que mapeia os nomes das colunas para os seus respectivos valores 
                 item_data = {column_name: value for column_name, value in zip(column_names, row)} 
  
         # Obter os valores das colunas, usando valores padrão para as colunas que estão faltando 
                 item_name = item_data.get("Nome", "Nome Desconhecido") 
                 categoria = item_data.get("Categoria", "Itens Importados") 
  
                 quantidade = item_data. get("Quantidade", 0) 
                 if quantidade is None or quantidade == '': 
                     quantidade = 0 
                 else: 
                         quantidade = float(quantidade) 
  
                 quantidade_minima = item_data.get("Quantidade Mínima", 0) 
                 if quantidade_minima is None or   quantidade_minima == '': 
                     quantidade_minima = 0 
                 else: 
                     quantidade_minima = float(quantidade_minima) 
  
                 vencimento = item_data.get("Vencimento", datetime.datetime(2100, 1, 1)) 
                 if isinstance(vencimento, datetime.datetime): 
                     vencimento = vencimento.strftime("%d/%m/%Y") 
                 if vencimento is None or vencimento == '': 
                     vencimento = "01/01/2100" 
  
                 entrada = item_data.get("Data do Primeiro Registro", datetime.datetime(2000, 1, 1)) 
                 if isinstance(entrada, datetime.datetime): 
                     entrada = entrada.strftime("%d/%m/%Y") 
                 if entrada is None or entrada == '': 
                     entrada = "01/01/2000" 
  
                 valor_unitario = item_data.get("Valor Unitário", 0) 
                 if valor_unitario is None or valor_unitario == '': 
                     valor_unitario = 0 
                 else: 
                     valor_unitario = float(valor_unitario) 
  
  
             # Calcular o valor total 
                 valor_total = quantidade * valor_unitario 
  
                 c.execute("INSERT INTO produtos (nome, categoria, quantidade, quantidade_minima, vencimento, data_registro, valor_unitario, valor_total) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", 
                           (item_name, categoria, quantidade, quantidade_minima, vencimento, entrada, valor_unitario, valor_total)) 
         conn.commit() 
         self.load_products() 
         self.update_total_value() 
       except Exception as e: 
           traceback.print_exc()  # Imprime o traceback completo do erro 
           QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao limpar os campos: {str(e)}")         
  
  
     def update_total_value(self): 
       try: 
         c.execute("SELECT SUM(valor_total) FROM produtos") 
 # No método update_total_value 
         total_value = c.fetchone()[0] 
         if total_value is None: 
             total_value = 0 
         total_value_str = 'R$ {:,.2f}'.format(total_value).replace('.', '#').replace(',', '.').replace('#', ',') 
         self.total_value_label.setText(f"Valor Total do Almoxarifado: {total_value_str}") 
       except Exception as e: 
           traceback.print_exc()  # Imprime o traceback completo do erro 
           QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao limpar os campos: {str(e)}") 
  
     def search_products(self): 
       try: 
         search_text = self.search_input.text()  # Obtém o texto digitado na caixa de busca 
         selected_category = self.    category_filter_combo.currentText() 
  
         if selected_category == "Todas as Categorias": 
             c.execute("SELECT nome, categoria, quantidade, quantidade_minima, vencimento, retirada, data_registro, valor_unitario, valor_total FROM produtos WHERE nome LIKE ?", (f'%{search_text}%',)) 
         else: 
             c.execute("SELECT nome, categoria, quantidade, quantidade_minima, vencimento, retirada, data_registro, valor_unitario, valor_total FROM produtos WHERE nome LIKE ? AND categoria=?", (f'%{search_text}%', selected_category)) 
  
         products = c.fetchall() 
         self.update_table(products) 
  
       except Exception as e: 
           traceback.print_exc()  # Imprime o traceback completo do erro 
           QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao limpar os campos: {str(e)}")         
  
     def update_table(self, products): 
       try: 
  
         self.table.setRowCount(0) 
  
         for row_number, row_data in   enumerate(products): 
             self.table.insertRow(row_number) 
             for column_number, data in enumerate(row_data): 
                 item = QTableWidgetItem(str(data)) 
                 self.table.setItem(row_number, column_number, item) 
  
                 quantidade = row_data[2] 
                 valor_unitario = row_data[7] 
                 valor_total = valor_unitario * quantidade 
                 item = QTableWidgetItem(str(valor_total)) 
                 self.table.setItem(row_number, 8, item) 
  
                 vencimento = row_data[4] 
                 vencimento_date = QDate.fromString(vencimento, "d/M/yyyy") 
                 if not vencimento_date.isValid(): 
                     vencimento_date = QDate.fromString(vencimento, "M/yyyy") 
                 retirada = row_data[5] 
                 if retirada: 
                     retirada_date = QDate.fromString(retirada, "dd/MM/yyyy") 
                 today = QDate.currentDate() 
                 if vencimento_date < today: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("red")) 
                             item.setForeground(QColor("white")) 
                 elif vencimento_date.addMonths(-2) <= today: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("yellow")) 
                 else: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("green")) 
                             item.setForeground(QColor("white")) 
  
                 quantidade = row_data[2] 
                 quantidade_minima = row_data[3] 
                 if quantidade == 0: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("purple")) 
                             item.setForeground(QColor("white")) 
                 elif quantidade <= quantidade_minima: 
                     font = item.font() 
                     font.setBold(True) 
                     item.setFont(font) 
  
                     if self.timer.isActive(): 
                         for column_number in range(self.table.columnCount()): 
                             item = self.table.item(row_number, column_number) 
                             if item is not None: 
                                 item.setBackground(QColor("blue")) 
                                 item.setForeground(QColor("white")) 
                     else: 
                         for column_number in range(self.table.columnCount()): 
                             item = self.table.item(row_number, column_number) 
                             if item is not None: 
                                 item.setBackground(QColor("white")) 
                                 item.setForeground(QColor("blue")) 
  
             button = QPushButton("Detalhes") 
             button.clicked.connect(lambda _, p=row_data[0]: self.open_withdrawal_details(p)) 
             self.table.setCellWidget(row_number, 6, button) 
  
             button = QPushButton("Detalhes") 
             button.clicked.connect(lambda _, p=row_data[0]: self.open_entry_details(p)) 
             self.table.setCellWidget(row_number, 5, button) 
       except Exception as e: 
  
           traceback.print_exc()  # Imprime o traceback completo do erro 
           QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao limpar os campos: {str(e)}")             
  
     def load_products(self): 
       try: 
         self.table.setRowCount(0) 
         selected_category = self.category_filter_combo.currentText() 
  
         if selected_category == "Todas as Categorias": 
             c.execute("SELECT nome, categoria, quantidade, quantidade_minima, vencimento, retirada, data_registro, valor_unitario, valor_total FROM produtos") 
         else: 
             c.execute("SELECT nome, categoria, quantidade, quantidade_minima, vencimento, retirada, data_registro, valor_unitario, valor_total FROM produtos WHERE categoria=?", (selected_category,)) 
  
         products = c.fetchall() 
  
         for row_number, row_data in   enumerate(products): 
             self.table.insertRow(row_number) 
             for column_number, data in enumerate(row_data): 
                 item = QTableWidgetItem(str(data)) 
                 self.table.setItem(row_number, column_number, item) 
  
                 quantidade = row_data[2] 
                 valor_unitario = row_data[7] 
                 valor_total = valor_unitario * quantidade 
                 item = QTableWidgetItem(str(valor_total)) 
                 self.table.setItem(row_number, 8, item) 
  
                 vencimento = row_data[4] 
                 vencimento_date = QDate.fromString(vencimento, "d/M/yyyy") 
                 if not vencimento_date.isValid(): 
                     vencimento_date = QDate.fromString(vencimento, "M/yyyy") 
                 retirada = row_data[5] 
                 if retirada: 
                     retirada_date = QDate.fromString(retirada, "dd/MM/yyyy") 
                 today = QDate.currentDate() 
                 if vencimento_date < today: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("red")) 
                             item.setForeground(QColor("white")) 
                 elif vencimento_date.addMonths(-2) <= today: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("yellow")) 
                 else: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("green")) 
                             item.setForeground(QColor("white")) 
  
                 quantidade = row_data[2] 
                 quantidade_minima = row_data[3] 
                 if quantidade == 0: 
                     for column_number in range(self.table.columnCount()): 
                         item = self.table.item(row_number, column_number) 
                         if item is not None: 
                             item.setBackground(QColor("purple")) 
                             item.setForeground(QColor("white")) 
                 elif quantidade <= quantidade_minima: 
                     font = item.font() 
                     font.setBold(True) 
                     item.setFont(font) 
  
                     if self.timer.isActive(): 
                         for column_number in range(self.table.columnCount()): 
                             item = self.table.item(row_number, column_number) 
                             if item is not None: 
                                 item.setBackground(QColor("blue")) 
                                 item.setForeground(QColor("white")) 
                     else: 
                         for column_number in range(self.table.columnCount()): 
                             item = self.table.item(row_number, column_number) 
                             if item is not None: 
                                 item.setBackground(QColor("white")) 
                                 item.setForeground(QColor("blue")) 
  
             button = QPushButton("Detalhes") 
             button.clicked.connect(lambda _, p=row_data[0]: self.open_withdrawal_details(p)) 
             self.table.setCellWidget(row_number, 6, button) 
  
             button = QPushButton("Detalhes") 
             button.clicked.connect(lambda _, p=row_data[0]: self.open_entry_details(p)) 
             self.table.setCellWidget(row_number, 5, button) 
  
         self.table.resizeColumnsToContents() 
         self.update_total_value() 
       except Exception as e: 
           traceback.print_exc()  # Imprime o traceback completo do erro 
           QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao limpar os campos: {str(e)}") 
     def toggle_blink(self): 
       try: 
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
       except Exception as e: 
           traceback.print_exc()  # Imprime o traceback completo do erro 
           QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao limpar os campos: {str(e)}")