import sys
import sqlite3
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
                             QLabel, QPushButton, QHBoxLayout, QDialog, QFormLayout, QLineEdit, QSpinBox,
                             QDoubleSpinBox, QComboBox, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt

class AddEditDialog(QDialog):
    def __init__(self, parent=None, table_name=None, record=None, factory_list=None):
        super().__init__(parent)
        self.setWindowTitle("Добавление/Редактирование записи")
        self.table_name = table_name
        self.record = record
        self.factory_list = factory_list or []

        layout = QFormLayout(self)

        if table_name == "Brands":
            self.brand_name = QLineEdit()
            self.engine_capacity = QDoubleSpinBox()
            self.engine_capacity.setRange(0, 10)
            self.engine_capacity.setDecimals(1)
            self.max_speed = QSpinBox()
            self.max_speed.setRange(0, 500)
            self.release_year = QSpinBox()
            self.release_year.setRange(1900, 2100)
            self.factory = QComboBox()
            self.factory.addItems(self.factory_list)

            layout.addRow("Наименование марки:", self.brand_name)
            layout.addRow("Объем двигателя:", self.engine_capacity)
            layout.addRow("Максимальная скорость:", self.max_speed)
            layout.addRow("Год появления:", self.release_year)
            layout.addRow("Авт. завод:", self.factory)

            if record:
                self.brand_name.setText(record[0])
                self.engine_capacity.setValue(float(record[1]))
                self.max_speed.setValue(int(record[2]))
                self.release_year.setValue(int(record[3]))
                self.factory.setCurrentText(record[4])

        elif table_name == "Factories":
            self.factory = QLineEdit()
            self.country = QLineEdit()

            layout.addRow("Наименование завода:", self.factory_name)
            layout.addRow("Страна:", self.country)

            if record:
                self.factory.setText(record[0])
                self.country.setText(record[1])

        buttons = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Отмена")
        buttons.addWidget(self.ok_button)
        buttons.addWidget(self.cancel_button)
        layout.addRow(buttons)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Билет №24")
        self.setGeometry(100, 100, 800, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        layout.addWidget(QLabel("Автомобильные заводы"))
        search_layout_1 = QHBoxLayout()
        self.search_input_1 = QLineEdit()
        self.search_input_1.setPlaceholderText("Поиск по наименованию или стране")
        self.search_button_1 = QPushButton("Поиск")
        search_layout_1.addWidget(self.search_input_1)
        search_layout_1.addWidget(self.search_button_1)
        layout.addLayout(search_layout_1)

        self.table_widget_1 = QTableWidget()
        self.table_widget_1.setSelectionMode(QTableWidget.SingleSelection)
        self.table_widget_1.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table_widget_1)

        button_layout_1 = QHBoxLayout()
        self.add_button_1 = QPushButton("Добавить запись")
        self.edit_button_1 = QPushButton("Редактировать запись")
        self.delete_button_1 = QPushButton("Удалить запись")
        button_layout_1.addWidget(self.add_button_1)
        button_layout_1.addWidget(self.edit_button_1)
        button_layout_1.addWidget(self.delete_button_1)
        layout.addLayout(button_layout_1)

        layout.addWidget(QLabel("Таблица марок автомобилей"))
        search_layout_2 = QHBoxLayout()
        self.search_input_2 = QLineEdit()
        self.search_input_2.setPlaceholderText("Поиск по наименованию марки, году или заводу")
        self.search_button_2 = QPushButton("Поиск")
        search_layout_2.addWidget(self.search_input_2)
        search_layout_2.addWidget(self.search_button_2)
        layout.addLayout(search_layout_2)

        self.table_widget_2 = QTableWidget()
        self.table_widget_2.setSelectionMode(QTableWidget.SingleSelection)
        self.table_widget_2.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table_widget_2)

        button_layout_2 = QHBoxLayout()
        self.add_button_2 = QPushButton("Добавить запись")
        self.edit_button_2 = QPushButton("Редактировать запись")
        self.delete_button_2 = QPushButton("Удалить запись")
        self.export_button = QPushButton("Экспорт в Excel")
        button_layout_2.addWidget(self.add_button_2)
        button_layout_2.addWidget(self.edit_button_2)
        button_layout_2.addWidget(self.delete_button_2)
        button_layout_2.addWidget(self.export_button)
        layout.addLayout(button_layout_2)

        self.add_button_1.clicked.connect(self.add_record_factory)
        self.edit_button_1.clicked.connect(self.edit_record_factory)
        self.delete_button_1.clicked.connect(self.delete_record_factory)
        self.search_button_1.clicked.connect(self.search_factory)

        self.add_button_2.clicked.connect(self.add_record_brand)
        self.edit_button_2.clicked.connect(self.edit_record_brand)
        self.delete_button_2.clicked.connect(self.delete_record_brand)
        self.search_button_2.clicked.connect(self.search_brand)

        self.export_button.clicked.connect(self.export_to_excel)

        self.load_data()

    def load_data(self):
        try:
            conn = sqlite3.connect('Car_Factory.db')
            cursor = conn.cursor()
            self.load_table_1(cursor)
            self.load_table_2(cursor)
            conn.close()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при работе с базой данных: {e}")

    def load_table_1(self, cursor, search_query=None):
        if search_query:
            query = "SELECT * FROM Factories WHERE Factory_name LIKE ? OR Country LIKE ?"
            cursor.execute(query, (f'%{search_query}%', f'%{search_query}%'))
        else:
            cursor.execute("SELECT * FROM Factories")
        data = cursor.fetchall()
        original_column_names = [description[0] for description in cursor.description]
        custom_column_names = ["Наименование", "Страна"]

        self.table_widget_1.setRowCount(len(data))
        self.table_widget_1.setColumnCount(len(original_column_names))
        self.table_widget_1.setHorizontalHeaderLabels(custom_column_names)

        for row_idx, row_data in enumerate(data):
            for col_idx, col_data in enumerate(row_data):
                self.table_widget_1.setItem(row_idx, col_idx, QTableWidgetItem(str(col_data)))

        self.table_widget_1.resizeColumnsToContents()

    def load_table_2(self, cursor, search_query=None):
        if search_query:
            query = "SELECT * FROM Brands WHERE Brand_name LIKE ? OR Release_year LIKE ? OR Factory LIKE ?"
            cursor.execute(query, (f'%{search_query}%', f'%{search_query}%', f'%{search_query}%'))
        else:
            cursor.execute("SELECT * FROM Brands")
        data = cursor.fetchall()
        original_column_names = [description[0] for description in cursor.description]
        custom_column_names = ["Наименование марки", "Объем двигателя", "Максимальная скорость", "Год появления", "Авт. завод"]

        self.table_widget_2.setRowCount(len(data))
        self.table_widget_2.setColumnCount(len(original_column_names))
        self.table_widget_2.setHorizontalHeaderLabels(custom_column_names)

        for row_idx, row_data in enumerate(data):
            for col_idx, col_data in enumerate(row_data):
                self.table_widget_2.setItem(row_idx, col_idx, QTableWidgetItem(str(col_data)))

        self.table_widget_2.resizeColumnsToContents()

    def search_factory(self):
        search_query = self.search_input_1.text()
        try:
            conn = sqlite3.connect('Car_Factory.db')
            cursor = conn.cursor()
            self.load_table_1(cursor, search_query)
            conn.close()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при поиске: {e}")

    def search_brand(self):
        search_query = self.search_input_2.text()
        try:
            conn = sqlite3.connect('Car_Factory.db')
            cursor = conn.cursor()
            self.load_table_2(cursor, search_query)
            conn.close()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при поиске: {e}")

    def export_to_excel(self):
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", "Car_report.xlsx", "Excel Files (*.xlsx)")
            if not file_path:
                return

            conn = sqlite3.connect('Car_Factory.db')

            df_factory = pd.read_sql_query("SELECT * FROM Factories", conn)
            df_factory.columns = ["Наименование", "Страна"]

            df_brand = pd.read_sql_query("SELECT * FROM Brands", conn)
            df_brand.columns = ["Наименование марки", "Объем двигателя", "Максимальная скорость", "Год появления", "Авт. завод"]

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_factory.to_excel(writer, sheet_name="Factories", index=False)
                df_brand.to_excel(writer, sheet_name="Brands", index=False)

            conn.close()
            QMessageBox.information(self, "Успех", f"Данные успешно экспортированы в {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте в Excel: {e}")

    def get_factory_list(self):
        try:
            conn = sqlite3.connect('Car_Factory.db')
            cursor = conn.cursor()
            cursor.execute("SELECT Factory_name FROM Factories")
            factory_list = [row[0] for row in cursor.fetchall()]
            conn.close()
            return factory_list
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при получении списка заводов: {e}")
            return []

    def add_record_factory(self):
        dialog = AddEditDialog(self, "Factories")
        if dialog.exec_():
            Factory_name = dialog.Factory_name.text()
            Country = dialog.country.text()

            if not name_factory or not country:
                QMessageBox.warning(self, "Предупреждение", "Все поля должны быть заполнены!")
                return

            try:
                conn = sqlite3.connect('Car_Factory.db')
                cursor = conn.cursor()
                cursor.execute("INSERT INTO Factories (Factory_name, Country) VALUES (?, ?)", (Factory_name, Country))
                conn.commit()
                conn.close()
                self.load_data()
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось добавить запись: {e}")

    def edit_record_factory(self):
        selected_row = self.table_widget_1.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Предупреждение", "Выберите запись для редактирования!")
            return

        record = [self.table_widget_1.item(selected_row, col).text() for col in range(self.table_widget_1.columnCount())]
        dialog = AddEditDialog(self, "Factories", record)
        if dialog.exec_():
            new_Factory_name = dialog.Factory_name.text()
            country = dialog.country.text()

            if not new_Factory_name or not Country:
                QMessageBox.warning(self, "Предупреждение", "Все поля должны быть заполнены!")
                return

            if new_Factory_name == record[0]:
                try:
                    conn = sqlite3.connect('Car_Factory.db')
                    cursor = conn.cursor()
                    cursor.execute("UPDATE Factories SET Country = ? WHERE Factory_name = ?", (Country, record[0]))
                    conn.commit()
                    conn.close()
                    self.load_data()
                except sqlite3.Error as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось обновить запись: {e}")
            else:
                try:
                    conn = sqlite3.connect('Car_Factory.db')
                    cursor = conn.cursor()
                    cursor.execute("UPDATE Brands SET Factory = ? WHERE Factory = ?", (new_Factory_name, record[0]))
                    cursor.execute("DELETE FROM Factories WHERE Factory_name = ?", (record[0],))
                    cursor.execute("INSERT INTO Factories (Factory_name, Country) VALUES (?, ?)", (new_Factory_name, Country))
                    conn.commit()
                    conn.close()
                    self.load_data()
                except sqlite3.Error as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось обновить запись: {e}")

    def delete_record_factory(self):
        selected_row = self.table_widget_1.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Предупреждение", "Выберите запись для удаления!")
            return

        Factory_name = self.table_widget_1.item(selected_row, 0).text()
        try:
            conn = sqlite3.connect('Car_Factory.db')
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM Brands WHERE Factory = ?", (Factory_name,))
            count = cursor.fetchone()[0]
            if count > 0:
                QMessageBox.warning(self, "Предупреждение", "Нельзя удалить завод, так как он связан с марками!")
                conn.close()
                return

            reply = QMessageBox.question(self, "Подтверждение", f"Вы уверены, что хотите удалить запись '{Factory_name}'?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return

            cursor.execute("DELETE FROM Factories WHERE Factory_name = ?", (Factory_name,))
            conn.commit()
            conn.close()
            self.load_data()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось удалить запись: {e}")

    def add_record_brand(self):
        factory_list = self.get_factory_list()
        if not factory_list:
            QMessageBox.warning(self, "Предупреждение", "Сначала добавьте хотя бы один завод!")
            return

        dialog = AddEditDialog(self, "Brands", factory_list=factory_list)
        if dialog.exec_():
            Brand_name = dialog.Brand_name.text()
            Engine_capacity = dialog.Engine_capacity.value()
            Max_speed = dialog.Max_speed.value()
            Release_year = dialog.Release_year.value()
            Factory = dialog.Factory.currentText()

            if not Brand_name:
                QMessageBox.warning(self, "Предупреждение", "Наименование марки должно быть заполнено!")
                return

            try:
                conn = sqlite3.connect('Car_Factory.db')
                cursor = conn.cursor()
                cursor.execute("INSERT INTO Brands (Brand_name, Engine_capacity, Max_speed, Release_year, Factory) VALUES (?, ?, ?, ?, ?)",
                               (Brand_name, Engine_capacity, Max_speed, Release_year, Factory))
                conn.commit()
                conn.close()
                self.load_data()
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось добавить запись: {e}")

    def edit_record_brand(self):
        selected_row = self.table_widget_2.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Предупреждение", "Выберите запись для редактирования!")
            return

        factory_list = self.get_factory_list()
        if not factory_list:
            QMessageBox.warning(self, "Предупреждение", "Список заводов пуст!")
            return

        record = [self.table_widget_2.item(selected_row, col).text() for col in range(self.table_widget_2.columnCount())]
        dialog = AddEditDialog(self, "Brands", record, factory_list)
        if dialog.exec_():
            new_Brand_name = dialog.Brand_name.text()
            Engine_capacity = dialog.Engine_capacity.value()
            Max_speed = dialog.Max_speed.value()
            Release_year = dialog.Release_year.value()
            Factory = dialog.Factory.currentText()

            if not new_Brand_name:
                QMessageBox.warning(self, "Предупреждение", "Наименование марки должно быть заполнено!")
                return

            if new_Brand_name == record[0]:
                try:
                    conn = sqlite3.connect('Car_Factory.db')
                    cursor = conn.cursor()
                    cursor.execute("UPDATE Brands SET Engine_capacity = ?, Max_speed= ?, Release_yeare = ?, Factory = ? WHERE Brand_name = ?",
                                   (Engine_capacity, Max_speed, Release_year, Factory, record[0]))
                    conn.commit()
                    conn.close()
                    self.load_data()
                except sqlite3.Error as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось обновить запись: {e}")
            else:
                try:
                    conn = sqlite3.connect('Car_Factory.db')
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM Brands WHERE Brand_name = ?", (record[0],))
                    cursor.execute("INSERT INTO Brands (Brand_name, Engine_capacity, Max_speed, Release_year, Factory) VALUES (?, ?, ?, ?, ?)",
                                   (new_Brand_name, Engine_capacity, Max_speed, Release_year, Factory))
                    conn.commit()
                    conn.close()
                    self.load_data()
                except sqlite3.Error as e:
                    QMessageBox.critical(self, "Ошибка", f"Не удалось обновить запись: {e}")

    def delete_record_brand(self):
        selected_row = self.table_widget_2.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Предупреждение", "Выберите запись для удаления!")
            return

        Brand_name = self.table_widget_2.item(selected_row, 0).text()
        reply = QMessageBox.question(self, "Подтверждение", f"Вы уверены, что хотите удалить запись '{Brand_name}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return

        try:
            conn = sqlite3.connect('Car_Factory.db')
            cursor = conn.cursor()
            cursor.execute("DELETE FROM Brands WHERE Brand_name = ?", (Brand_name,))
            conn.commit()
            conn.close()
            self.load_data()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось удалить запись: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())