import os
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QFileDialog)
from gui.styles import MAIN_STYLE
from core.excel_manager import ExcelManager

class CotizacionesApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_manager = ExcelManager()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Registro de Cotizaciones")
        self.setFixedSize(400, 250)
        
        layout = QVBoxLayout()
        
        # Archivo Excel actual
        texto_archivo = os.path.basename(self.excel_manager.excel_path) if self.excel_manager.excel_path else "Ninguno seleccionado"
        self.lbl_archivo = QLabel(f"Archivo Excel: {texto_archivo}")
        self.lbl_archivo.setStyleSheet("color: gray; font-style: italic;")
        
        layout_archivo = QHBoxLayout()
        layout_archivo.addWidget(self.lbl_archivo)
        
        btn_cambiar_archivo = QPushButton("Seleccionar otro Excel...")
        btn_cambiar_archivo.clicked.connect(self.cambiar_archivo)
        layout_archivo.addWidget(btn_cambiar_archivo)
        
        layout.addLayout(layout_archivo)
        
        # Separador visual
        layout.addSpacing(10)
        
        # Nombre
        layout_nombre = QHBoxLayout()
        layout_nombre.addWidget(QLabel("Nombre del artículo:"))
        self.input_nombre = QLineEdit()
        layout_nombre.addWidget(self.input_nombre)
        layout.addLayout(layout_nombre)
        
        # Precio
        layout_precio = QHBoxLayout()
        layout_precio.addWidget(QLabel("Precio estimado ($):"))
        self.input_precio = QLineEdit()
        layout_precio.addWidget(self.input_precio)
        layout.addLayout(layout_precio)
        
        # Lugar
        layout_lugar = QHBoxLayout()
        layout_lugar.addWidget(QLabel("Lugar o Enlace:"))
        self.input_lugar = QLineEdit()
        layout_lugar.addWidget(self.input_lugar)
        layout.addLayout(layout_lugar)
        
        # Categoría
        layout_cat = QHBoxLayout()
        layout_cat.addWidget(QLabel("Categoría:"))
        self.combo_cat = QComboBox()
        self.combo_cat.addItems(["Supplier", "Sales force"])
        layout_cat.addWidget(self.combo_cat)
        layout.addLayout(layout_cat)
        
        layout.addSpacing(15)
        
        # Botón Guardar
        btn_guardar = QPushButton("Guardar Cotización")
        btn_guardar.clicked.connect(self.guardar_cotizacion)
        btn_guardar.setStyleSheet(MAIN_STYLE)
        layout.addWidget(btn_guardar)
        
        self.setLayout(layout)
        
    def cambiar_archivo(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Seleccionar archivo Excel', os.getcwd(), "Archivos Excel (*.xlsx *.xls)")
        if fname:
            self.excel_manager.excel_path = fname
            self.lbl_archivo.setText(f"Archivo Excel: {os.path.basename(self.excel_manager.excel_path)}")
            
    def guardar_cotizacion(self):
        if not self.excel_manager.excel_path or not os.path.exists(self.excel_manager.excel_path):
            QMessageBox.warning(self, "Archivo no seleccionado", "Por favor, selecciona un archivo Excel existente antes de guardar la cotización.")
            return

        nombre = self.input_nombre.text().strip()
        precio_str = self.input_precio.text().strip()
        lugar = self.input_lugar.text().strip()
        categoria = self.combo_cat.currentText()
        
        if not nombre or not precio_str or not lugar:
            QMessageBox.warning(self, "Campos Incompletos", "Por favor, completa todos los campos antes de guardar.")
            return
            
        try:
            precio = float(precio_str.replace(',', '.'))
        except ValueError:
            QMessageBox.warning(self, "Error de Formato", "El precio debe ser un número válido (ej. 1500.50).")
            return
            
        try:
            self.excel_manager.guardar_cotizacion(nombre, precio, lugar, categoria)
            QMessageBox.information(self, "¡Guardado Exitoso!", f"La cotización de '{nombre}' ha sido registrada en el Excel.")
            
            # Limpiar campos para la siguiente entrada
            self.input_nombre.clear()
            self.input_precio.clear()
            self.input_lugar.clear()
            
        except PermissionError:
            QMessageBox.critical(self, "Archivo Abierto", f"El archivo Excel ({os.path.basename(self.excel_manager.excel_path)}) está abierto.\nPor favor, ciérralo antes de guardar una nueva cotización.")
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar la información en el Excel:\n{str(e)}")
