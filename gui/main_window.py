import os
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QFormLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QFileDialog, QScrollArea)
from PyQt6.QtCore import Qt
from gui.styles import MAIN_STYLE
from core.excel_manager import ExcelManager

class CotizacionesApp(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_manager = ExcelManager()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Registro de Cotizaciones")
        self.resize(500, 700)
        
        main_layout = QVBoxLayout()
        
        # Archivo Excel actual
        texto_archivo = os.path.basename(self.excel_manager.excel_path) if self.excel_manager.excel_path else "Ninguno seleccionado"
        self.lbl_archivo = QLabel(f"Archivo Excel: {texto_archivo}")
        self.lbl_archivo.setStyleSheet("color: gray; font-style: italic;")
        
        layout_archivo = QHBoxLayout()
        layout_archivo.addWidget(self.lbl_archivo)
        
        btn_cambiar_archivo = QPushButton("Seleccionar otro Excel...")
        btn_cambiar_archivo.clicked.connect(self.cambiar_archivo)
        layout_archivo.addWidget(btn_cambiar_archivo)
        
        main_layout.addLayout(layout_archivo)
        main_layout.addSpacing(10)
        
        # Scroll Area para el formulario con los 18 campos
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        form_layout = QFormLayout(scroll_content)
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.inputs = {}
        self.field_names = [
            ("FECHA DEL PEDIDO", "fecha_pedido"),
            ("ORDEN DE COMPRA", "orden_compra"),
            ("COTIZACIÓN", "cotizacion"),
            ("ATENDIO", "atendio"),
            ("USUARIO", "usuario"),
            ("EMPRESA", "empresa"),
            ("CANTIDAD", "cantidad"),
            ("DESCRIPCIÓN DE LA MERCANCIA", "descripcion"),
            ("PRECIO UNITARIO (USD)", "precio_unitario"),
            ("PRECIO TOTAL", "precio_total"),
            ("TIEMPO ESTIMADO DE LLEGADA", "tiempo_llegada"),
            ("FECHA DE LLEGADA DEL PEDIDO", "fecha_llegada"),
            ("EMPRESA DE TRANSPORTE", "transporte"),
            ("No. DE GUIA", "guia"),
            ("No. DE UPS", "ups"),
            ("PREALERTADO", "prealertado"),
            ("INVOICE", "invoice"),
            ("POSICIÓN ARANCELARIA", "posicion_arancelaria")
        ]
        
        for label, name in self.field_names:
            widget = QLineEdit()
            self.inputs[name] = widget
            form_layout.addRow(QLabel(f"{label}:"), widget)
            
        # Conectar campos de cantidad y precio para autocalcular el total
        self.inputs["cantidad"].textChanged.connect(self.calcular_total)
        self.inputs["precio_unitario"].textChanged.connect(self.calcular_total)

        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        main_layout.addSpacing(15)
        
        # Botón Guardar
        btn_guardar = QPushButton("Guardar Cotización")
        btn_guardar.clicked.connect(self.guardar_cotizacion)
        btn_guardar.setStyleSheet(MAIN_STYLE)
        main_layout.addWidget(btn_guardar)
        
        self.setLayout(main_layout)
        
    def calcular_total(self):
        try:
            cantidad_str = self.inputs["cantidad"].text().replace(',', '.')
            precio_str = self.inputs["precio_unitario"].text().replace(',', '.')
            if cantidad_str and precio_str:
                cantidad = float(cantidad_str)
                precio = float(precio_str)
                total = cantidad * precio
                self.inputs["precio_total"].setText(f"{total:.2f}")
        except ValueError:
            pass # Si hay texto no numerico, no calculamos

    def cambiar_archivo(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Seleccionar archivo Excel', os.getcwd(), "Archivos Excel (*.xlsx *.xls)")
        if fname:
            self.excel_manager.excel_path = fname
            self.lbl_archivo.setText(f"Archivo Excel: {os.path.basename(self.excel_manager.excel_path)}")
            
    def guardar_cotizacion(self):
        if not self.excel_manager.excel_path or not os.path.exists(self.excel_manager.excel_path):
            QMessageBox.warning(self, "Archivo no seleccionado", "Por favor, selecciona un archivo Excel existente antes de guardar el registro.")
            return

        # Recopilar todos los datos en el orden de las columnas del Excel
        datos_guardar = []
        for label, name in self.field_names:
            widget = self.inputs[name]
            texto = widget.text().strip()
            datos_guardar.append(texto)
            
        # Comprobar que al menos haya agregado algun dato importante (e.g., descripcion o empresa)
        if not any(datos_guardar):
            QMessageBox.warning(self, "Campos Vacíos", "Por favor, completa al menos un campo antes de guardar.")
            return
            
        try:
            self.excel_manager.guardar_cotizacion(datos_guardar)
            QMessageBox.information(self, "¡Guardado Exitoso!", "Los datos han sido registrados en el Excel correctamente.")
            
            # Limpiar campos para la siguiente entrada
            for _, name in self.field_names:
                self.inputs[name].clear()
            
        except PermissionError:
            QMessageBox.critical(self, "Archivo Abierto", f"El archivo Excel ({os.path.basename(self.excel_manager.excel_path)}) está abierto.\nPor favor, ciérralo antes de guardar.")
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar la información en el Excel:\n{str(e)}")
