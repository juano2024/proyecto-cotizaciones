import os
import openpyxl

class ExcelManager:
    def __init__(self, excel_path=""):
        self.excel_path = excel_path

    def guardar_cotizacion(self, datos):
        if not self.excel_path or not os.path.exists(self.excel_path):
            raise FileNotFoundError("No se ha seleccionado un archivo Excel de destino válido.")

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        ws.append(datos)
        wb.save(self.excel_path)
