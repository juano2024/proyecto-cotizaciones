import sys
from PyQt6.QtWidgets import QApplication
from gui.main_window import CotizacionesApp

def main():
    app = QApplication(sys.argv)
    
    # Estilo general de la aplicación
    app.setStyle("Fusion")
    
    ventana = CotizacionesApp()
    ventana.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
