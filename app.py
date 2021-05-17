from components import Ui_AccountEngine as Ui
from PyQt5.QtWidgets import *
import sys
from funciones import init

class App( QDialog ):
    def __init__( self ):
        super( App, self ).__init__()
        self.app = Ui()
        self.app.setupUi( self )
        self.app.comenzar.clicked.connect(self.comenzar)
    def comenzar(self):
        self.app.comenzar.setEnabled(False)
        self.app.advices.setText("Cargando...")
        filename = self.app.filename.text()
        try:
            init( filename )
            self.app.advices.setStyleSheet("color: white;")
            self.app.advices.setText("Hecho.")
        except PermissionError:
            self.app.advices.setStyleSheet("color: red;")
            self.app.advices.setText("Procura tener cerrado el documento excel antes de comenzar.")
        except Exception as e:
            if type(e).__name__ == "InvalidFileException":
                self.app.advices.setStyleSheet("color: gold;")
                filename = filename if filename != "" else "   " 
                self.app.advices.setText(f"No existe un archivo con el nombre '{ filename }'.")
            else:
                self.app.advices.setStyleSheet("color: red;")
                self.app.advices.setText(f"Se ha detectado un error desconocido. { type( e ).__name__ }")
        self.app.comenzar.setEnabled(True)
        


if __name__ == "__main__":
    app = QApplication( sys.argv )
    miApp = App()
    miApp.show()
    sys.exit( app.exec_() )