import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from interface import MainWindow

def main():
    app = QApplication(sys.argv)
    
    # Define o ícone do aplicativo
    icon_path = os.path.join(os.path.dirname(__file__), 'icons', 'icon.ico')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    window = MainWindow()
    window.resize(750, 320)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
