from openpyxl import load_workbook, Workbook
from pyxl_class import excell
from pyxl_msgs import msgs
from PyQt5 import QtWidgets, uic
import sys
import os

def callClass(call):
    while True:
        try:
            loadclass = call
            return loadclass
        except PermissionError:
            input("\n FECHAR ARQUIVO ABERTO \n")

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__() # Call the inherited classes __init__ method
        uic.loadUi('qt\init.ui', self) # Load the .ui file
        #self.button = self.findChild(QtWidgets.QPushButton, 'optQ')  # Find the button
        #self.button.clicked.connect(self.endProgram)  # Remember to pass the definition/method, not the return value!
        self.show() # Show the GUI

    def endProgram(self):
        return

app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()