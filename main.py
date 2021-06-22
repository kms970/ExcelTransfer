import EventGUI
from PySide2 import QtWidgets

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    myWindow = EventGUI.MainView()
    app.exec_()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
