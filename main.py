import sys
from PyQt5.QtWidgets import QApplication
from gui import CraneOperationSystem

def main():
    app = QApplication(sys.argv)
    crane_system = CraneOperationSystem()
    crane_system.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()