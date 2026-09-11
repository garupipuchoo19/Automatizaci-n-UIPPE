import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.utils.helpers import inicializar_directorios
from src.gui.app import AppUIPPE

def main():
    inicializar_directorios()
    app = AppUIPPE()
    app.mainloop()

if __name__ == "__main__":
    main()