import tkinter as tk
from tkinter import ttk
import os
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Import your separated tabs!
from tab1_preprocessing import PreProcessingTab
from tab2_quantification import QuantificationTab
from tab3_golgi import GolgiTab  # <-- NEW: Import the third tab

class CytoQuantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CytoQuant Version 12")
        self.root.geometry("1920x1080") # Adjust default size as needed
        
        # Load the Icon safely using the resource_path function
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        # Create the Main Notebook (Tab manager)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Initialize and add Tab 1
        self.tab1 = PreProcessingTab(self.notebook, self)
        self.notebook.add(self.tab1, text="1. Pre-Processing (3D)       ")

        # Initialize and add Tab 2
        self.tab2 = QuantificationTab(self.notebook)
        self.notebook.add(self.tab2, text="2. Image Analysis (2D)       ")

        # Initialize and add Tab 3  <-- NEW
        self.tab3 = GolgiTab(self.notebook)
        self.notebook.add(self.tab3, text="3. Golgi Analysis (Beta)     ")

if __name__ == "__main__":
    root = tk.Tk()
    app = CytoQuantApp(root)
    root.mainloop()