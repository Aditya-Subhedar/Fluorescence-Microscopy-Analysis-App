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

class NeuroQuantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NeuroQuant V3: Complete Workflow")
        self.root.geometry("1400x900") # Adjust default size as needed
        
        # Load the Icon safely using the resource_path function
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        # Create the Main Notebook (Tab manager)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Initialize and add Tab 1
        self.tab1 = PreProcessingTab(self.notebook, main_app=self)
        self.notebook.add(self.tab1, text="1. Pre-Processing (CZI Layers/Channels)")

        # Initialize and add Tab 2
        self.tab2 = QuantificationTab(self.notebook)
        self.notebook.add(self.tab2, text="2. NeuroQuant Analysis")

if __name__ == "__main__":
    root = tk.Tk()
    app = NeuroQuantApp(root)
    root.mainloop()