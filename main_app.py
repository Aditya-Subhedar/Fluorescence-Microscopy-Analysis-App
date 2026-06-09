import tkinter as tk
from tkinter import ttk
import os
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Import separated tabs
from tab1_preprocessing import PreProcessingTab
from tab2_quantification import QuantificationTab
from tab3_representation import MaskMergerTab
from tab4_panel_creation import PanelCreationTab 
from tab5_golgi_cox_analysis import GolgiTab
from tab6_oft_tracking import OFTTrackingTab

class CytoQuantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CytoQuant Version 20")
        self.root.geometry("1920x1080")

        # Load Icon
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        # Create Notebook (Tab manager)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Tab 1: Pre-Processing
        self.tab1 = PreProcessingTab(self.notebook, self)
        self.notebook.add(self.tab1, text="1. Pre-Processing (3D)   ")

        # Tab 2: Quantification
        self.tab2 = QuantificationTab(self.notebook)
        self.notebook.add(self.tab2, text="2. Quantification (2D)   ")

        # Tab 3: Representation
        self.tab3 = MaskMergerTab(self.notebook, self)
        self.notebook.add(self.tab3, text="3. Representation    ")

        # Tab 4: Plate Creation (Placeholder)
        self.tab4 = PanelCreationTab(self.notebook, self)
        self.notebook.add(self.tab4, text="4. Panel Creation    ")

        # Tab 5: Golgi-Cox Analysis
        self.tab5 = GolgiTab(self.notebook)
        self.notebook.add(self.tab5, text="5. Golgi-Cox Analysis    ")

        # Tab 6: OFT Tracking
        self.tab6 = OFTTrackingTab(self.notebook)
        self.notebook.add(self.tab6, text="6. OFT Tracking  ")

if __name__ == "__main__":
    root = tk.Tk()
    app = CytoQuantApp(root)
    root.mainloop()