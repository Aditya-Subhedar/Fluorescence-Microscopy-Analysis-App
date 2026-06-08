import tkinter as tk
from tkinter import ttk

class PlateCreationPlaceholderTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Center the placeholder message
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        lbl = ttk.Label(
            self, 
            text="Tab 4: Plate Creation\n\n(Module Currently Under Construction)", 
            font=("Helvetica", 16, "bold"),
            justify="center"
        )
        lbl.grid(row=0, column=0)