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
        self.root.title("CytoQuant Version 24")
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

        # -----------------------------------------------------------------
        # CENTRAL ROUTING & AUTO-FOCUS FUNCTIONALITY
        # -----------------------------------------------------------------
        # 1. Bind global hotkeys to the root window object
        self.root.bind("<Left>", self.route_prev_image)
        self.root.bind("<Right>", self.route_next_image)
        self.root.bind("<Control-o>", self.route_open_file)
        self.root.bind("<Control-O>", self.route_open_file)

        # ---> NEW: CENTRAL UNDO / REDO BINDINGS <---
        self.root.bind("<Control-z>", self.route_undo)
        self.root.bind("<Control-Z>", self.route_undo)
        self.root.bind("<Control-y>", self.route_redo)
        self.root.bind("<Control-Y>", self.route_redo)

        # ---> NEW: CENTRAL UP / DOWN ARROW BINDINGS FOR Z-STACKS <---
        self.root.bind("<Up>", self.route_up_arrow)
        self.root.bind("<Down>", self.route_down_arrow)

        # ---> NEW: CENTRAL COPY / PASTE BINDINGS <---
        self.root.bind("<Control-c>", self.route_copy_box)
        self.root.bind("<Control-C>", self.route_copy_box)
        self.root.bind("<Control-v>", self.route_paste_box)
        self.root.bind("<Control-V>", self.route_paste_box)

        # ---> CENTRAL DELETE BINDING <---
        self.root.bind("<Delete>", self.route_delete_item)
        self.root.bind("<BackSpace>", self.route_delete_item)

        # 2. Watch for tab switches to shift focus dynamically
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_switched)

    def on_tab_switched(self, event):
        """Forces keyboard focus into the selected tab canvas for immediate shortcut usage."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 0 and hasattr(self.tab1, 'canvas'):
            self.tab1.canvas.focus_force()
        elif active_idx == 1 and hasattr(self.tab2, 'canvas'):
            self.tab2.canvas.focus_force()

    def route_open_file(self, event=None):
        """Routes Ctrl+O execution natively based on which tab viewport is visible."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 0:
            if hasattr(self.tab1, 'load_czi'): self.tab1.load_czi()
        elif active_idx == 1:
            if hasattr(self.tab2, 'load_files'): self.tab2.load_files()
        return "break" # Stops native TKinter control hooks from overriding dialog outputs

    def route_prev_image(self, event=None):
        """Routes Left Arrow cleanly depending on the active notebook panel layout context."""
        active_idx = self.notebook.index(self.notebook.select())
        
        if active_idx == 0:
            # Tab 1: Check if the button exists and read its state using [] syntax
            if hasattr(self.tab1, 'btn_prev_img') and self.tab1.btn_prev_img['state'] == tk.NORMAL:
                if hasattr(self.tab1, 'prev_image'): 
                    self.tab1.prev_image()
        elif active_idx == 1:
            # Tab 2: Check if the button exists and read its state safely
            if hasattr(self.tab2, 'btn_prev_img') and self.tab2.btn_prev_img['state'] == tk.NORMAL:
                if hasattr(self.tab2, 'prev_image'): 
                    self.tab2.prev_image()

    def route_next_image(self, event=None):
        """Routes Right Arrow cleanly depending on the active notebook panel layout context."""
        active_idx = self.notebook.index(self.notebook.select())
        
        if active_idx == 0:
            # Tab 1: Validate next image button exists and is enabled
            if hasattr(self.tab1, 'btn_next_img') and self.tab1.btn_next_img['state'] == tk.NORMAL:
                if hasattr(self.tab1, 'next_image'): 
                    self.tab1.next_image()
        elif active_idx == 1:
            # Tab 2: Validate next image button exists and is enabled
            if hasattr(self.tab2, 'btn_next_img') and self.tab2.btn_next_img['state'] == tk.NORMAL:
                if hasattr(self.tab2, 'next_image'): 
                    self.tab2.next_image()

    def route_undo(self, event=None):
        """Intercepts Ctrl+Z and safely triggers undo for active tabs."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1: # Tab 2
            if hasattr(self.tab2, 'undo_action'): self.tab2.undo_action()
        elif active_idx == 3: # Tab 4
            if hasattr(self.tab4, 'undo_action'): self.tab4.undo_action()
        return "break"

    def route_redo(self, event=None):
        """Intercepts Ctrl+Y and safely triggers redo for active tabs."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1: # Tab 2
            if hasattr(self.tab2, 'redo_action'): self.tab2.redo_action()
        elif active_idx == 3: # Tab 4
            if hasattr(self.tab4, 'redo_action'): self.tab4.redo_action()
        return "break"
    
    def route_up_arrow(self, event=None):
        """Intercepts Up Arrow and moves up 1 slice in Tab 1 Z-stack."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 0:  # Tab 1 is index 0
            if hasattr(self.tab1, 'change_z_slice'):
                self.tab1.change_z_slice(1)  # +1 moves up the stack
        return "break"  # Stops default widget scrolling overrides

    def route_down_arrow(self, event=None):
        """Intercepts Down Arrow and moves down 1 slice in Tab 1 Z-stack."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 0:  # Tab 1 is index 0
            if hasattr(self.tab1, 'change_z_slice'):
                self.tab1.change_z_slice(-1) # -1 moves down the stack
        return "break"  # Stops default widget scrolling overrides

    def route_copy_box(self, event=None):
        """Intercepts Ctrl+C and routes to Tab 4 copy function."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 3:  # Tab 4 (Panel Creation) is index 3
            if hasattr(self.tab4, 'copy_selected_box'):
                self.tab4.copy_selected_box()
        return "break"

    def route_paste_box(self, event=None):
        """Intercepts Ctrl+V and routes to Tab 4 paste function."""
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 3:  # Tab 4 (Panel Creation) is index 3
            if hasattr(self.tab4, 'paste_box'):
                self.tab4.paste_box()
        return "break"

    def route_delete_item(self, event=None):
        # Ensure we are checking for index 3 (the 4th tab)
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 3:
            # Check that tab4 exists and has the required method
            if hasattr(self.tab4, 'delete_selected_item'):
                self.tab4.delete_selected_item()
        return "break"

if __name__ == "__main__":
    root = tk.Tk()
    app = CytoQuantApp(root)
    root.mainloop()