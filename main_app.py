import ctypes
try:
    # Tells Windows to make the Tkinter app DPI-aware for crisp fonts
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    
    # Force Windows Taskbar to treat KytoQuant as an independent application
    myappid = 'KytoQuant.workspace.v1.2.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass
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
from tab6_graph_creation import GraphCreationTab
from tab7_about_us import AboutUsTab

class SplashScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.overrideredirect(True) # Remove window borders
        
        # INCREASED WINDOW SIZE: 600x400 to give the GIF room to breathe
        width, height = 600, 400
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.configure(bg="#1e272e")

        # GIF Animation handling
        self.gif_path = resource_path("loading.gif")
        self.frames = []
        if os.path.exists(self.gif_path):
            try:
                idx = 0
                while True:
                    self.frames.append(tk.PhotoImage(file=self.gif_path, format=f"gif -index {idx}"))
                    idx += 1
            except Exception:
                pass # Reached the last frame of the GIF
                
        self.current_frame = 0
        
        # CONDITIONAL UI: If GIF exists, show GIF. If no GIF, show default title.
        if self.frames:
            self.gif_label = tk.Label(self, image=self.frames[0], bg="#1e272e")
            self.gif_label.pack(expand=True)
            self.animate_gif()
        else:
            self.lbl_title = tk.Label(self, text="KytoQuant", font=("Arial", 36, "bold"), bg="#1e272e", fg="#ffffff")
            self.lbl_title.pack(expand=True)

        # Always show the small loading text at the bottom
        self.lbl_loading = tk.Label(self, text="Loading workspace modules...", font=("Arial", 10), bg="#1e272e", fg="#dcdde1")
        self.lbl_loading.pack(side="bottom", pady=20)
            
    def animate_gif(self):
        if self.frames:
            self.current_frame = (self.current_frame + 1) % len(self.frames)
            self.gif_label.configure(image=self.frames[self.current_frame])
            self.after(50, self.animate_gif) # 50ms per frame


class KytoQuantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KytoQuant v1.1.0")
        self.root.state('zoomed')

        # Apply icon globally to root and all future top levels
        icon_path = resource_path("app_logo.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(default=icon_path)
            except Exception:
                pass

        # Style configurations for a clean, cohesive layout
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # 1. Set the background behind the tabs to match the workspace
        self.style.configure("TNotebook", background="#f8f9fa", borderwidth=0)
        
        # 2. Style the inactive tabs (Slimmer profile)
        self.style.configure("TNotebook.Tab", 
                             background="#e9ecef",  
                             foreground="#495057",
                             padding=[12, 2],       
                             font=("Arial", 9),     
                             borderwidth=0)
                             
        # 3. Style the active/selected tab (Crisp white, bold text)
        self.style.map("TNotebook.Tab",
                       background=[("selected", "#ffffff")],  
                       foreground=[("selected", "#1e272e")],  
                       font=[("selected", ("Arial", 9, "bold"))],
                       padding=[("selected", [12, 2])])
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # -----------------------------------------------------------------
        # DASHBOARD LANDING HOME PAGE INITIALIZATION (INDEX 0)
        # -----------------------------------------------------------------
        self.tab_home = tk.Frame(self.notebook, bg="#f8f9fa")
        
        # Change the empty string to a Unicode home icon and add a little padding
        self.notebook.add(self.tab_home, text="🏠")
        
        # Instantiate your analytical tabs (Indices 1 to 6)
        self.tab1 = PreProcessingTab(self.notebook, self)
        self.notebook.add(self.tab1, text="1. Pre-Processing (3D)   ")

        self.tab2 = QuantificationTab(self.notebook)
        self.notebook.add(self.tab2, text="2. Quantification (2D)   ")

        self.tab3 = MaskMergerTab(self.notebook, self)
        self.notebook.add(self.tab3, text="3. Annotation    ")

        self.tab4 = PanelCreationTab(self.notebook, self)
        self.notebook.add(self.tab4, text="4. Panel Creation    ")

        self.tab5 = GolgiTab(self.notebook)
        self.notebook.add(self.tab5, text="5. Golgi-Cox Analysis    ")

        self.tab6 = GraphCreationTab(self.notebook)
        self.notebook.add(self.tab6, text="6. Graph Generation  ")

        self.tab7 = AboutUsTab(self.notebook)
        self.notebook.add(self.tab7, text="ℹ️ About Us ")

        # Build the visual components of the grid menu on our home page
        self.build_welcome_dashboard()

        # -----------------------------------------------------------------
        # CENTRAL ROUTING & AUTO-FOCUS FUNCTIONALITY
        # -----------------------------------------------------------------
        self.root.bind("<Left>", self.route_prev_image)
        self.root.bind("<Right>", self.route_next_image)
        self.root.bind("<Control-o>", self.route_open_file)
        self.root.bind("<Control-O>", self.route_open_file)
        self.root.bind("<Control-z>", self.route_undo)
        self.root.bind("<Control-Z>", self.route_undo)
        self.root.bind("<Control-y>", self.route_redo)
        self.root.bind("<Control-Y>", self.route_redo)
        self.root.bind("<Up>", self.route_up_arrow)
        self.root.bind("<Down>", self.route_down_arrow)
        self.root.bind("<Control-c>", self.route_copy_box)
        self.root.bind("<Control-C>", self.route_copy_box)
        self.root.bind("<Control-v>", self.route_paste_box)
        self.root.bind("<Control-V>", self.route_paste_box)
        self.root.bind("<Delete>", self.route_delete_item)
        self.root.bind("<BackSpace>", self.route_delete_item)

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_switched)

    def build_welcome_dashboard(self):
        """Creates a grid of modular module launch tiles with smooth entrance animation math."""
        # Top Header Banner
        header_frame = tk.Frame(self.tab_home, bg="#1e272e", pady=25)
        header_frame.pack(fill=tk.X)
        
        lbl_welcome = tk.Label(header_frame, text="KytoQuant Analysis Workspace", font=("Arial", 26, "bold"), fg="#ffffff", bg="#1e272e")
        lbl_welcome.pack()
        
        lbl_sub = tk.Label(header_frame, text="Select an automated pipeline to begin image processing and extraction workflows", font=("Arial", 11), fg="#dcdde1", bg="#1e272e")
        lbl_sub.pack(pady=(5,0))

        # Grid container frame for tile layouts
        grid_container = tk.Frame(self.tab_home, bg="#f8f9fa")
        grid_container.pack(expand=True, fill=tk.BOTH, padx=40, pady=20)
        
        # Configure a 3x2 uniform grid layout profile
        for col in range(3): 
            grid_container.columnconfigure(col, weight=1, uniform="equal")
        for row in range(2): 
            grid_container.rowconfigure(row, weight=1, uniform="equal")

        # Configuration database definitions tailored for researchers
        self.modules_config = [
            {
                "target_idx": 1, 
                "title": "1. Pre-Processing (3D)", 
                "desc": "Load raw microscopy formats (OME-TIFF, LIF, CZI, ND2). Collapse Z-stacks into single multiple intensity projection methods, apply multi-channel pseudo-coloring for specific stains (e.g., DAPI, GFAP), adjust brightness and contrast of channels, and navigate high-resolution images smoothly.", 
                "color": "#0066cc", "icon": "🔬", "grid": (0, 0)
            },
            {
                "target_idx": 2, 
                "title": "2. Quantification (2D)", 
                "desc": "Automate cell counting and fluorescence intensity measurements of neuronal tissue sectins for IHC. Filter specific cells or fibers by hue, fluroscence, size and morphology, measure cell count and fluoroscent area, and export images, contours, and data directly to CSV.", 
                "color": "#2e7d32", "icon": "📊", "grid": (0, 1)
            },
            {
                "target_idx": 3, 
                "title": "3. Figure Layer Merger", 
                "desc": "Overlay quantitative boundary masks exactly onto your original grayscale or multi-channel microscopy images. Add customizable text annotations to generate clear, transparent, or solid-background visuals.", 
                "color": "#d35400", "icon": "🖋️", "grid": (0, 2)
            },
            {
                "target_idx": 4, 
                "title": "4. Publication Panel Creator", 
                "desc": "Design multi-image panels optimized for scientific manuscripts. Automatically handles image layout and labeling, assign sequential subpanel indexing (A, B, C), and export high-resolution, publication-ready figures.", 
                "color": "#6d214f", "icon": "📋", "grid": (1, 0)
            },
            {
                "target_idx": 5, 
                "title": "5. Golgi Morphological Profiling", 
                "desc": "Perform specialized Sholl analysis for neuronal morphology. Map dendritic networks, measure branching complexity, and quantify active spine densities across concentric distances from the soma.", 
                "color": "#c62828", "icon": "🧠", "grid": (1, 1)
            },
            {
                "target_idx": 6, 
                "title": "6. Graph Generation", 
                "desc": "Convert quantitative fluorescent and morphometric analysis sheets into publication-ready figures. Customize layouts, map statistical significance comparisons, compile error bars, and export high-DPI vector graphics.", 
                "color": "#00838f", "icon": "📈", "grid": (1, 2)
            }
        ]

        # --- ANIMATED ENTRANCE CASCADE ---
        # Stagger the tiles by 60ms each so they waterfall upwards
        for index, item in enumerate(self.modules_config):
            self.root.after(index * 60, lambda conf=item: self.render_dashboard_tile(grid_container, conf))

    def render_dashboard_tile(self, parent, conf):
        """Assembles an interactive module button card with a smooth slide-up animation."""
        row_idx, col_idx = conf["grid"]
        
        # 1. Static invisible container cell to hold the layout structure
        cell_frame = tk.Frame(parent, bg="#f8f9fa")
        cell_frame.grid(row=row_idx, column=col_idx, padx=15, pady=15, sticky="nsew")

        # 2. Outer card boundary container
        tile_card = tk.Frame(cell_frame, bd=1, relief=tk.SOLID, bg="#ffffff", cursor="hand2")
        
        # SLIDE ANIMATION: Start the tile 400 pixels below its intended location
        start_y_pixels = 400
        tile_card.place(relx=0, rely=0, relwidth=1.0, relheight=1.0, y=start_y_pixels)

        # Decorative side accent color bar
        accent_bar = tk.Frame(tile_card, bg=conf["color"], height=4)
        accent_bar.pack(fill=tk.X, side=tk.TOP)

        # Padding body container
        body = tk.Frame(tile_card, bg="#ffffff", padx=20, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        # Icon Label
        lbl_icon = tk.Label(body, text=conf["icon"], font=("Arial", 28), bg="#ffffff", fg=conf["color"])
        lbl_icon.pack(anchor="nw")

        # Title Label
        lbl_title = tk.Label(body, text=conf["title"], font=("Arial", 12, "bold"), bg="#ffffff", fg="#2c3e50")
        lbl_title.pack(anchor="nw", pady=(10, 5))

        # Description text block
        lbl_desc = tk.Label(body, text=conf["desc"], font=("Arial", 10), bg="#ffffff", fg="#7f8c8d", justify=tk.LEFT)
        lbl_desc.pack(anchor="nw", fill=tk.BOTH, expand=True)

        def update_wraplength(event, label=lbl_desc):
            safe_width = max(200, event.width - 45) 
            label.config(wraplength=safe_width)
            
        body.bind("<Configure>", update_wraplength)

        for widget in (tile_card, body, lbl_icon, lbl_title, lbl_desc):
            widget.bind("<Button-1>", lambda event, idx=conf["target_idx"]: self.switch_to_pipeline_tab(idx))

        # Hover states
        tile_card.bind("<Enter>", lambda e: tile_card.config(relief=tk.RAISED, bd=2))
        tile_card.bind("<Leave>", lambda e: tile_card.config(relief=tk.SOLID, bd=1))
        body.bind("<Enter>", lambda e: tile_card.config(relief=tk.RAISED, bd=2))
        body.bind("<Leave>", lambda e: tile_card.config(relief=tk.SOLID, bd=1))

        # 3. Trigger the slide-up animation loop
        self.animate_slide_up(tile_card, start_y_pixels)

    def animate_slide_up(self, widget, current_y):
        """Creates a smooth, hardware-like ease-out deceleration effect."""
        target_y = 0
        
        # Ease-out formula: move 20% of the remaining distance per frame
        new_y = current_y + (target_y - current_y) * 0.20
        
        if abs(new_y) < 0.5:
            # Snap to final position and finish
            widget.place(relx=0, rely=0, relwidth=1.0, relheight=1.0, y=0)
        else:
            # Update position and schedule next frame (~60 FPS)
            widget.place(relx=0, rely=0, relwidth=1.0, relheight=1.0, y=new_y)
            widget.after(16, self.animate_slide_up, widget, new_y)

    def switch_to_pipeline_tab(self, index_id):
        """Transitions notebook selections seamlessly into active data loops."""
        self.notebook.select(index_id)

    def on_tab_switched(self, event):
        """Forces hardware focus into selected tab configurations safely."""
        active_idx = self.notebook.index(self.notebook.select())
        
        tab_mapping = {
            1: self.tab1,
            2: self.tab2,
            3: self.tab3,
            4: self.tab4,
            5: self.tab5,
            6: self.tab6,
            7: self.tab7
        }
        
        target_tab = tab_mapping.get(active_idx)
        if target_tab and hasattr(target_tab, 'canvas'):
            target_tab.canvas.focus_force()

    def route_open_file(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1:
            if hasattr(self.tab1, 'load_czi'): self.tab1.load_czi()
        elif active_idx == 2:
            if hasattr(self.tab2, 'load_files'): self.tab2.load_files()
        return "break"

    def route_prev_image(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1:
            if hasattr(self.tab1, 'btn_prev_img') and self.tab1.btn_prev_img['state'] == tk.NORMAL:
                if hasattr(self.tab1, 'prev_image'): self.tab1.prev_image()
        elif active_idx == 2:
            if hasattr(self.tab2, 'btn_prev_img') and self.tab2.btn_prev_img['state'] == tk.NORMAL:
                if hasattr(self.tab2, 'prev_image'): self.tab2.prev_image()

    def route_next_image(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1:
            if hasattr(self.tab1, 'btn_next_img') and self.tab1.btn_next_img['state'] == tk.NORMAL:
                if hasattr(self.tab1, 'next_image'): self.tab1.next_image()
        elif active_idx == 2:
            if hasattr(self.tab2, 'btn_next_img') and self.tab2.btn_next_img['state'] == tk.NORMAL:
                if hasattr(self.tab2, 'next_image'): self.tab2.next_image()

    def route_undo(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 2:
            if hasattr(self.tab2, 'undo_action'): self.tab2.undo_action()
        elif active_idx == 4:
            if hasattr(self.tab4, 'undo_action'): self.tab4.undo_action()
        return "break"

    def route_redo(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 2:
            if hasattr(self.tab2, 'redo_action'): self.tab2.redo_action()
        elif active_idx == 4:
            if hasattr(self.tab4, 'redo_action'): self.tab4.redo_action()
        return "break"

    def route_up_arrow(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1:
            if hasattr(self.tab1, 'change_z_slice'): self.tab1.change_z_slice(1)
        return "break"

    def route_down_arrow(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 1:
            if hasattr(self.tab1, 'change_z_slice'): self.tab1.change_z_slice(-1)
        return "break"

    def route_copy_box(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 4:
            if hasattr(self.tab4, 'copy_selected_box'): self.tab4.copy_selected_box()
        return "break"

    def route_paste_box(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 4:
            if hasattr(self.tab4, 'paste_box'): self.tab4.paste_box()
        return "break"

    def route_delete_item(self, event=None):
        active_idx = self.notebook.index(self.notebook.select())
        if active_idx == 4:
            if hasattr(self.tab4, 'delete_selected_item'): self.tab4.delete_selected_item()
        return "break"

if __name__ == "__main__":
    root = tk.Tk()
    
    # SET ICON GLOBALLY BEFORE CREATING ANY WINDOWS
    icon_path = resource_path("app_logo.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(default=icon_path)
        except Exception:
            pass

    # Hide the main window while the splash screen is active
    root.withdraw()
    splash = SplashScreen(root)
    
    def launch_main_app():
        app = KytoQuantApp(root)
        splash.destroy()
        root.deiconify() 
        
    root.after(25, launch_main_app)
    root.mainloop()