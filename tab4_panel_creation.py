import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import copy

class PanelCreationTab(ttk.Frame):
    def __init__(self, parent, main_app=None):
        super().__init__(parent)
        self.main_app = main_app
        
        # State Data
        self.cell_images = {}      # key: (row, col), value: absolute filepath
        self.cell_rotations = {}   # key: (row, col), value: degrees (0, 90, 180, 270)
        self.grid_dims = (2, 2)
        self.cell_aspect_ratio = 1.0  
        # --- Bounding Box & Viewport State ---
        self.panel_boxes = []         # List of dicts: {'x1', 'y1', 'x2', 'y2', 'color', 'scalable'}
        self.clipboard_box = None     # NEW: Stores copied box
        
        # Previews
        self.composite_pil = None  
        self.tk_img = None        
        self.cell_thumbnails = {} 

        # Annotation / Figure settings
        self.sublabel_color_rgb = (255, 255, 255)  
        self.title_color_rgb = (0, 0, 0)          
        
        # Undo/Redo (for selections)
        self.undo_stack = [{'images': {}, 'rotations': {}}]
        self.redo_stack = []

        # --- Bounding Box & Viewport State ---
        self.panel_boxes = []         # List of dicts: {'x1', 'y1', 'x2', 'y2', 'color'}
        self.selected_box_idx = None
        self.interaction_state = "IDLE" # IDLE, DRAWING, DRAGGING, RESIZING
        self.resize_handle = None       # top_left, top, right, etc.
        
        # Viewport parameters
        self.view_scale = 1.0
        self.pan_x = 0
        self.pan_y = 0
        self.start_x = 0
        self.start_y = 0
        self.drag_start_box = None

        self.setup_ui()
        self.setup_keybindings()

    def setup_ui(self):
        root_frame = tk.Frame(self)
        root_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # --- LEFT PANEL: CONTROLLER ---
        left_panel = tk.Frame(root_frame, width=280)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # 1. Grid Configuration
        tk.Label(left_panel, text="Step 1: Define Grid", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(5, 5))
        grid_frame = tk.LabelFrame(left_panel, text="Grid Dims & Geometry", padx=5, pady=5)
        grid_frame.pack(fill=tk.X, pady=(0, 10))
        
        row_col_frame = tk.Frame(grid_frame)
        row_col_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(row_col_frame, text="Rows:").pack(side=tk.LEFT)
        self.spin_rows = tk.Spinbox(row_col_frame, from_=1, to=20, width=5)
        self.spin_rows.pack(side=tk.LEFT, padx=(2, 10))
        self.spin_rows.delete(0, tk.END); self.spin_rows.insert(0, "2")
        
        tk.Label(row_col_frame, text="Cols:").pack(side=tk.LEFT)
        self.spin_cols = tk.Spinbox(row_col_frame, from_=1, to=20, width=5)
        self.spin_cols.pack(side=tk.LEFT, padx=(2, 0))
        self.spin_cols.delete(0, tk.END); self.spin_cols.insert(0, "2")
        
        tk.Label(grid_frame, text="Image Orientation:").pack(anchor=tk.W, pady=(5, 0))
        self.combo_aspect = ttk.Combobox(grid_frame, values=["Original Ratio", "Square (1:1)", "Landscape (4:3)", "Portrait (3:4)"], state="readonly")
        self.combo_aspect.pack(fill=tk.X, pady=2)
        self.combo_aspect.set("Original Ratio")
        
        gap_frame = tk.Frame(grid_frame)
        gap_frame.pack(fill=tk.X, pady=(5, 2))
        tk.Label(gap_frame, text="Gap (pixels):").pack(side=tk.LEFT)
        self.spin_gap = tk.Spinbox(gap_frame, from_=0, to=500, increment=10, width=8)
        self.spin_gap.pack(side=tk.RIGHT)
        self.spin_gap.delete(0, tk.END); self.spin_gap.insert(0, "40")
        
        tk.Button(left_panel, text="🎬 Generate / Reset Grid", command=self.action_generate_grid, bg="#0288d1", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=(0, 15))

        # 2. Titles & Labels
        tk.Label(left_panel, text="Step 2: Add Titles & Labels", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(5, 5))
        titles_frame = tk.LabelFrame(left_panel, text="Figure Captions (Locked to Black)", padx=5, pady=5)
        titles_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(titles_frame, text="Column Titles (Top - comma sep):").pack(anchor=tk.W)
        self.entry_col_titles = tk.Entry(titles_frame)
        self.entry_col_titles.pack(fill=tk.X, pady=2)

        tk.Label(titles_frame, text="Row Titles (Left - comma sep):").pack(anchor=tk.W)
        self.entry_row_titles = tk.Entry(titles_frame)
        self.entry_row_titles.pack(fill=tk.X, pady=2)

        tk.Label(titles_frame, text="Row Title Orientation:").pack(anchor=tk.W)
        self.combo_row_orient = ttk.Combobox(titles_frame, values=["Horizontal", "Sideways (90° CCW)"], state="readonly")
        self.combo_row_orient.pack(fill=tk.X, pady=2)
        self.combo_row_orient.set("Sideways (90° CCW)")
        
        labels_config_frame = tk.Frame(titles_frame)
        labels_config_frame.pack(fill=tk.X, pady=(5, 0))
        tk.Label(labels_config_frame, text="Subpanel labels (A, B, C):").pack(side=tk.LEFT)
        self.combo_sublabels = ttk.Combobox(labels_config_frame, values=["Inside Top-Left", "Outside Top-Left", "None"], state="readonly", width=15)
        self.combo_sublabels.pack(side=tk.RIGHT)
        self.combo_sublabels.set("Inside Top-Left")

        # 3. Text & Font Settings
        font_frame = tk.LabelFrame(left_panel, text="🎨 Font Settings", padx=5, pady=5)
        font_frame.pack(fill=tk.X, pady=(0, 15))
        
        # New Font Family Selection
        font_family_frame = tk.Frame(font_frame)
        font_family_frame.pack(fill=tk.X, pady=2)
        tk.Label(font_family_frame, text="Font Family:").pack(side=tk.LEFT)
        self.combo_font = ttk.Combobox(font_family_frame, values=["Arial", "Times New Roman", "Courier New", "Verdana", "Tahoma", "Georgia"], state="readonly", width=12)
        self.combo_font.pack(side=tk.RIGHT)
        self.combo_font.set("Arial")

        color_row = tk.Frame(font_frame)
        color_row.pack(fill=tk.X, pady=2)
        tk.Label(color_row, text="Img Label Color:").pack(side=tk.LEFT)
        self.btn_font_color = tk.Button(color_row, text="Label Color", command=self.pick_sublabel_color, bg="#ffffff", fg="black", width=12)
        self.btn_font_color.pack(side=tk.RIGHT)
        
        font_sizes_frame = tk.Frame(font_frame)
        font_sizes_frame.pack(fill=tk.X, pady=2)
        
        grid_sizer = tk.Frame(font_sizes_frame)
        grid_sizer.pack(fill=tk.X)
        tk.Label(grid_sizer, text="Label Size (px):").grid(row=0, column=0, sticky=tk.W)
        self.spin_sublabel_size = tk.Spinbox(grid_sizer, from_=10, to=500, increment=10, width=5)
        self.spin_sublabel_size.grid(row=0, column=1)
        self.spin_sublabel_size.delete(0, tk.END); self.spin_sublabel_size.insert(0, "48")
        
        tk.Label(grid_sizer, text="Title Size (px):").grid(row=1, column=0, sticky=tk.W)
        self.spin_title_size = tk.Spinbox(grid_sizer, from_=10, to=500, increment=10, width=5)
        self.spin_title_size.grid(row=1, column=1)
        self.spin_title_size.delete(0, tk.END); self.spin_title_size.insert(0, "56")

        # Box Options
        box_frame = tk.Frame(left_panel)
        box_frame.pack(fill=tk.X, pady=5)
        tk.Button(box_frame, text="🗑 Delete Selected Box", command=self.delete_selected_box, bg="#d32f2f", fg="white").pack(fill=tk.X)
        
        # 4. Global Actions / Export
        tk.Button(left_panel, text="🔄 Build / Refresh Preview", command=self.action_build_composite, bg="#1976d2", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=(10, 5))
        
        export_frame = tk.Frame(left_panel)
        export_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
        
        undo_redo_row = tk.Frame(export_frame)
        undo_redo_row.pack(fill=tk.X, pady=2)
        tk.Button(undo_redo_row, text="↩️ Undo Sel.", command=self.undo, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        tk.Button(undo_redo_row, text="↪️ Redo Sel.", command=self.redo, width=10).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        tk.Button(
            export_frame, text="💾 Export Panel High-Res", 
            command=self.action_export_panel, bg="#2e7d32", fg="white", 
            font=("Arial", 11, "bold"), height=2
        ).pack(fill=tk.X, pady=5)
        
        # --- RIGHT PANEL: CANVAS PRESENTATION ---
        self.preview_frame = tk.Frame(root_frame, bg="#1a1a1a") 
        self.preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.preview_frame, bg="#1a1a1a", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Canvas Interactions
        self.canvas.bind("<Configure>", lambda e: self.render_canvas_preview())
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<ButtonPress-2>", self.start_pan)
        self.canvas.bind("<B2-Motion>", self.do_pan)
        
        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)

    def setup_keybindings(self):
        toplevel = self.winfo_toplevel()
        toplevel.bind("<Control-z>", self.undo)
        toplevel.bind("<Control-y>", self.redo)
        toplevel.bind("<Control-s>", self.action_export_panel)
        toplevel.bind("<Control-g>", self.action_generate_grid)
        toplevel.bind("<Delete>", lambda e: self.delete_selected_box())
        toplevel.bind("<BackSpace>", lambda e: self.delete_selected_box())

    # --- MATH & COORDINATE CONVERSION ---
    def screen_to_img(self, sx, sy):
        return (sx - self.pan_x) / self.view_scale, (sy - self.pan_y) / self.view_scale

    def img_to_screen(self, ix, iy):
        return ix * self.view_scale + self.pan_x, iy * self.view_scale + self.pan_y

    # --- INTERACTION CONTROLS (PAN, ZOOM, DRAW) ---
    def on_mouse_wheel(self, event):
        if not self.composite_pil: return
        zoom_factor = 1.1 if event.delta > 0 else 0.9
        
        sx, sy = event.x, event.y
        ix, iy = self.screen_to_img(sx, sy)
        
        self.view_scale *= zoom_factor
        
        self.pan_x = sx - ix * self.view_scale
        self.pan_y = sy - iy * self.view_scale
        
        self.render_canvas_preview()

    def start_pan(self, event):
        self.canvas.config(cursor="fleur")
        self.start_x = event.x
        self.start_y = event.y
        self.start_pan_x = self.pan_x
        self.start_pan_y = self.pan_y

    def do_pan(self, event):
        dx = event.x - self.start_x
        dy = event.y - self.start_y
        self.pan_x = self.start_pan_x + dx
        self.pan_y = self.start_pan_y + dy
        
        self.canvas.coords("bg", self.pan_x, self.pan_y)
        self.draw_interactive_elements()
        
    def get_handle_rects(self, box):
        x1, y1 = self.img_to_screen(box['x1'], box['y1'])
        x2, y2 = self.img_to_screen(box['x2'], box['y2'])
        hw = 5 
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        return {
            "top_left": (x1-hw, y1-hw, x1+hw, y1+hw),
            "top_right": (x2-hw, y1-hw, x2+hw, y1+hw),
            "bottom_left": (x1-hw, y2-hw, x1+hw, y2+hw),
            "bottom_right": (x2-hw, y2-hw, x2+hw, y2+hw),
            "top": (cx-hw, y1-hw, cx+hw, y1+hw),
            "bottom": (cx-hw, y2-hw, cx+hw, y2+hw),
            "left": (x1-hw, cy-hw, x1+hw, cy+hw),
            "right": (x2-hw, cy-hw, x2+hw, cy+hw)
        }

    def on_left_down(self, event):
        if not self.composite_pil: return
        self.start_x, self.start_y = self.screen_to_img(event.x, event.y)
        
        if self.selected_box_idx is not None:
            box = self.panel_boxes[self.selected_box_idx]
            
            # NEW: Only allow resizing if the box is permitted to scale
            if box.get('scalable', True):
                handles = self.get_handle_rects(box)
                for handle_name, (hx1, hy1, hx2, hy2) in handles.items():
                    if hx1 <= event.x <= hx2 and hy1 <= event.y <= hy2:
                        self.interaction_state = "RESIZING"
                        self.resize_handle = handle_name
                        self.drag_start_box = copy.deepcopy(box)
                        return

        for i, box in reversed(list(enumerate(self.panel_boxes))):
            x1, y1 = self.img_to_screen(box['x1'], box['y1'])
            x2, y2 = self.img_to_screen(box['x2'], box['y2'])
            min_x, max_x = min(x1, x2), max(x1, x2)
            min_y, max_y = min(y1, y2), max(y1, y2)
            
            if min_x <= event.x <= max_x and min_y <= event.y <= max_y:
                self.selected_box_idx = i
                self.interaction_state = "DRAGGING"
                self.drag_start_box = copy.deepcopy(box)
                self.draw_interactive_elements()
                return

        self.selected_box_idx = None
        self.interaction_state = "DRAWING"
        # NEW: Ensure fresh hand-drawn boxes are explicitly marked as scalable
        new_box = {'x1': self.start_x, 'y1': self.start_y, 'x2': self.start_x, 'y2': self.start_y, 'color': '#FFFF00', 'scalable': True}
        self.panel_boxes.append(new_box)
        self.selected_box_idx = len(self.panel_boxes) - 1
        self.draw_interactive_elements()

    def on_left_drag(self, event):
        if self.interaction_state == "IDLE" or self.selected_box_idx is None: return
        
        curr_img_x, curr_img_y = self.screen_to_img(event.x, event.y)
        box = self.panel_boxes[self.selected_box_idx]

        if self.interaction_state == "DRAWING":
            box['x2'] = curr_img_x
            box['y2'] = curr_img_y

        elif self.interaction_state == "DRAGGING":
            dx = curr_img_x - self.start_x
            dy = curr_img_y - self.start_y
            box['x1'] = self.drag_start_box['x1'] + dx
            box['y1'] = self.drag_start_box['y1'] + dy
            box['x2'] = self.drag_start_box['x2'] + dx
            box['y2'] = self.drag_start_box['y2'] + dy

        elif self.interaction_state == "RESIZING":
            if "left" in self.resize_handle: box['x1'] = curr_img_x
            if "right" in self.resize_handle: box['x2'] = curr_img_x
            if "top" in self.resize_handle: box['y1'] = curr_img_y
            if "bottom" in self.resize_handle: box['y2'] = curr_img_y

        self.draw_interactive_elements()

    def on_left_up(self, event):
        self.canvas.config(cursor="")
        if self.interaction_state == "DRAWING":
            box = self.panel_boxes[self.selected_box_idx]
            if abs(box['x2'] - box['x1']) < 5 or abs(box['y2'] - box['y1']) < 5:
                self.panel_boxes.pop(self.selected_box_idx)
                self.selected_box_idx = None
        
        if self.selected_box_idx is not None:
            box = self.panel_boxes[self.selected_box_idx]
            box['x1'], box['x2'] = min(box['x1'], box['x2']), max(box['x1'], box['x2'])
            box['y1'], box['y2'] = min(box['y1'], box['y2']), max(box['y1'], box['y2'])

        self.interaction_state = "IDLE"
        self.draw_interactive_elements()

    def delete_selected_box(self):
        if self.selected_box_idx is not None and self.selected_box_idx < len(self.panel_boxes):
            self.panel_boxes.pop(self.selected_box_idx)
            self.selected_box_idx = None
            self.draw_interactive_elements()

    def copy_selected_box(self):
        if self.selected_box_idx is not None and self.selected_box_idx < len(self.panel_boxes):
            self.clipboard_box = copy.deepcopy(self.panel_boxes[self.selected_box_idx])

    def paste_box(self):
        if self.clipboard_box is not None:
            new_box = copy.deepcopy(self.clipboard_box)
            
            # Offset the box slightly so it doesn't paste invisibly on top of the old one
            offset = 20 / self.view_scale 
            new_box['x1'] += offset
            new_box['y1'] += offset
            new_box['x2'] += offset
            new_box['y2'] += offset
            
            # Restrict scaling on the pasted box
            new_box['scalable'] = False
            
            self.panel_boxes.append(new_box)
            self.selected_box_idx = len(self.panel_boxes) - 1
            self.draw_interactive_elements()

    # --- GRID / SELECTION LOGIC ---
    def save_selection_state(self):
        state = {
            'images': copy.deepcopy(self.cell_images),
            'rotations': copy.deepcopy(self.cell_rotations)
        }
        self.undo_stack.append(state)
        self.redo_stack.clear()
        if len(self.undo_stack) > 50: self.undo_stack.pop(0)

    def undo(self, event=None):
        if len(self.undo_stack) > 1:
            self.redo_stack.append(self.undo_stack.pop())
            top_state = self.undo_stack[-1]
            self.cell_images = copy.deepcopy(top_state['images'])
            self.cell_rotations = copy.deepcopy(top_state['rotations'])
            self.cell_aspect_ratio = self.get_cell_aspect()
            self.generate_canvas_placeholders()
            if self.composite_pil is not None: self.action_build_composite()
        else:
            messagebox.showinfo("Info", "Nothing to undo in selections.")

    def redo(self, event=None):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.cell_images = copy.deepcopy(state['images'])
            self.cell_rotations = copy.deepcopy(state['rotations'])
            self.undo_stack.append(state)
            self.cell_aspect_ratio = self.get_cell_aspect()
            self.generate_canvas_placeholders()
            if self.composite_pil is not None: self.action_build_composite()
        else:
            messagebox.showinfo("Info", "Nothing to redo in selections.")

    def get_cell_aspect(self):
        choice = self.combo_aspect.get()
        if "Square" in choice: return 1.0
        if "Landscape" in choice: return 4.0 / 3.0
        if "Portrait" in choice: return 3.0 / 4.0
        if "Original" in choice:
            for (r, c), path in self.cell_images.items():
                if path and os.path.exists(path):
                    try:
                        with Image.open(path) as img:
                            rot = self.cell_rotations.get((r, c), 0)
                            if rot in [90, 270]: return img.size[1] / img.size[0]
                            return img.size[0] / img.size[1]
                    except: pass
            return 1.0
        return 1.0

    def pick_sublabel_color(self):
        color_choice = colorchooser.askcolor(initialcolor="#ffffff", title="Select Subpanel Label Color")
        if color_choice and color_choice[0] is not None:
            self.sublabel_color_rgb = tuple(int(c) for c in color_choice[0])
            hex_color = color_choice[1]
            btn_fg = "black" if (sum(self.sublabel_color_rgb)/3) > 128 else "white"
            self.btn_font_color.config(text=f"Label: {hex_color}", bg=hex_color, fg=btn_fg)

    def rotate_cell(self, row, col, angle):
        if (row, col) in self.cell_images:
            self.save_selection_state()
            curr_rot = self.cell_rotations.get((row, col), 0)
            self.cell_rotations[(row, col)] = (curr_rot + angle) % 360
            self.cell_aspect_ratio = self.get_cell_aspect()
            self.generate_canvas_placeholders()
            if self.composite_pil is not None: self.action_build_composite()

    def action_generate_grid(self, event=None):
        try:
            r = int(self.spin_rows.get())
            c = int(self.spin_cols.get())
            self.grid_dims = (r, c)
            self.cell_aspect_ratio = self.get_cell_aspect()
            
            self.composite_pil = None 
            self.tk_img = None
            self.panel_boxes.clear() 
            
            if list(self.cell_images.keys()) and (r != len(set(k[0] for k in self.cell_images.keys())) or c != len(set(k[1] for k in self.cell_images.keys()))):
                if messagebox.askyesno("Clear Images?", "Grid dimensions changed. Clear currently selected images?"):
                    self.cell_images.clear()
                    self.cell_rotations.clear()
            
            self.generate_canvas_placeholders()
        except ValueError:
            messagebox.showerror("Error", "Invalid Row/Col numbers.")

    def generate_canvas_placeholders(self):
        self.canvas.delete("all")
        self.cell_thumbnails.clear()
        
        cw = max(800, self.canvas.winfo_width())
        ch = max(600, self.canvas.winfo_height())

        rows, cols = self.grid_dims
        ar = self.cell_aspect_ratio
        gap = int(self.spin_gap.get()) // 10
        if gap < 2: gap = 5
        
        top_margin = 50
        left_margin = 70
        
        avail_w = cw - (2 * gap) - left_margin
        avail_h = ch - (2 * gap) - top_margin
        
        raw_cell_w = (avail_w - (cols-1)*gap) / cols
        raw_cell_h = (avail_h - (rows-1)*gap) / rows
        
        if raw_cell_w / raw_cell_h > ar:
            self.canvas_cell_h = raw_cell_h
            self.canvas_cell_w = raw_cell_h * ar
        else:
            self.canvas_cell_w = raw_cell_w
            self.canvas_cell_h = raw_cell_w / ar

        start_x = left_margin + gap
        start_y = top_margin + gap
        
        for r in range(rows):
            for c in range(cols):
                x1 = start_x + c * (self.canvas_cell_w + gap)
                y1 = start_y + r * (self.canvas_cell_h + gap)
                x2 = x1 + self.canvas_cell_w
                y2 = y1 + self.canvas_cell_h
                
                self.canvas.create_rectangle(x1, y1, x2, y2, outline="#444444", dash=(4,4), fill="#2a2a2a")
                
                path = self.cell_images.get((r, c))
                if path:
                    self.draw_cell_thumbnail(r, c, x1, y1, x2, y2, path)
                else:
                    btn = tk.Button(self.canvas, text=f"Select Img(s)\n({r},{c})", font=("Arial", 8), bg="#333333", fg="white", command=lambda row=r, col=c: self.select_images_for_cell(row, col))
                    self.canvas.create_window(x1 + self.canvas_cell_w//2, y1 + self.canvas_cell_h//2, window=btn, anchor=tk.CENTER)

    def select_images_for_cell(self, start_row, start_col):
        f_types = [("Microscopy Images", "*.tif *.tiff *.png *.jpg *.jpeg"), ("All Files", "*.*")]
        paths = filedialog.askopenfilenames(title=f"Select Image(s) starting from Row {start_row}, Col {start_col}", filetypes=f_types)
        if not paths: return
        
        self.save_selection_state()
        
        r, c = start_row, start_col
        rows, cols = self.grid_dims
        
        for path in paths:
            if r >= rows: 
                break 
            
            self.cell_images[(r, c)] = path
            if (r, c) not in self.cell_rotations:
                self.cell_rotations[(r, c)] = 0
            
            c += 1
            if c >= cols:
                c = 0
                r += 1
                
        self.cell_aspect_ratio = self.get_cell_aspect()
        self.generate_canvas_placeholders()

    def draw_cell_thumbnail(self, r, c, x1, y1, x2, y2, path):
        try:
            pil_img = Image.open(path)
            rot = self.cell_rotations.get((r, c), 0)
            if rot != 0: pil_img = pil_img.rotate(rot, expand=True)

            target_w = max(1, int(self.canvas_cell_w))
            target_h = max(1, int(self.canvas_cell_h))
            
            thumb = pil_img.resize((target_w, target_h), Image.Resampling.BILINEAR)
            tk_thumb = ImageTk.PhotoImage(thumb)
            
            self.cell_thumbnails[(r, c)] = tk_thumb
            self.canvas.create_image(x1, y1, image=tk_thumb, anchor=tk.NW)
            
            labels_mode = self.combo_sublabels.get()
            if "Inside" in labels_mode:
                label_char = chr(65 + r * self.grid_dims[1] + c)
                hex_lbl = f"#{self.sublabel_color_rgb[0]:02x}{self.sublabel_color_rgb[1]:02x}{self.sublabel_color_rgb[2]:02x}"
                self.canvas.create_text(x1 + 8, y1 + 8, text=f"{label_char}.", fill=hex_lbl, anchor=tk.NW, font=("Arial", 11, "bold"))
            
            text_id = self.canvas.create_text(x2-6, y2-6, text="Change...", fill="#00e676", anchor=tk.SE, font=("Arial", 8, "bold"))
            self.canvas.tag_bind(text_id, "<Button-1>", lambda e, row=r, col=c: self.select_images_for_cell(row, col))
            
            rot_ccw_id = self.canvas.create_text(x1+6, y2-6, text="⟲ CCW", fill="#ff9100", anchor=tk.SW, font=("Arial", 8, "bold"))
            self.canvas.tag_bind(rot_ccw_id, "<Button-1>", lambda e, row=r, col=c: self.rotate_cell(row, col, -90))

            rot_cw_id = self.canvas.create_text(x1+55, y2-6, text="CW ⟳", fill="#ff9100", anchor=tk.SW, font=("Arial", 8, "bold"))
            self.canvas.tag_bind(rot_cw_id, "<Button-1>", lambda e, row=r, col=c: self.rotate_cell(row, col, 90))
            
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load image:\n{str(e)}")
            if (r, c) in self.cell_images: del self.cell_images[(r, c)]

    def action_build_composite(self):
        if not self.cell_images:
            messagebox.showwarning("Warning", "Select images first.")
            return

        rows, cols = self.grid_dims
        gap_orig = int(self.spin_gap.get())
        
        try:
            ref_path = next(iter(self.cell_images.values()))
            with Image.open(ref_path) as img:
                ref_cell_w = img.size[0]
        except:
            ref_cell_w = 1200

        target_cell_w = ref_cell_w
        target_cell_h = int(ref_cell_w / self.cell_aspect_ratio)
        
        row_title_strs = [s.strip() for s in self.entry_row_titles.get().split(',') if s.strip()]
        col_title_strs = [s.strip() for s in self.entry_col_titles.get().split(',') if s.strip()]
        
        labels_mode = self.combo_sublabels.get()
        title_font_size = int(self.spin_title_size.get())
        sublabel_font_size = int(self.spin_sublabel_size.get())
        
        selected_font = self.combo_font.get().lower().replace(" ", "")
        
        def get_pil_font(size_px, bold=False):
            font_map = {
                "arial": ("arialbd.ttf", "arial.ttf"),
                "timesnewroman": ("timesbd.ttf", "times.ttf"),
                "couriernew": ("courbd.ttf", "cour.ttf"),
                "verdana": ("verdanab.ttf", "verdana.ttf"),
                "tahoma": ("tahomabd.ttf", "tahoma.ttf"),
                "georgia": ("georgiab.ttf", "georgia.ttf")
            }
            
            fonts_to_try = font_map.get(selected_font, ("arialbd.ttf", "arial.ttf"))
            if not bold:
                fonts_to_try = (fonts_to_try[1], fonts_to_try[0])
            
            for f in fonts_to_try + ("DejaVuSans-Bold.ttf", "DejaVuSans.ttf"):
                try: return ImageFont.truetype(f, size_px)
                except: continue
            return ImageFont.load_default()

        font_titles = get_pil_font(title_font_size, bold=True)
        font_sublabels = get_pil_font(sublabel_font_size, bold=True)
        
        dummy_img = Image.new("RGBA", (10, 10))
        dummy_draw = ImageDraw.Draw(dummy_img)
        
        max_row_text_w = 0
        for title in row_title_strs:
            bbox = dummy_draw.textbbox((0, 0), title, font=font_titles) if hasattr(dummy_draw, 'textbbox') else (0,0,len(title)*title_font_size*0.6, title_font_size)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            current_w = th if self.combo_row_orient.get() == "Sideways (90° CCW)" else tw
            if current_w > max_row_text_w: max_row_text_w = current_w
            
        max_col_text_h = 0
        for title in col_title_strs:
            bbox = dummy_draw.textbbox((0, 0), title, font=font_titles) if hasattr(dummy_draw, 'textbbox') else (0,0,title_font_size, title_font_size)
            th = bbox[3] - bbox[1]
            if th > max_col_text_h: max_col_text_h = th

        left_margin = int(max_row_text_w + (gap_orig if row_title_strs else gap_orig // 2))
        top_margin = int(max_col_text_h + (gap_orig if col_title_strs else gap_orig // 2))
        if "Outside" in labels_mode: left_margin = max(left_margin, int(sublabel_font_size * 1.5))
            
        left_margin += 20; top_margin += 20
        right_margin = gap_orig // 2 + 20
        bottom_margin = gap_orig // 2 + 20
        
        total_w = left_margin + cols * target_cell_w + (cols - 1) * gap_orig + right_margin
        total_h = top_margin + rows * target_cell_h + (rows - 1) * gap_orig + bottom_margin
        
        self.composite_pil = Image.new("RGBA", (int(total_w), int(total_h)), (255, 255, 255, 255))
        draw = ImageDraw.Draw(self.composite_pil)
        
        title_color_rgba = self.title_color_rgb + (255,)
        sublabel_color_rgba = self.sublabel_color_rgb + (255,)
        COLUMN_TITLE_PADDING = 35
        
        for i, title in enumerate(col_title_strs):
            if i >= cols: break
            col_cx = left_margin + i * (target_cell_w + gap_orig) + target_cell_w // 2
            bbox = draw.textbbox((0, 0), title, font=font_titles) if hasattr(draw, 'textbbox') else (0,0,len(title)*title_font_size*0.6, title_font_size)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]            
            draw.text((col_cx - tw // 2, top_margin - th - COLUMN_TITLE_PADDING), title, fill=title_color_rgba, font=font_titles)
            
        for i, title in enumerate(row_title_strs):
            if i >= rows: break
            row_cy = top_margin + i * (target_cell_h + gap_orig) + target_cell_h // 2
            bbox = draw.textbbox((0, 0), title, font=font_titles) if hasattr(draw, 'textbbox') else (0,0,len(title)*title_font_size*0.6, title_font_size)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            
            if self.combo_row_orient.get() == "Sideways (90° CCW)":
                txt_w, txt_h = max(1, int(tw + 4)), max(1, int(th + 4))
                text_layer = Image.new("RGBA", (txt_w, txt_h), (0, 0, 0, 0))
                layer_draw = ImageDraw.Draw(text_layer)
                layer_draw.text((-bbox[0], -bbox[1]), title, fill=title_color_rgba, font=font_titles)
                rotated_txt = text_layer.rotate(90, expand=True)
                rt_w, rt_h = rotated_txt.size
                self.composite_pil.alpha_composite(rotated_txt, (int(left_margin - rt_w - 20), int(row_cy - rt_h // 2)))
            else:
                draw.text((left_margin - tw - 20, row_cy - th // 2), title, fill=title_color_rgba, font=font_titles)

        for r in range(rows):
            for c in range(cols):
                path = self.cell_images.get((r, c))
                px = left_margin + c * (target_cell_w + gap_orig)
                py = top_margin + r * (target_cell_h + gap_orig)
                
                if path:
                    try:
                        with Image.open(path) as img:
                            rot = self.cell_rotations.get((r, c), 0)
                            if rot != 0: img = img.rotate(rot, expand=True)
                            scaled_img = img.convert("RGBA").resize((int(target_cell_w), int(target_cell_h)), Image.Resampling.LANCZOS)
                            self.composite_pil.alpha_composite(scaled_img, (int(px), int(py)))
                    except Exception:
                        draw.rectangle([px, py, px+target_cell_w, py+target_cell_h], outline="red", width=4)
                else:
                    draw.rectangle([px, py, px+target_cell_w, py+target_cell_h], fill=(240, 240, 240, 255), outline=(200, 200, 200, 255), width=2)
                
                label_char = chr(65 + r * cols + c)
                if "Inside" in labels_mode:
                    offset = sublabel_font_size // 2
                    draw.text((px + offset, py + offset), f"{label_char}.", fill=sublabel_color_rgba, font=font_sublabels)
                elif "Outside" in labels_mode:
                    draw.text((px - int(sublabel_font_size * 0.9), py), f"{label_char}.", fill=sublabel_color_rgba, font=font_sublabels)
        
        # Reset Viewport
        cw = max(50, self.canvas.winfo_width())
        ch = max(50, self.canvas.winfo_height())
        iw, ih = self.composite_pil.size
        
        self.view_scale = min((cw - 40) / iw, (ch - 40) / ih)
        if self.view_scale > 1.0: self.view_scale = 1.0
        
        self.pan_x = (cw - (iw * self.view_scale)) / 2
        self.pan_y = (ch - (ih * self.view_scale)) / 2

        self.render_canvas_preview()
        messagebox.showinfo("Success", "High-Resolution composite built. You can now zoom (scroll), pan (middle click), and draw boxes (left click) directly on it.")

    # --- SEPARATED RENDERING ENGINE ---
    
    def render_canvas_preview(self, event=None):
        if self.composite_pil is None:
             if self.grid_dims: self.generate_canvas_placeholders()
             return

        # Crucial Fix: Wipe the entire canvas to clear out the interactive placeholder text
        self.canvas.delete("all")
        
        target_w = max(1, int(self.composite_pil.size[0] * self.view_scale))
        target_h = max(1, int(self.composite_pil.size[1] * self.view_scale))
        
        resized_comp = self.composite_pil.resize((target_w, target_h), Image.Resampling.BILINEAR)
        self.tk_img = ImageTk.PhotoImage(resized_comp)
        
        self.canvas.create_image(self.pan_x, self.pan_y, anchor=tk.NW, image=self.tk_img, tags="bg")
        self.canvas.tag_lower("bg")
        
        self.draw_interactive_elements()

    def draw_interactive_elements(self):
        self.canvas.delete("overlay")
        for i, box in enumerate(self.panel_boxes):
            x1, y1 = self.img_to_screen(box['x1'], box['y1'])
            x2, y2 = self.img_to_screen(box['x2'], box['y2'])
            
            is_selected = (i == self.selected_box_idx)
            outline_color = "#FFD700" if is_selected else "#FFFFFF"
            dash_pattern = (4, 4) if not is_selected else None

            self.canvas.create_rectangle(x1, y1, x2, y2, outline=outline_color, width=2, dash=dash_pattern, tags="overlay")
            
            if is_selected:
                # NEW: Only draw white resize handles if the box is scalable
                if box.get('scalable', True):
                    handles = self.get_handle_rects(box)
                    for hx1, hy1, hx2, hy2 in handles.values():
                        self.canvas.create_rectangle(hx1, hy1, hx2, hy2, fill="white", outline="black", tags="overlay")

    def draw_dashed_line(self, draw, pt1, pt2, fill, width, dash_len=10):
        x1, y1 = pt1
        x2, y2 = pt2
        dist = ((x2 - x1)**2 + (y2 - y1)**2)**0.5
        if dist == 0: return
        dx, dy = (x2 - x1) / dist, (y2 - y1) / dist
        
        for i in range(0, int(dist), dash_len * 2):
            start = (x1 + dx * i, y1 + dy * i)
            end = (x1 + dx * min(i + dash_len, dist), y1 + dy * min(i + dash_len, dist))
            draw.line([start, end], fill=fill, width=width)

    def action_export_panel(self, event=None):
        if self.composite_pil is None:
            if not self.cell_images:
                messagebox.showwarning("Warning", "Assign images first.")
                return
            self.action_build_composite()
            if self.composite_pil is None: return
            
        f_types = [("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("TIFF Image", "*.tif")]
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=f_types, title="Export Final Panel")
        if not path: return
        
        try:
            export_pil = self.composite_pil.copy()
            draw = ImageDraw.Draw(export_pil)
            
            img_w, img_h = export_pil.size
            dyn_thick = max(3, int(min(img_w, img_h) * 0.003))
            dash_len = dyn_thick * 4

            for box in self.panel_boxes:
                x1, y1 = box['x1'], box['y1']
                x2, y2 = box['x2'], box['y2']
                self.draw_dashed_line(draw, (x1, y1), (x2, y1), fill="yellow", width=dyn_thick, dash_len=dash_len)
                self.draw_dashed_line(draw, (x1, y2), (x2, y2), fill="yellow", width=dyn_thick, dash_len=dash_len)
                self.draw_dashed_line(draw, (x1, y1), (x1, y2), fill="yellow", width=dyn_thick, dash_len=dash_len)
                self.draw_dashed_line(draw, (x2, y1), (x2, y2), fill="yellow", width=dyn_thick, dash_len=dash_len)

            _, ext = os.path.splitext(path)
            mode = "RGBA" if ext.lower() in [".png", ".tif", ".tiff"] else "RGB"
            final_export = export_pil.convert(mode)
            final_export.save(path)
            
            messagebox.showinfo("Export Complete", f"Figure exported successfully:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save asset:\n{str(e)}")