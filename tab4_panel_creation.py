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
        self.cell_images = {}  # key: (row, col), value: absolute filepath
        self.grid_dims = (2, 2)
        self.cell_aspect_ratio = 1.0  # Width/Height (1.0 = Square)
        
        # Previews
        self.composite_pil = None  # Master high-res image
        self.tk_img = None        # Scaled canvas image
        self.cell_thumbnails = {} # Scaled thumbnails for canvas

        # Annotation / Figure settings
        self.sublabel_color_rgb = (255, 255, 255)  # Labels inside image default to White
        self.title_color_rgb = (0, 0, 0)          # Row/Col titles always Black
        
        # Undo/Redo (for image selections)
        self.undo_stack = [[]]
        self.redo_stack = []

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
        self.combo_aspect = ttk.Combobox(grid_frame, values=["Square (1:1)", "Landscape (4:3)", "Portrait (3:4)"], state="readonly")
        self.combo_aspect.pack(fill=tk.X, pady=2)
        self.combo_aspect.set("Square (1:1)")
        
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
        
        labels_config_frame = tk.Frame(titles_frame)
        labels_config_frame.pack(fill=tk.X, pady=(5, 0))
        tk.Label(labels_config_frame, text="Subpanel labels (A, B, C):").pack(side=tk.LEFT)
        self.combo_sublabels = ttk.Combobox(labels_config_frame, values=["Inside Top-Left", "Outside Top-Left", "None"], state="readonly", width=15)
        self.combo_sublabels.pack(side=tk.RIGHT)
        self.combo_sublabels.set("Inside Top-Left")

        # 3. Text & Font Settings
        font_frame = tk.LabelFrame(left_panel, text="🎨 Font Settings", padx=5, pady=5)
        font_frame.pack(fill=tk.X, pady=(0, 15))
        
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
        
        self.v_scroll = tk.Scrollbar(self.preview_frame, orient=tk.VERTICAL)
        self.h_scroll = tk.Scrollbar(self.preview_frame, orient=tk.HORIZONTAL)
        self.canvas = tk.Canvas(self.preview_frame, bg="#1a1a1a", highlightthickness=0, yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        
        self.v_scroll.config(command=self.canvas.yview)
        self.h_scroll.config(command=self.canvas.xview)
        
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<Configure>", lambda e: self.render_canvas_preview())

    def setup_keybindings(self):
        toplevel = self.winfo_toplevel()
        toplevel.bind("<Control-z>", self.undo)
        toplevel.bind("<Control-y>", self.redo)
        toplevel.bind("<Control-s>", self.action_export_panel)
        toplevel.bind("<Control-g>", self.action_generate_grid)

    def save_selection_state(self):
        state = copy.deepcopy(self.cell_images)
        self.undo_stack.append(state)
        self.redo_stack.clear()
        if len(self.undo_stack) > 50: self.undo_stack.pop(0)

    def undo(self, event=None):
        if len(self.undo_stack) > 1:
            self.redo_stack.append(self.undo_stack.pop())
            self.cell_images = copy.deepcopy(self.undo_stack[-1])
            self.generate_canvas_placeholders()
        else:
            messagebox.showinfo("Info", "Nothing to undo in selections.")

    def redo(self, event=None):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.cell_images = copy.deepcopy(state)
            self.undo_stack.append(state)
            self.generate_canvas_placeholders()
        else:
            messagebox.showinfo("Info", "Nothing to redo in selections.")

    def get_cell_aspect(self):
        choice = self.combo_aspect.get()
        if "Square" in choice: return 1.0
        if "Landscape" in choice: return 4.0 / 3.0
        if "Portrait" in choice: return 3.0 / 4.0
        return 1.0

    def pick_sublabel_color(self):
        """Allows modifying color fields exclusively for internal panel markings."""
        color_choice = colorchooser.askcolor(initialcolor="#ffffff", title="Select Subpanel Label Color")
        if color_choice and color_choice[0] is not None:
            self.sublabel_color_rgb = tuple(int(c) for c in color_choice[0])
            hex_color = color_choice[1]
            btn_fg = "black" if (sum(self.sublabel_color_rgb)/3) > 128 else "white"
            self.btn_font_color.config(text=f"Label: {hex_color}", bg=hex_color, fg=btn_fg)

    def action_generate_grid(self, event=None):
        try:
            r = int(self.spin_rows.get())
            c = int(self.spin_cols.get())
            self.grid_dims = (r, c)
            self.cell_aspect_ratio = self.get_cell_aspect()
            
            self.composite_pil = None 
            self.tk_img = None
            
            if list(self.cell_images.keys()) and (r != len(set(k[0] for k in self.cell_images.keys())) or c != len(set(k[1] for k in self.cell_images.keys()))):
                if messagebox.askyesno("Clear Images?", "Grid dimensions changed. Clear currently selected images?"):
                    self.cell_images.clear()
            
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
        left_margin = 60
        
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
                
                rect_id = self.canvas.create_rectangle(x1, y1, x2, y2, outline="#444444", dash=(4,4), fill="#2a2a2a")
                
                path = self.cell_images.get((r, c))
                if path:
                    self.draw_cell_thumbnail(r, c, x1, y1, x2, y2, path)
                else:
                    btn = tk.Button(self.canvas, text=f"Select Img\n({r},{c})", font=("Arial", 8), bg="#333333", fg="white", command=lambda row=r, col=c: self.select_image_for_cell(row, col))
                    self.canvas.create_window(x1 + self.canvas_cell_w//2, y1 + self.canvas_cell_h//2, window=btn, anchor=tk.CENTER)

        self.canvas.config(scrollregion=(0, 0, cw, ch))

    def select_image_for_cell(self, row, col):
        f_types = [("Microscopy Images", "*.tif *.tiff *.png *.jpg *.jpeg"), ("All Files", "*.*")]
        path = filedialog.askopenfilename(title=f"Select Image for Row {row}, Col {col}", filetypes=f_types)
        if not path: return
        
        self.save_selection_state()
        self.cell_images[(row, col)] = path
        self.generate_canvas_placeholders()

    def draw_cell_thumbnail(self, r, c, x1, y1, x2, y2, path):
        try:
            pil_img = Image.open(path)
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
            self.canvas.tag_bind(text_id, "<Button-1>", lambda e, row=r, col=c: self.select_image_for_cell(row, col))
            
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
                ref_cell_w, _ = img.size
        except:
            ref_cell_w = 1200

        target_cell_w = ref_cell_w
        target_cell_h = int(ref_cell_w / self.cell_aspect_ratio)
        
        row_title_strs = [s.strip() for s in self.entry_row_titles.get().split(',') if s.strip()]
        col_title_strs = [s.strip() for s in self.entry_col_titles.get().split(',') if s.strip()]
        
        labels_mode = self.combo_sublabels.get()
        title_font_size = int(self.spin_title_size.get())
        sublabel_font_size = int(self.spin_sublabel_size.get())
        
        def get_pil_font(size_px, bold=False):
            possible_fonts = ["arialbd.ttf" if bold else "arial.ttf", "arial.ttf", "DejaVuSans.ttf"]
            for f in possible_fonts:
                try: return ImageFont.truetype(f, size_px)
                except: continue
            return ImageFont.load_default()

        font_titles = get_pil_font(title_font_size, bold=True)
        font_sublabels = get_pil_font(sublabel_font_size, bold=True)
        
        # --- DYNAMIC MARGIN CALCULATION ENGINE ---
        dummy_img = Image.new("RGBA", (10, 10))
        dummy_draw = ImageDraw.Draw(dummy_img)
        
        max_row_text_w = 0
        for title in row_title_strs:
            bbox = dummy_draw.textbbox((0, 0), title, font=font_titles) if hasattr(dummy_draw, 'textbbox') else (0,0,len(title)*title_font_size*0.6, title_font_size)
            tw = bbox[2] - bbox[0]
            if tw > max_row_text_w: max_row_text_w = tw
            
        max_col_text_h = 0
        for title in col_title_strs:
            bbox = dummy_draw.textbbox((0, 0), title, font=font_titles) if hasattr(dummy_draw, 'textbbox') else (0,0,title_font_size, title_font_size)
            th = bbox[3] - bbox[1]
            if th > max_col_text_h: max_col_text_h = th

        left_margin = int(max_row_text_w + (gap_orig if row_title_strs else gap_orig // 2))
        top_margin = int(max_col_text_h + (gap_orig if col_title_strs else gap_orig // 2))
        
        if "Outside" in labels_mode:
            left_margin = max(left_margin, int(sublabel_font_size * 1.5))
            
        left_margin += 20
        top_margin += 20
        right_margin = gap_orig // 2 + 20
        bottom_margin = gap_orig // 2 + 20
        
        total_w = left_margin + cols * target_cell_w + (cols - 1) * gap_orig + right_margin
        total_h = top_margin + rows * target_cell_h + (rows - 1) * gap_orig + bottom_margin
        
        # --- INITIALIZE WHITE BACKGROUND SHEET ---
        self.composite_pil = Image.new("RGBA", (int(total_w), int(total_h)), (255, 255, 255, 255))
        draw = ImageDraw.Draw(self.composite_pil)
        
        # Color mapping splits
        title_color_rgba = self.title_color_rgb + (255,)       # Always Black
        sublabel_color_rgba = self.sublabel_color_rgb + (255,)   # Adjustable (Default White)
        
        # --- ADD COLUMN TITLE PADDING VARIABLE ---
        COLUMN_TITLE_PADDING = 35
        
        # Render Column Titles (Permanently Black)
        for i, title in enumerate(col_title_strs):
            if i >= cols: break
            col_cx = left_margin + i * (target_cell_w + gap_orig) + target_cell_w // 2
            bbox = draw.textbbox((0, 0), title, font=font_titles) if hasattr(draw, 'textbbox') else (0,0,len(title)*title_font_size*0.6, title_font_size)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]            
            draw.text((col_cx - tw // 2, top_margin - th - COLUMN_TITLE_PADDING), title, fill=title_color_rgba, font=font_titles)
            
        # Render Row Titles (Permanently Black)
        for i, title in enumerate(row_title_strs):
            if i >= rows: break
            row_cy = top_margin + i * (target_cell_h + gap_orig) + target_cell_h // 2
            bbox = draw.textbbox((0, 0), title, font=font_titles) if hasattr(draw, 'textbbox') else (0,0,len(title)*title_font_size*0.6, title_font_size)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            draw.text((left_margin - tw - 20, row_cy - th // 2), title, fill=title_color_rgba, font=font_titles)

        # Compositing High-Res Grid
        for r in range(rows):
            for c in range(cols):
                path = self.cell_images.get((r, c))
                px = left_margin + c * (target_cell_w + gap_orig)
                py = top_margin + r * (target_cell_h + gap_orig)
                
                if path:
                    try:
                        with Image.open(path) as img:
                            scaled_img = img.convert("RGBA").resize((int(target_cell_w), int(target_cell_h)), Image.Resampling.LANCZOS)
                            self.composite_pil.alpha_composite(scaled_img, (int(px), int(py)))
                    except Exception as e:
                        draw.rectangle([px, py, px+target_cell_w, py+target_cell_h], outline="red", width=4)
                else:
                    draw.rectangle([px, py, px+target_cell_w, py+target_cell_h], fill=(240, 240, 240, 255), outline=(200, 200, 200, 255), width=2)
                
                # Apply Internal Labels (Adjustable Color - Default White)
                label_char = chr(65 + r * cols + c)
                label_str = f"{label_char}."
                
                if "Inside" in labels_mode:
                    offset = sublabel_font_size // 2
                    draw.text((px + offset, py + offset), label_str, fill=sublabel_color_rgba, font=font_sublabels)
                elif "Outside" in labels_mode:
                    # Outside prints on the margin gap, adjust fill context if needed
                    draw.text((px - int(sublabel_font_size * 0.9), py), label_str, fill=sublabel_color_rgba, font=font_sublabels)
        
        self.render_canvas_preview()
        messagebox.showinfo("Success", "High-Resolution composite built cleanly with decoupled text properties.")

    def render_canvas_preview(self, event=None):
        if self.composite_pil is None:
             if self.grid_dims: self.generate_canvas_placeholders()
             return

        cw = max(50, self.canvas.winfo_width())
        ch = max(50, self.canvas.winfo_height())

        iw, ih = self.composite_pil.size
        scale = min((cw - 40) / iw, (ch - 40) / ih)
        if scale > 1.0: scale = 1.0
        
        target_w = max(1, int(iw * scale))
        target_h = max(1, int(ih * scale))
        
        resized_comp = self.composite_pil.resize((target_w, target_h), Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(resized_comp)
        
        self.canvas.delete("all")
        self.canvas.create_image(cw//2, ch//2, anchor=tk.CENTER, image=self.tk_img)
        self.canvas.config(scrollregion=(0, 0, cw, ch))

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
            _, ext = os.path.splitext(path)
            mode = "RGBA" if ext.lower() in [".png", ".tif", ".tiff"] else "RGB"
            export_img = self.composite_pil.convert(mode)
            export_img.save(path)
            messagebox.showinfo("Export Complete", f"Figure exported successfully:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save asset:\n{str(e)}")