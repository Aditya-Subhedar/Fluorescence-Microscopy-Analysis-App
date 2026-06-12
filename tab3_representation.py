import os
import math
import copy
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont

class MaskMergerTab(ttk.Frame):
    def __init__(self, parent, main_app):
        super().__init__(parent)
        self.main_app = main_app
        
        # Layer Management
        self.microscope_image_path = None 
        self.mask_paths = []
        self.composite_pil = None
        self.tk_img = None
        self.global_rotation = 0  # Tracks absolute rotation state
        
        # Annotation & Gesture Variables
        self.annotations = []
        self.selected_idx = None
        self.active_text_color_rgb = (255, 255, 255) 
        self.bg_color_rgb = None 
        
        # Undo / Redo Stacks
        self.undo_stack = [[]]
        self.redo_stack = []
        self.action_occurred = False
        
        # Interaction States
        self.interaction_state = "Select"
        self.draw_start_pt = None
        self.temp_draw_data = None
        
        self.setup_ui()
        self.setup_keybindings()

    def setup_ui(self):
        root_frame = tk.Frame(self)
        root_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # --- LEFT PANEL ---
        left_panel = tk.Frame(root_frame, width=260)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # 1. Mask Layers Queue
        tk.Label(left_panel, text="Outline Mask Layers Queue:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(5, 5))
        self.list_box = tk.Listbox(left_panel, selectmode=tk.SINGLE, font=("Arial", 9), height=4)
        self.list_box.pack(fill=tk.X, pady=2)
        
        btn_frame = tk.Frame(left_panel)
        btn_frame.pack(fill=tk.X, pady=2)
        tk.Button(btn_frame, text="➕ Add Mask File", command=self.add_mask_file, bg="#1976d2", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=1)
        tk.Button(btn_frame, text="❌ Remove Selected", command=self.remove_mask_file, bg="#d32f2f", fg="white").pack(fill=tk.X, pady=1)
        
        # 2. Background Settings
        bg_frame = tk.LabelFrame(left_panel, text="🖼️ Background Settings", font=("Arial", 9, "bold"), padx=5, pady=5)
        bg_frame.pack(fill=tk.X, pady=5)
        
        self.btn_load_microscope = tk.Button(bg_frame, text="🔬 Load Base Image", command=self.load_microscope_image, bg="#673ab7", fg="white")
        self.btn_load_microscope.pack(fill=tk.X, pady=(2, 0))
        
        self.lbl_microscope_name = tk.Label(bg_frame, text="No base image loaded", fg="gray", font=("Arial", 8, "italic"), wraplength=230)
        self.lbl_microscope_name.pack(fill=tk.X, pady=(0, 4))
        
        self.btn_bg_color = tk.Button(bg_frame, text="🎨 Select Solid Color", command=self.pick_background_color, bg="#424242", fg="white")
        self.btn_bg_color.pack(fill=tk.X, pady=2)

        # ROTATION CONTROLS
        rot_frame = tk.Frame(bg_frame)
        rot_frame.pack(fill=tk.X, pady=(4, 2))
        tk.Button(rot_frame, text="↺ CCW", command=self.action_rotate_ccw, bg="#ff9800", fg="black").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        tk.Button(rot_frame, text="↻ CW", command=self.action_rotate_cw, bg="#ff9800", fg="black").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))
        
        # 3. Interactive Figure Tools
        tool_frame = tk.LabelFrame(left_panel, text="🔤 Interactive Tools", font=("Arial", 9, "bold"), padx=5, pady=5)
        tool_frame.pack(fill=tk.X, pady=5)
        
        self.tool_mode = tk.StringVar(value="Select")
        modes = [
            ("👆 Select / Edit", "Select"),
            ("🔲 Draw Dashed Box", "Box"),
            ("↗️ Draw Arrow", "Arrow"),
            ("🔤 Add Text (Click Canvas)", "Text")
        ]
        for text, mode in modes:
            tk.Radiobutton(tool_frame, text=text, variable=self.tool_mode, value=mode, 
                           command=self.on_tool_change, indicatoron=0, width=20, 
                           selectcolor="#b3e5fc", pady=4).pack(fill=tk.X, pady=1)
        
        # Slider Settings
        self.slider_frame = tk.Frame(tool_frame)
        self.slider_frame.pack(fill=tk.X, pady=(8, 2))
        tk.Label(self.slider_frame, text="🔍 Adjust Size / Scale:").pack(anchor=tk.W)
        self.size_var = tk.DoubleVar(value=32)
        self.size_slider = ttk.Scale(self.slider_frame, from_=10, to_=400, variable=self.size_var, command=self.on_slider_slide)
        self.size_slider.pack(fill=tk.X)
        self.size_slider.bind("<ButtonRelease-1>", lambda e: self.save_state())  # Only save state when slider is released
        
        self.btn_text_color = tk.Button(tool_frame, text="🎨 Set Annotation Color", command=self.pick_text_color, bg="#ffffff", fg="black")
        self.btn_text_color.pack(fill=tk.X, pady=6)
        
        # Undo / Redo / Delete row
        action_frame = tk.Frame(tool_frame)
        action_frame.pack(fill=tk.X, pady=2)
        tk.Button(action_frame, text="↩️ Undo", command=self.undo).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        tk.Button(action_frame, text="↪️ Redo", command=self.redo).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))
        tk.Button(tool_frame, text="🗑️ Delete Selected Item", command=self.delete_selected, fg="red").pack(fill=tk.X, pady=4)
        
        # Export Button
        self.btn_export = tk.Button(left_panel, text="💾 Export Merged Figure", command=self.export_merged_figure, bg="#2e7d32", fg="white", font=("Arial", 11, "bold"), height=2)
        self.btn_export.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
        
        # --- RIGHT CANVAS ---
        self.preview_frame = tk.Frame(root_frame, bg="black")
        self.preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.preview_frame, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<Configure>", lambda e: self.render_preview())
        
        self.on_tool_change()

    def setup_keybindings(self):
        toplevel = self.winfo_toplevel()
        toplevel.bind("<Control-z>", self.undo)
        toplevel.bind("<Control-y>", self.redo)
        toplevel.bind("<Control-s>", self.export_merged_figure)

    # --- ACTION HANDLERS ---
    def action_rotate_cw(self):
        """Rotates 90° Clockwise"""
        if not self.composite_pil: return
        old_w, old_h = self.composite_pil.size
        
        # PIL rotate is CCW by default, so subtract 90 for CW
        self.global_rotation = (self.global_rotation - 90) % 360
        
        # Matrix translation for annotations to match CW rotation
        for anno in self.annotations:
            old_x, old_y = anno["x"], anno["y"]
            anno["x"] = old_h - old_y
            anno["y"] = old_x
            if "angle" in anno:
                anno["angle"] = (anno["angle"] + 90) % 360
                
        self.generate_composite()
        self.save_state()

    def action_rotate_ccw(self):
        """Rotates 90° Counter-Clockwise"""
        if not self.composite_pil: return
        old_w, old_h = self.composite_pil.size
        
        # PIL rotate is CCW by default, so add 90 for CCW
        self.global_rotation = (self.global_rotation + 90) % 360
        
        # Matrix translation for annotations to match CCW rotation
        for anno in self.annotations:
            old_x, old_y = anno["x"], anno["y"]
            anno["x"] = old_y
            anno["y"] = old_w - old_x
            if "angle" in anno:
                anno["angle"] = (anno["angle"] - 90) % 360
                
        self.generate_composite()
        self.save_state()

    # --- UNDO / REDO LOGIC ---
    def save_state(self):
        self.undo_stack.append(copy.deepcopy(self.annotations))
        self.redo_stack.clear()
        if len(self.undo_stack) > 50: 
            self.undo_stack.pop(0)

    def undo(self, event=None):
        self.finalize_inline_text()  
        if len(self.undo_stack) > 1:
            self.redo_stack.append(self.undo_stack.pop())
            self.annotations = copy.deepcopy(self.undo_stack[-1])
            self.selected_idx = None
            self.render_preview()

    def redo(self, event=None):
        self.finalize_inline_text()
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.annotations = copy.deepcopy(state)
            self.undo_stack.append(state)
            self.selected_idx = None
            self.render_preview()

    def on_tool_change(self):
        self.finalize_inline_text()
        self.selected_idx = None
        self.render_preview()

    # --- MATH & COORD HELPERS ---
    def rotate_point(self, x, y, cx, cy, angle_deg):
        rad = math.radians(angle_deg)
        cos_a, sin_a = math.cos(rad), math.sin(rad)
        return cx + (x - cx) * cos_a - (y - cy) * sin_a, cy + (x - cx) * sin_a + (y - cy) * cos_a

    def image_to_canvas(self, img_x, img_y):
        if not self.composite_pil: return 0, 0
        img_w, img_h = self.composite_pil.size
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        scale = min(cw / img_w, ch / img_h)
        return int(img_x * scale) + (cw - int(img_w * scale)) // 2, int(img_y * scale) + (ch - int(img_h * scale)) // 2

    def canvas_to_image(self, cx, cy):
        if not self.composite_pil: return 0, 0
        img_w, img_h = self.composite_pil.size
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        scale = min(cw / img_w, ch / img_h)
        img_x = (cx - (cw - int(img_w * scale)) // 2) / scale
        img_y = (cy - (ch - int(img_h * scale)) // 2) / scale
        return img_x, img_y

    def get_scale_factor(self):
        if not self.composite_pil: return 1
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        return min(cw / self.composite_pil.size[0], ch / self.composite_pil.size[1])

    # --- SLIDER LOGIC ---
    def update_slider_from_selection(self):
        if self.selected_idx is None: return
        anno = self.annotations[self.selected_idx]
        if anno["type"] == "Text":
            self.size_var.set(anno["size"])
        elif anno["type"] == "Arrow":
            self.size_var.set(anno["length"])
        elif anno["type"] == "Box":
            self.size_var.set(max(anno["w"], anno["h"]))

    def on_slider_slide(self, val):
        if self.selected_idx is None: return
        new_size = float(val)
        anno = self.annotations[self.selected_idx]
        
        if anno["type"] == "Text":
            anno["size"] = new_size
            anno["w"] = anno["size"] * len(anno["text"]) * 0.6
            anno["h"] = anno["size"] * 1.5
        elif anno["type"] == "Arrow":
            anno["length"] = new_size
        elif anno["type"] == "Box":
            old_max = max(anno["w"], anno["h"])
            if old_max > 0:
                ratio = new_size / old_max
                anno["w"] *= ratio
                anno["h"] *= ratio
        self.render_preview()

    # --- INLINE TEXT ENTRY ---
    def finalize_inline_text(self, event=None):
        if hasattr(self, 'inline_entry') and self.inline_entry.winfo_exists():
            text_val = self.inline_entry.get().strip()
            self.inline_entry.destroy()
            if text_val:
                self.annotations.append({
                    "type": "Text", "text": text_val, "x": self.inline_x, "y": self.inline_y,
                    "w": len(text_val) * 19, "h": 48, "size": 32, "angle": 0, "color": self.active_text_color_rgb
                })
                self.save_state()
                self.selected_idx = len(self.annotations) - 1
                self.tool_mode.set("Select")
                self.update_slider_from_selection()
            self.render_preview()

    # --- GESTURE LOGIC ---
    def on_mouse_down(self, event):
        if not self.composite_pil: return
        self.finalize_inline_text()
        
        self.action_occurred = False
        img_x, img_y = self.canvas_to_image(event.x, event.y)
        mode = self.tool_mode.get()
        
        if mode == "Select":
            self.interaction_state = "Select"
            if self.selected_idx is not None:
                anno = self.annotations[self.selected_idx]
                cx, cy = self.image_to_canvas(anno["x"], anno["y"])
                scale = self.get_scale_factor()
                
                if anno["type"] in ["Box", "Text"]:
                    bw = anno.get("w", 100) * scale / 2
                    bh = anno.get("h", 30) * scale / 2
                    rot_hx, rot_hy = self.rotate_point(cx, cy - bh - 25, cx, cy, anno["angle"])
                    scl_hx, scl_hy = self.rotate_point(cx + bw, cy + bh, cx, cy, anno["angle"])
                else: 
                    al = (anno["length"] * scale) / 2
                    rot_hx, rot_hy = self.rotate_point(cx, cy - 30, cx, cy, anno["angle"])
                    scl_hx, scl_hy = self.rotate_point(cx + al, cy, cx, cy, anno["angle"])

                if math.hypot(event.x - rot_hx, event.y - rot_hy) < 15:
                    self.interaction_state = "Rotate"
                    return
                if math.hypot(event.x - scl_hx, event.y - scl_hy) < 15:
                    self.interaction_state = "Scale"
                    return

            self.selected_idx = None
            for idx, anno in enumerate(reversed(self.annotations)):
                real_idx = len(self.annotations) - 1 - idx
                ax, ay = self.image_to_canvas(anno["x"], anno["y"])
                if math.hypot(event.x - ax, event.y - ay) < 30: 
                    self.selected_idx = real_idx
                    self.interaction_state = "Move"
                    self.draw_start_pt = (event.x, event.y)
                    self.update_slider_from_selection()
                    self.render_preview()
                    return
            self.render_preview() 

        elif mode in ["Box", "Arrow"]:
            self.interaction_state = f"Draw_{mode}"
            self.draw_start_pt = (img_x, img_y)
            self.temp_draw_data = {"start": (img_x, img_y), "end": (img_x, img_y)}

        elif mode == "Text":
            self.inline_x, self.inline_y = img_x, img_y
            self.inline_entry = tk.Entry(self.canvas, font=("Arial", 16), justify="center", bg="#333333", fg="white", insertbackground="white")
            self.inline_entry.place(x=event.x, y=event.y, anchor=tk.CENTER)
            self.inline_entry.focus_set()
            self.inline_entry.bind("<Return>", self.finalize_inline_text)

    def on_mouse_drag(self, event):
        if hasattr(self, 'inline_entry') and self.inline_entry.winfo_exists(): return
        if not self.composite_pil or self.interaction_state == "Select": return
        self.action_occurred = True
        img_x, img_y = self.canvas_to_image(event.x, event.y)
        
        if self.interaction_state == "Move" and self.selected_idx is not None:
            dx_img = img_x - self.canvas_to_image(*self.draw_start_pt)[0]
            dy_img = img_y - self.canvas_to_image(*self.draw_start_pt)[1]
            self.annotations[self.selected_idx]["x"] += dx_img
            self.annotations[self.selected_idx]["y"] += dy_img
            self.draw_start_pt = (event.x, event.y)

        elif self.interaction_state == "Rotate" and self.selected_idx is not None:
            anno = self.annotations[self.selected_idx]
            cx, cy = self.image_to_canvas(anno["x"], anno["y"])
            angle_rad = math.atan2(event.y - cy, event.x - cx)
            anno["angle"] = math.degrees(angle_rad) + 90 

        elif self.interaction_state == "Scale" and self.selected_idx is not None:
            anno = self.annotations[self.selected_idx]
            nx, ny = self.rotate_point(img_x, img_y, anno["x"], anno["y"], -anno["angle"])
            if anno["type"] == "Box":
                anno["w"] = max(10, abs(nx - anno["x"]) * 2)
                anno["h"] = max(10, abs(ny - anno["y"]) * 2)
                self.size_var.set(max(anno["w"], anno["h"]))
            elif anno["type"] == "Text":
                anno["size"] = int(max(10, abs(nx - anno["x"])))
                anno["w"] = anno["size"] * len(anno["text"]) * 0.6
                anno["h"] = anno["size"] * 1.5
                self.size_var.set(anno["size"])
            elif anno["type"] == "Arrow":
                dist = math.hypot(img_x - anno["x"], img_y - anno["y"])
                anno["length"] = max(20, dist * 2)
                self.size_var.set(anno["length"])

        elif self.interaction_state.startswith("Draw_"):
            self.temp_draw_data["end"] = (img_x, img_y)

        self.render_preview()

    def on_mouse_up(self, event):
        if not self.composite_pil: return
        
        if self.interaction_state.startswith("Draw_") and self.temp_draw_data:
            sx, sy = self.temp_draw_data["start"]
            ex, ey = self.temp_draw_data["end"]
            
            if math.hypot(ex-sx, ey-sy) > 10: 
                if self.interaction_state == "Draw_Box":
                    self.annotations.append({
                        "type": "Box", "x": (sx+ex)/2, "y": (sy+ey)/2,
                        "w": abs(ex-sx), "h": abs(ey-sy), "angle": 0, "color": self.active_text_color_rgb
                    })
                elif self.interaction_state == "Draw_Arrow":
                    angle = math.degrees(math.atan2(ey-sy, ex-sx))
                    self.annotations.append({
                        "type": "Arrow", "x": (sx+ex)/2, "y": (sy+ey)/2,
                        "length": math.hypot(ex-sx, ey-sy), "angle": angle, "color": self.active_text_color_rgb
                    })
                self.selected_idx = len(self.annotations) - 1
                self.action_occurred = True
            
            self.temp_draw_data = None
            self.tool_mode.set("Select")
            self.update_slider_from_selection()
            self.on_tool_change()
            
        if self.action_occurred:
            self.save_state()
            
        self.interaction_state = "Select"
        self.render_preview()

    def delete_selected(self):
        if self.selected_idx is not None:
            self.annotations.pop(self.selected_idx)
            self.selected_idx = None
            self.save_state()
            self.render_preview()

    # --- RENDER PIPELINE ---
    def render_preview(self):
        if not self.composite_pil: return
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw < 10: cw, ch = 800, 600
        
        scale = self.get_scale_factor()
        resized = self.composite_pil.resize((max(1, int(self.composite_pil.width * scale)), max(1, int(self.composite_pil.height * scale))), Image.Resampling.NEAREST)
        self.tk_img = ImageTk.PhotoImage(resized)
        self.canvas.delete("all")
        self.canvas.create_image(cw//2, ch//2, anchor=tk.CENTER, image=self.tk_img)
        
        for idx, anno in enumerate(self.annotations):
            self.draw_canvas_element(anno, scale, is_selected=(idx == self.selected_idx))
            
        if self.temp_draw_data:
            sx, sy = self.image_to_canvas(*self.temp_draw_data["start"])
            ex, ey = self.image_to_canvas(*self.temp_draw_data["end"])
            color = '#%02x%02x%02x' % self.active_text_color_rgb
            if self.interaction_state == "Draw_Box":
                self.canvas.create_rectangle(sx, sy, ex, ey, outline=color, width=6, dash=(12, 8))
            elif self.interaction_state == "Draw_Arrow":
                self.canvas.create_line(sx, sy, ex, ey, fill=color, width=4, arrow=tk.LAST, arrowshape=(24, 30, 12))

    def draw_canvas_element(self, anno, scale, is_selected=False):
        cx, cy = self.image_to_canvas(anno["x"], anno["y"])
        tk_color = '#%02x%02x%02x' % anno["color"]
        
        if anno["type"] == "Text":
            self.canvas.create_text(cx, cy, text=anno["text"], fill=tk_color, font=("Arial", max(8, int(anno["size"] * scale)), "bold"), anchor=tk.CENTER, angle=-anno["angle"])
        
        elif anno["type"] == "Box":
            bw, bh = (anno["w"] * scale) / 2, (anno["h"] * scale) / 2
            pts = [self.rotate_point(cx+px, cy+py, cx, cy, anno["angle"]) for px, py in [(-bw,-bh), (bw,-bh), (bw,bh), (-bw,bh)]]
            for i in range(4):
                self.canvas.create_line(pts[i][0], pts[i][1], pts[(i+1)%4][0], pts[(i+1)%4][1], fill=tk_color, width=6, dash=(12, 8))
                
        elif anno["type"] == "Arrow":
            al = (anno["length"] * scale) / 2
            x1, y1 = self.rotate_point(cx - al, cy, cx, cy, anno["angle"])
            x2, y2 = self.rotate_point(cx + al, cy, cx, cy, anno["angle"])
            self.canvas.create_line(x1, y1, x2, y2, fill=tk_color, width=4, arrow=tk.LAST, arrowshape=(24, 30, 12))

        if is_selected:
            if anno["type"] in ["Box", "Text"]:
                bw = anno.get("w", 100) * scale / 2
                bh = anno.get("h", 30) * scale / 2
                rot_x, rot_y = self.rotate_point(cx, cy - bh - 25, cx, cy, anno["angle"])
                anc_x, anc_y = self.rotate_point(cx, cy - bh, cx, cy, anno["angle"])
                scl_x, scl_y = self.rotate_point(cx + bw, cy + bh, cx, cy, anno["angle"])
            else: 
                al = (anno["length"] * scale) / 2
                rot_x, rot_y = self.rotate_point(cx, cy - 30, cx, cy, anno["angle"])
                anc_x, anc_y = cx, cy
                scl_x, scl_y = self.rotate_point(cx + al, cy, cx, cy, anno["angle"])

            self.canvas.create_line(anc_x, anc_y, rot_x, rot_y, fill="#00e676", width=1, dash=(2,2)) 
            self.canvas.create_oval(rot_x-6, rot_y-6, rot_x+6, rot_y+6, fill="#00e676", outline="white", width=2) 
            self.canvas.create_rectangle(scl_x-6, scl_y-6, scl_x+6, scl_y+6, fill="#ffeb3b", outline="black", width=1) 
            self.canvas.create_oval(cx-3, cy-3, cx+3, cy+3, fill="red") 

    # --- FILE / BACKGROUND SETTINGS ---
    def load_microscope_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Images", "*.tif *.tiff *.png *.jpg *.jpeg")], title="Select Background")
        if not file_path: return
        self.microscope_image_path = file_path
        self.lbl_microscope_name.config(text=os.path.basename(file_path), fg="lightgreen")
        self.generate_composite()

    def pick_background_color(self):
        c = colorchooser.askcolor(initialcolor="#000000", title="Canvas Background Color")
        if c[0]: 
            self.bg_color_rgb = tuple(int(v) for v in c[0])
            self.btn_bg_color.config(bg=c[1], fg="black" if sum(self.bg_color_rgb)/3 > 128 else "white")
            self.generate_composite()

    def pick_text_color(self):
        c = colorchooser.askcolor(initialcolor="#ffffff", title="Annotation Color")
        if c[0]: 
            self.active_text_color_rgb = tuple(int(v) for v in c[0])
            self.btn_text_color.config(bg=c[1], fg="black" if sum(self.active_text_color_rgb)/3 > 128 else "white")

    def add_mask_file(self):
        files = filedialog.askopenfilenames(filetypes=[("PNG", "*.png")])
        for f in files:
            if f not in self.mask_paths:
                self.mask_paths.append(f)
                self.list_box.insert(tk.END, os.path.basename(f))
        self.generate_composite()

    def remove_mask_file(self):
        idx = self.list_box.curselection()
        if idx:
            self.list_box.delete(idx[0])
            self.mask_paths.pop(idx[0])
            self.generate_composite()

    def generate_composite(self):
        if not self.mask_paths and not self.microscope_image_path:
            self.composite_pil = None
            self.canvas.delete("all")
            return
        
        # Load Base Image
        if self.microscope_image_path:
            master = Image.open(self.microscope_image_path).convert("RGBA")
        else:
            first = Image.open(self.mask_paths[0]).convert("RGBA")
            master = Image.new("RGBA", first.size, self.bg_color_rgb + (255,) if self.bg_color_rgb else (0,0,0,255))
            
        # Apply Global Rotation
        if self.global_rotation % 360 != 0:
            master = master.rotate(self.global_rotation, expand=True, resample=Image.Resampling.BICUBIC)
            
        w, h = master.size
        
        # Overlay Masks
        for path in self.mask_paths:
            layer = Image.open(path).convert("RGBA")
            if self.global_rotation % 360 != 0:
                layer = layer.rotate(self.global_rotation, expand=True, resample=Image.Resampling.BICUBIC)
            if layer.size != (w, h): 
                layer = layer.resize((w, h), Image.Resampling.NEAREST)
            master.alpha_composite(layer)
            
        self.composite_pil = master
        self.render_preview()

    def export_merged_figure(self, event=None):
        self.finalize_inline_text()
        if not self.composite_pil: return messagebox.showwarning("Empty", "Add layout sheets before exporting.")
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
        if not path: return
        
        w, h = self.composite_pil.size
        
        # Build Export Master Setup with Rotation Logic
        if self.microscope_image_path:
            exp = Image.open(self.microscope_image_path).convert("RGBA")
            if self.global_rotation % 360 != 0:
                exp = exp.rotate(self.global_rotation, expand=True, resample=Image.Resampling.BICUBIC)
        else:
            exp = Image.new("RGBA", (w, h), self.bg_color_rgb + (255,) if self.bg_color_rgb else (0,0,0,255))
            
        for p in self.mask_paths:
            l = Image.open(p).convert("RGBA")
            if self.global_rotation % 360 != 0:
                l = l.rotate(self.global_rotation, expand=True, resample=Image.Resampling.BICUBIC)
            if l.size != (w, h): 
                l = l.resize((w, h), Image.Resampling.NEAREST)
            exp.alpha_composite(l)
            
        # Embed Annotations
        draw = ImageDraw.Draw(exp)
        for a in self.annotations:
            rgba = a["color"] + (255,)
            if a["type"] == "Text":
                try: font = ImageFont.truetype("arial.ttf", int(a["size"]))
                except: font = ImageFont.load_default()
                bbox = draw.textbbox((0,0), a["text"], font=font) if hasattr(draw, 'textbbox') else (0,0,10,10)
                txt_img = Image.new('RGBA', (bbox[2]-bbox[0]+20, bbox[3]-bbox[1]+20), (0,0,0,0))
                ImageDraw.Draw(txt_img).text((10,10), a["text"], fill=rgba, font=font)
                rotated = txt_img.rotate(-a["angle"], expand=1, resample=Image.Resampling.BICUBIC)
                exp.alpha_composite(rotated, (int(a["x"] - rotated.width//2), int(a["y"] - rotated.height//2)))
                
            elif a["type"] == "Box":
                bw, bh = a["w"]/2, a["h"]/2
                pts = [self.rotate_point(a["x"]+px, a["y"]+py, a["x"], a["y"], a["angle"]) for px, py in [(-bw,-bh), (bw,-bh), (bw,bh), (-bw,bh)]]
                for i in range(4):
                    dx, dy = pts[(i+1)%4][0] - pts[i][0], pts[(i+1)%4][1] - pts[i][1]
                    dist = math.hypot(dx, dy)
                    if dist == 0: continue
                    dash_len, gap_len = 16, 10
                    step_x, step_y = dx / dist, dy / dist
                    curr_dist = 0
                    while curr_dist < dist:
                        start_x = pts[i][0] + curr_dist * step_x
                        start_y = pts[i][1] + curr_dist * step_y
                        curr_dist += dash_len
                        end_x = pts[i][0] + min(dist, curr_dist) * step_x
                        end_y = pts[i][1] + min(dist, curr_dist) * step_y
                        draw.line([(start_x, start_y), (end_x, end_y)], fill=rgba, width=6)
                        curr_dist += gap_len
                        
            elif a["type"] == "Arrow":
                al = a["length"]/2
                x1, y1 = self.rotate_point(a["x"]-al, a["y"], a["x"], a["y"], a["angle"])
                x2, y2 = self.rotate_point(a["x"]+al, a["y"], a["x"], a["y"], a["angle"])
                draw.line([(x1, y1), (x2, y2)], fill=rgba, width=4)
                actual_angle = math.atan2(y2 - y1, x2 - x1)
                hl, hw = 30, 16 
                back_x = x2 - hl * math.cos(actual_angle)
                back_y = y2 - hl * math.sin(actual_angle)
                p1_x = back_x + hw * math.cos(actual_angle + math.pi/2)
                p1_y = back_y + hw * math.sin(actual_angle + math.pi/2)
                p2_x = back_x + hw * math.cos(actual_angle - math.pi/2)
                p2_y = back_y + hw * math.sin(actual_angle - math.pi/2)
                draw.polygon([(x2, y2), (p1_x, p1_y), (p2_x, p2_y)], fill=rgba)
                
        exp.save(path, "PNG")
        messagebox.showinfo("Export Complete", f"Saved:\n{path}")