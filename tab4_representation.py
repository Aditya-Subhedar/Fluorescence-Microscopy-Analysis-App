import os
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont

class MaskMergerTab(ttk.Frame):
    def __init__(self, parent, main_app):
        super().__init__(parent)
        self.main_app = main_app
        
        # Layer Management Stacks
        self.microscope_image_path = None # Absolute background asset path tracker
        self.mask_paths = []
        self.composite_pil = None
        self.tk_img = None
        
        # Annotation Vector Lists
        self.annotations = []
        self.selected_text_idx = None
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.active_text_color_rgb = (255, 255, 255) 
        self.bg_color_rgb = None 
        
        self.setup_ui()

    def setup_ui(self):
        root_frame = tk.Frame(self)
        root_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left Panel Controller (Fixed width to maximize work canvas)
        left_panel = tk.Frame(root_frame, width=250)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # 1. Outline Mask Layers Queue Section
        tk.Label(left_panel, text="Outline Mask Layers Queue:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(5, 10))
        
        self.list_box = tk.Listbox(left_panel, selectmode=tk.SINGLE, font=("Arial", 9), height=8)
        self.list_box.pack(fill=tk.X, pady=2)
        
        # Queue Adjustment Layout Buttons
        btn_frame = tk.Frame(left_panel)
        btn_frame.pack(fill=tk.X, pady=2)
        tk.Button(btn_frame, text="➕ Add Mask File", command=self.add_mask_file, bg="#1976d2", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=1)
        tk.Button(btn_frame, text="❌ Remove Selected Mask", command=self.remove_mask_file, bg="#d32f2f", fg="white").pack(fill=tk.X, pady=1)
        tk.Button(btn_frame, text="🧹 Clear All Masks", command=self.clear_queue).pack(fill=tk.X, pady=1)
        
        # --- CONSOLIDATED: BACKGROUND SETTINGS FRAME ---
        bg_frame = tk.LabelFrame(left_panel, text="🖼️ Background Settings", font=("Arial", 9, "bold"), padx=5, pady=5)
        bg_frame.pack(fill=tk.X, pady=10)
        
        # Microscope Image Sub-controls
        self.btn_load_microscope = tk.Button(bg_frame, text="🔬 Load Microscope Image", command=self.load_microscope_image, bg="#673ab7", fg="white", font=("Arial", 9, "bold"))
        self.btn_load_microscope.pack(fill=tk.X, pady=(2, 0))
        
        self.lbl_microscope_name = tk.Label(bg_frame, text="No base image loaded", fg="gray", font=("Arial", 8, "italic"), wraplength=230)
        self.lbl_microscope_name.pack(fill=tk.X, pady=(0, 4))
        
        # Separator line for visual structure
        ttk.Separator(bg_frame, orient='horizontal').pack(fill=tk.X, pady=6)
        
        # Solid Color and Transparency fallbacks
        self.btn_bg_color = tk.Button(bg_frame, text="🎨 Select Solid Color", command=self.pick_background_color, bg="#424242", fg="white")
        self.btn_bg_color.pack(fill=tk.X, pady=2)
        
        tk.Button(bg_frame, text="❌ Clear / Set Transparent", command=self.reset_to_transparent, fg="red", font=("Arial", 9)).pack(fill=tk.X, pady=2)
        # -----------------------------------------------

        # Figure Annotation Tools
        text_tool_frame = tk.LabelFrame(left_panel, text="🔤 Figure Annotation Tools", font=("Arial", 9, "bold"), padx=5, pady=5)
        text_tool_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(text_tool_frame, text="Label Text:").pack(anchor=tk.W)
        self.entry_text = tk.Entry(text_tool_frame)
        self.entry_text.pack(fill=tk.X, pady=2)
        self.entry_text.insert(0, "")
        
        size_frame = tk.Frame(text_tool_frame)
        size_frame.pack(fill=tk.X, pady=2)
        tk.Label(size_frame, text="Font Size:").pack(side=tk.LEFT)
        self.spin_font_size = tk.Spinbox(size_frame, from_=10, to=150, width=5)
        self.spin_font_size.pack(side=tk.RIGHT)
        self.spin_font_size.delete(0, tk.END)
        self.spin_font_size.insert(0, "24")
        
        self.btn_text_color = tk.Button(text_tool_frame, text="🎨 Choose Text Color", command=self.pick_text_color, bg="#ffffff", fg="black")
        self.btn_text_color.pack(fill=tk.X, pady=4)
        
        tk.Button(text_tool_frame, text="📝 Place Text on Image", command=self.place_text_annotation, bg="#0288d1", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=2)
        tk.Button(text_tool_frame, text="🗑️ Remove Selected Text", command=self.delete_selected_text, fg="red").pack(fill=tk.X, pady=2)
        
        # Master Export Trigger Button
        self.btn_export = tk.Button(
            left_panel, text="💾 Export Merged Figure", 
            command=self.export_merged_figure, bg="#2e7d32", fg="white", 
            font=("Arial", 11, "bold"), height=2
        )
        self.btn_export.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
        
        # Right Presentation Board Canvas Panel Container
        self.preview_frame = tk.Frame(root_frame, bg="black")
        self.preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.preview_frame, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<Configure>", lambda e: self.render_preview())

    def load_microscope_image(self):
        """Loads a TIFF, JPEG, or PNG microscope file to serve as the bottom image layer."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Microscopy Images", "*.tif *.tiff *.png *.jpg *.jpeg"), ("All Files", "*.*")],
            title="Select Raw Microscope Image for Background"
        )
        if not file_path: return
        self.microscope_image_path = file_path
        self.lbl_microscope_name.config(text=os.path.basename(file_path), fg="lightgreen")
        self.generate_composite()

    def reset_to_transparent(self):
        """Clears both microscope images and background color registers back to fully transparent."""
        self.microscope_image_path = None
        self.bg_color_rgb = None
        self.lbl_microscope_name.config(text="No base image loaded", fg="gray")
        self.btn_bg_color.config(text="🎨 Select Solid Color", bg="#424242", fg="white")
        self.generate_composite()

    def pick_background_color(self):
        """Opens native color panel selector to configure a solid publication background sheet."""
        color_choice = colorchooser.askcolor(initialcolor="#000000", title="Select Canvas Background Color")
        if color_choice and color_choice[0] is not None:
            self.bg_color_rgb = tuple(int(c) for c in color_choice[0])
            hex_color = color_choice[1]
            btn_fg = "black" if (sum(self.bg_color_rgb)/3) > 128 else "white"
            self.btn_bg_color.config(text=f"Backdrop: {hex_color}", bg=hex_color, fg=btn_fg)
            self.generate_composite()

    def pick_text_color(self):
        """Configures active text color variables."""
        color_choice = colorchooser.askcolor(initialcolor="#ffffff", title="Select Text Label Color")
        if color_choice and color_choice[0] is not None:
            self.active_text_color_rgb = tuple(int(c) for c in color_choice[0])
            hex_color = color_choice[1]
            btn_fg = "black" if (sum(self.active_text_color_rgb)/3) > 128 else "white"
            self.btn_text_color.config(bg=hex_color, fg=btn_fg)

    def place_text_annotation(self):
        """Places text labels onto screen space indices."""
        text_str = self.entry_text.get().strip()
        if not text_str:
            messagebox.showwarning("Text Field Empty", "Please type descriptive text before placement.")
            return
        if self.composite_pil is None:
            messagebox.showwarning("Empty Canvas", "Load layers or a base image before adding text labels.")
            return
        w, h = self.composite_pil.size
        new_annotation = {
            "text": text_str,
            "x": w // 2,
            "y": h // 2,
            "color": self.active_text_color_rgb,
            "size": int(self.spin_font_size.get()),
            "id": None
        }
        self.annotations.append(new_annotation)
        self.render_preview()

    def delete_selected_text(self):
        """Removes the selected annotation text from the image preview."""
        if self.selected_text_idx is not None and self.selected_text_idx < len(self.annotations):
            self.annotations.pop(self.selected_text_idx)
            self.selected_text_idx = None
            self.render_preview()

    def on_canvas_click(self, event):
        """Initializes selection checking routines."""
        self.selected_text_idx = None
        canvas_x = event.x
        canvas_y = event.y
        for idx, anno in enumerate(reversed(self.annotations)):
            real_idx = len(self.annotations) - 1 - idx
            cx, cy = self.image_to_canvas_coords(anno["x"], anno["y"])
            click_radius = anno["size"] + 10
            if abs(canvas_x - cx) < click_radius and abs(canvas_y - cy) < (click_radius // 2):
                self.selected_text_idx = real_idx
                self.drag_start_x = canvas_x
                self.drag_start_y = canvas_y
                self.canvas.create_rectangle(cx - click_radius, cy - 15, cx + click_radius, cy + 15, outline="#00e676", width=2, tags="select_ring")
                break

    def on_canvas_drag(self, event):
        """Tracks delta movements to scale true positioning components."""
        if self.selected_text_idx is None or self.composite_pil is None: return
        dx_canvas = event.x - self.drag_start_x
        dy_canvas = event.y - self.drag_start_y
        img_w, img_h = self.composite_pil.size
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        scale = min(canvas_w / img_w, canvas_h / img_h)
        if scale <= 0: return
        dx_img = dx_canvas / scale
        dy_img = dy_canvas / scale
        anno = self.annotations[self.selected_text_idx]
        anno["x"] = int(max(0, min(img_w, anno["x"] + dx_img)))
        anno["y"] = int(max(0, min(img_h, anno["y"] + dy_img)))
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.render_preview()

    def image_to_canvas_coords(self, img_x, img_y):
        """Translates coordinate structures positionally across active viewport frames."""
        if self.composite_pil is None: return (0, 0)
        img_w, img_h = self.composite_pil.size
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        scale = min(canvas_w / img_w, canvas_h / img_h)
        new_w = int(img_w * scale)
        new_h = int(img_h * scale)
        offset_x = (canvas_w - new_w) // 2
        offset_y = (canvas_h - new_h) // 2
        cx = int(img_x * scale) + offset_x
        cy = int(img_y * scale) + offset_y
        return cx, cy

    def add_mask_file(self):
        """Appends new files to the mask sequence list box."""
        files = filedialog.askopenfilenames(filetypes=[("PNG Image", "*.png")], title="Select Outline Masks")
        if not files: return
        for file_path in files:
            if file_path not in self.mask_paths:
                self.mask_paths.append(file_path)
                self.list_box.insert(tk.END, os.path.basename(file_path))
        self.generate_composite()

    def remove_mask_file(self):
        """Pops files from layout loops."""
        selected_idx = self.list_box.curselection()
        if not selected_idx: return
        idx = selected_idx[0]
        self.list_box.delete(idx)
        self.mask_paths.pop(idx)
        self.generate_composite()

    def clear_queue(self):
        """Resets panel data containers."""
        self.list_box.delete(0, tk.END)
        self.mask_paths.clear()
        self.annotations.clear()
        self.composite_pil = None
        self.canvas.delete("all")

    def generate_composite(self):
        """Layers elements sequentially, prioritizing microscope captures on the bottom layer."""
        if not self.mask_paths and self.microscope_image_path is None:
            self.composite_pil = None
            self.canvas.delete("all")
            return
        try:
            if self.microscope_image_path is not None:
                base_img = Image.open(self.microscope_image_path).convert("RGBA")
                w, h = base_img.size
                master_canvas = base_img
            else:
                first_mask = Image.open(self.mask_paths[0]).convert("RGBA")
                w, h = first_mask.size
                if self.bg_color_rgb is not None:
                    master_canvas = Image.new("RGBA", (w, h), self.bg_color_rgb + (255,))
                else:
                    master_canvas = Image.new("RGBA", (w, h), (0, 0, 0, 255))
            for path in self.mask_paths:
                layer = Image.open(path).convert("RGBA")
                if layer.size != (w, h):
                    layer = layer.resize((w, h), Image.Resampling.NEAREST)
                master_canvas.alpha_composite(layer)
            self.composite_pil = master_canvas
            self.render_preview()
        except Exception as e:
            messagebox.showerror("Blending Failure", f"Failed to layer masks:\n{str(e)}")

    def render_preview(self):
        """Updates layout preview states inside canvas boxes."""
        if self.composite_pil is None: return
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: canvas_w, canvas_h = 800, 600
        img_w, img_h = self.composite_pil.size
        scale = min(canvas_w / img_w, canvas_h / img_h)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))
        resized_img = self.composite_pil.resize((new_w, new_h), Image.Resampling.NEAREST)
        self.tk_img = ImageTk.PhotoImage(resized_img)
        self.canvas.delete("all")
        self.canvas.create_image(canvas_w // 2, canvas_h // 2, anchor=tk.CENTER, image=self.tk_img)
        for idx, anno in enumerate(self.annotations):
            cx, cy = self.image_to_canvas_coords(anno["x"], anno["y"])
            tk_color = '#%02x%02x%02x' % anno["color"]
            self.canvas.create_text(
                cx, cy, text=anno["text"], 
                fill=tk_color, font=("Arial", max(8, int(anno["size"] * scale)), "bold"),
                anchor=tk.CENTER
            )

    def export_merged_figure(self):
        """Saves everything out together, preserving underlying raw channel values."""
        if self.composite_pil is None:
            messagebox.showwarning("Empty Canvas", "Add layout sheets before exporting.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png")], title="Save Combined Overlay Figure")
        if not file_path: return
        try:
            w, h = self.composite_pil.size
            if self.microscope_image_path is not None:
                transparent_export = Image.open(self.microscope_image_path).convert("RGBA")
            elif self.bg_color_rgb is not None:
                transparent_export = Image.new("RGBA", (w, h), self.bg_color_rgb + (255,))
            else:
                transparent_export = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            for path in self.mask_paths:
                layer = Image.open(path).convert("RGBA")
                if layer.size != (w, h):
                    layer = layer.resize((w, h), Image.Resampling.NEAREST)
                transparent_export.alpha_composite(layer)
            draw = ImageDraw.Draw(transparent_export)
            for anno in self.annotations:
                try:
                    font = ImageFont.truetype("arial.ttf", anno["size"])
                except IOError:
                    font = ImageFont.load_default()
                text_str = anno["text"]
                if hasattr(draw, 'textbbox'):
                    bbox = draw.textbbox((0, 0), text_str, font=font)
                    tw = bbox[2] - bbox[0]
                    th = bbox[3] - bbox[1]
                else:
                    tw, th = draw.textsize(text_str, font=font)
                tx = anno["x"] - (tw // 2)
                ty = anno["y"] - (th // 2)
                draw.text((tx, ty), text_str, fill=anno["color"] + (255,), font=font)
            transparent_export.save(file_path, "PNG")
            messagebox.showinfo("Export Complete", f"Merged publication figure successfully written to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to render file vectors to disk:\n{str(e)}")
