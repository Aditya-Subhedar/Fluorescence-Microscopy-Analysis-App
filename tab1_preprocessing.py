import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cv2
import numpy as np
from PIL import Image, ImageTk
import os
import czifile
import tifffile

class PreProcessingTab(ttk.Frame):
    def __init__(self, parent, main_app):
        super().__init__(parent)
        self.main_app = main_app
        
        self.original_raw_volume = None # Keeps the uncropped backup
        self.raw_volume = None      
        self.original_filename = ""
        self.max_z = 0
        self.is_merged_preview = False
        self.channel_baselines = [] 
        
        # 16-bit Adjustments (Global Only)
        self.adj_settings = {'contrast': 1.0, 'brightness': 0.0}
        
        # Cropping variables
        self.rect_id = None
        self.start_x = None
        self.start_y = None
        self.current_rect = None
        self.img_offset_x = 0
        self.img_offset_y = 0
        self.img_scale = 1.0

        self.setup_ui()
        self.maximize_window()

    def maximize_window(self):
        top = self.winfo_toplevel()
        try:
            top.state('zoomed')  
        except Exception:
            try:
                top.attributes('-zoomed', True)  
            except Exception:
                top.attributes('-fullscreen', True) 

    def setup_ui(self):
        # Left Panel: Controls
        control_frame = tk.Frame(self, width=350, padx=10, pady=10)
        control_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # Action Frame
        action_frame = tk.Frame(control_frame)
        action_frame.pack(fill=tk.X, pady=(0, 5))

        # Changed text to imply multiple files can be loaded
        tk.Button(action_frame, text="1. Load CZI or Multi-Stack TIF(s)", command=self.load_czi, font=("Arial", 11, "bold"), bg="#4a90e2", fg="white").pack(fill=tk.X)
        
        self.lbl_filename = tk.Label(action_frame, text="No file loaded", fg="gray", wraplength=330)
        self.lbl_filename.pack(pady=(5, 0))

        # --- NEW: Multi-Image Navigation Bar ---
        self.nav_images_frame = tk.Frame(action_frame)
        self.nav_images_frame.pack(fill=tk.X, pady=5)
        
        self.btn_prev_img = tk.Button(self.nav_images_frame, text="◀ Prev", command=self.prev_image, state=tk.DISABLED, width=6)
        self.btn_prev_img.pack(side=tk.LEFT)
        
        self.lbl_img_count = tk.Label(self.nav_images_frame, text="0 / 0", font=("Arial", 10, "bold"))
        self.lbl_img_count.pack(side=tk.LEFT, expand=True)
        
        self.btn_next_img = tk.Button(self.nav_images_frame, text="Next ▶", command=self.next_image, state=tk.DISABLED, width=6)
        self.btn_next_img.pack(side=tk.RIGHT)
        # ---------------------------------------

        tk.Button(action_frame, text="2. Save Processed Image As...", command=self.save_image_to_disk, font=("Arial", 11, "bold"), bg="#2e7d32", fg="white", height=2).pack(fill=tk.X, pady=(5, 5))

        # Z-Navigation
        nav_frame = tk.LabelFrame(control_frame, text="Stack Preview Navigation", padx=10, pady=5)
        nav_frame.pack(fill=tk.X, pady=2)
        self.lbl_z_current = tk.Label(nav_frame, text="Current Stack: 0")
        self.lbl_z_current.pack()
        self.scale_z = tk.Scale(nav_frame, from_=0, to=0, orient=tk.HORIZONTAL, command=self.update_preview)
        self.scale_z.pack(fill=tk.X)

        # Merge Range (Z-Projection)
        proj_frame = tk.LabelFrame(control_frame, text="Z-Projection Merge Range", padx=10, pady=5)
        proj_frame.pack(fill=tk.X, pady=2)
        tk.Label(proj_frame, text="Start Stack:").grid(row=0, column=0, sticky="w", pady=2)
        self.spin_z_start = tk.Spinbox(proj_frame, from_=0, to=0, width=5, command=self.update_preview)
        self.spin_z_start.grid(row=0, column=1, pady=2)
        tk.Label(proj_frame, text="End Stack:").grid(row=1, column=0, sticky="w", pady=2)
        self.spin_z_end = tk.Spinbox(proj_frame, from_=0, to=0, width=5, command=self.update_preview)
        self.spin_z_end.grid(row=1, column=1, pady=2)
        self.btn_preview_merge = tk.Button(proj_frame, text="👁 Preview Merged Stacks", command=self.toggle_merge_preview)
        self.btn_preview_merge.grid(row=2, column=0, columnspan=2, pady=5, sticky="we")

        # Cropping Tool
        crop_frame = tk.LabelFrame(control_frame, text="Cropping Tool", padx=10, pady=5)
        crop_frame.pack(fill=tk.X, pady=2)
        tk.Label(crop_frame, text="Draw a rectangle on the image to crop.", fg="gray").pack(anchor="w", pady=(0,2))
        tk.Button(crop_frame, text="✂ Crop to Selection", command=self.apply_crop, bg="#f39c12", fg="white").pack(fill=tk.X, pady=2)
        tk.Button(crop_frame, text="🔄 Reset Original Image", command=self.reset_crop).pack(fill=tk.X, pady=2)

        # Channels
        chan_frame = tk.LabelFrame(control_frame, text="Channel Visibility", padx=10, pady=5)
        chan_frame.pack(fill=tk.X, pady=2)
        self.var_ch_r = tk.BooleanVar(value=True)
        self.var_ch_g = tk.BooleanVar(value=True)
        self.var_ch_b = tk.BooleanVar(value=True)
        tk.Checkbutton(chan_frame, text="Alexa Fluor 568 (Red)", variable=self.var_ch_r, fg="red", command=self.update_preview).pack(anchor="w")
        tk.Checkbutton(chan_frame, text="Alexa Fluor 488 (Green)", variable=self.var_ch_g, fg="green", command=self.update_preview).pack(anchor="w")
        tk.Checkbutton(chan_frame, text="DAPI (Blue)", variable=self.var_ch_b, fg="blue", command=self.update_preview).pack(anchor="w")

        # Adjustments
        adj_frame = tk.LabelFrame(control_frame, text="Image Adjustments", padx=10, pady=5)
        adj_frame.pack(fill=tk.X, pady=2)
        
        # Configure the frame to have two equal-width columns
        adj_frame.columnconfigure(0, weight=1)
        adj_frame.columnconfigure(1, weight=1)
        
        tk.Label(adj_frame, text="Contrast:").grid(row=0, column=0, sticky="w")
        # Reduced max to 5.0 for better slider precision in the 1-3 range
        self.scale_contrast = tk.Scale(adj_frame, from_=0.1, to=5.0, resolution=0.1, orient=tk.HORIZONTAL, command=self.on_slider_move)
        self.scale_contrast.set(1.0)
        self.scale_contrast.grid(row=1, column=0, sticky="ew", padx=(0, 5))
        
        tk.Label(adj_frame, text="Brightness:").grid(row=0, column=1, sticky="w")
        self.scale_brightness = tk.Scale(adj_frame, from_=-1.0, to=1.0, resolution=0.05, orient=tk.HORIZONTAL, command=self.on_slider_move)
        self.scale_brightness.set(0.0)
        self.scale_brightness.grid(row=1, column=1, sticky="ew", padx=(5, 0))

        # Right Panel: Canvas
        self.canvas_frame = tk.Frame(self, bg="black")
        self.canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.canvas = tk.Canvas(self.canvas_frame, bg="black", cursor="crosshair")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Mouse Bindings for Cropping
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_release)

        # --- NEW: Keyboard bindings for Left/Right arrows ---
        # Bound to top level window so they work regardless of where focus is
        top = self.winfo_toplevel()
        top.bind("<Left>", lambda e: self.prev_image() if self.btn_prev_img['state'] == tk.NORMAL else None)
        top.bind("<Right>", lambda e: self.next_image() if self.btn_next_img['state'] == tk.NORMAL else None)

    # --- Mouse Events for Cropping ---
    def on_mouse_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="yellow", dash=(4, 4), width=2)

    def on_mouse_drag(self, event):
        if self.rect_id:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_mouse_release(self, event):
        if self.rect_id:
            self.current_rect = (self.start_x, self.start_y, event.x, event.y)

    def apply_crop(self):
        if not self.current_rect or self.raw_volume is None: return
        x1, y1, x2, y2 = self.current_rect
        
        x_start, x_end = min(x1, x2), max(x1, x2)
        y_start, y_end = min(y1, y2), max(y1, y2)
        
        # Convert canvas coordinates to raw image indices
        img_x1 = int((x_start - self.img_offset_x) / self.img_scale)
        img_y1 = int((y_start - self.img_offset_y) / self.img_scale)
        img_x2 = int((x_end - self.img_offset_x) / self.img_scale)
        img_y2 = int((y_end - self.img_offset_y) / self.img_scale)
        
        # Constrain to image bounds
        h, w = self.raw_volume.shape[1:3]
        img_x1, img_x2 = max(0, min(img_x1, w)), max(0, min(img_x2, w))
        img_y1, img_y2 = max(0, min(img_y1, h)), max(0, min(img_y2, h))
        
        if img_x2 <= img_x1 or img_y2 <= img_y1:
            messagebox.showwarning("Crop Error", "Invalid crop area selected.")
            return
        
        # Slice the numpy array (applies to all Z-layers and Channels)
        self.raw_volume = self.raw_volume[:, img_y1:img_y2, img_x1:img_x2, :]
        
        self.current_rect = None
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
            
        self.update_preview()

    def reset_crop(self):
        if self.original_raw_volume is not None:
            self.raw_volume = self.original_raw_volume
            self.current_rect = None
            if self.rect_id:
                self.canvas.delete(self.rect_id)
                self.rect_id = None
            self.update_preview()

    # --- Adjustment Logic ---
    def on_slider_move(self, event=None):
        self.adj_settings['contrast'] = self.scale_contrast.get()
        self.adj_settings['brightness'] = self.scale_brightness.get()
        self.update_preview()

    def toggle_merge_preview(self):
        self.is_merged_preview = not self.is_merged_preview
        if self.is_merged_preview:
            self.btn_preview_merge.config(text="🔙 Back to Single Stack", bg="#e0e0e0")
            self.scale_z.config(state=tk.DISABLED) 
        else:
            self.btn_preview_merge.config(text="👁 Preview Merged Stacks", bg="SystemButtonFace")
            self.scale_z.config(state=tk.NORMAL)
        self.update_preview()

   # --- Loading ---
    def load_czi(self):
        # Use askopenfilenames (plural) to allow selecting multiple files
        file_paths = filedialog.askopenfilenames(
            title="Select CZI or Z-Stack Files", 
            filetypes=[("CZI/TIFF Files", "*.czi *.tif *.tiff")]
        )
        if not file_paths: return

        # Store the list of files and initialize the index
        self.loaded_files = list(file_paths)
        self.current_file_index = 0
        
        # Load the first image in the selection
        self.load_image_from_index()

    def prev_image(self):
        if hasattr(self, 'current_file_index') and self.current_file_index > 0:
            self.current_file_index -= 1
            self.load_image_from_index()

    def next_image(self):
        if hasattr(self, 'current_file_index') and self.current_file_index < len(self.loaded_files) - 1:
            self.current_file_index += 1
            self.load_image_from_index()

    def load_image_from_index(self):
        if hasattr(self, 'loaded_files') == False or not self.loaded_files:
            return

        file_path = self.loaded_files[self.current_file_index]

        # --- Update Navigation UI ---
        total = len(self.loaded_files)
        current = self.current_file_index + 1
        self.lbl_img_count.config(text=f"{current} / {total}")
        
        self.btn_prev_img.config(state=tk.NORMAL if self.current_file_index > 0 else tk.DISABLED)
        self.btn_next_img.config(state=tk.NORMAL if self.current_file_index < total - 1 else tk.DISABLED)

        # --- Begin Image Loading ---
        self.lbl_filename.config(text="Loading...")
        
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: canvas_w, canvas_h = 800, 600
        
        self.canvas.delete("all")
        self.canvas.create_text(
            canvas_w // 2, canvas_h // 2, 
            text=f"Please wait, loading image {current} of {total}...", 
            fill="white", font=("Arial", 16, "bold")
        )
        self.update() 
        
        try:
            self.original_filename = os.path.basename(file_path)
            
            if file_path.lower().endswith('.czi'):
                img = czifile.imread(file_path)
            else:
                img = tifffile.imread(file_path)
                
            # Safely strip 1-length dimensions
            img = np.squeeze(img)
            
            # --- ROBUST DIMENSION MAPPING (ImageJ Style) ---
            if img.ndim == 2:
                img = img[np.newaxis, ..., np.newaxis] 
            elif img.ndim == 3:
                if img.shape[2] <= 4:
                    img = img[np.newaxis, ...] 
                elif img.shape[0] <= 4:
                    img = np.moveaxis(img, 0, -1)[np.newaxis, ...]
                else:
                    img = img[..., np.newaxis]
            elif img.ndim == 4: 
                if img.shape[0] <= 4:
                    img = np.moveaxis(img, 0, -1)
                elif img.shape[1] <= 4:
                    img = np.moveaxis(img, 1, -1)
            
            if 0 in img.shape: raise ValueError(f"Empty dimensions: {img.shape}")
                
            # Track the REAL number of channels before we pad it
            self.original_num_channels = img.shape[-1]
                
            # Guarantee at least 3 channels for rendering logic mapping
            if img.shape[-1] == 1:
                img = np.concatenate([img]*3, axis=-1)
            elif img.shape[-1] == 2:
                img = np.concatenate([img, np.zeros_like(img[..., :1])], axis=-1)

            # Preserve memory: Original backup vs Active Working Volume
            self.original_raw_volume = img.astype(np.float32)
            self.raw_volume = self.original_raw_volume
            self.max_z = img.shape[0] - 1
            
            mid_z = self.max_z // 2
            self.channel_baselines = []
            
            # ImageJ-style Min/Max display mapping per channel
            for c in range(img.shape[-1]):
                slice_data = self.raw_volume[mid_z, ..., c]
                p_min, p_max = np.percentile(slice_data[slice_data > 0], (0.1, 99.9)) if np.any(slice_data > 0) else (0, 1)
                if p_max <= p_min: p_max = p_min + 1
                self.channel_baselines.append({'min': float(p_min), 'max': float(p_max)})

            self.scale_z.config(to=self.max_z)
            self.scale_z.set(mid_z)
            self.spin_z_start.config(to=self.max_z)
            self.spin_z_end.config(to=self.max_z)
            self.spin_z_start.delete(0, tk.END)
            self.spin_z_start.insert(0, 0)
            self.spin_z_end.delete(0, tk.END)
            self.spin_z_end.insert(0, str(self.max_z))
            
            self.adj_settings = {'contrast': 1.0, 'brightness': 0.0}
            self.scale_contrast.set(1.0)
            self.scale_brightness.set(0.0)
            
            self.lbl_filename.config(text=self.original_filename)
            self.canvas.after(100, self.update_preview)
            
        except Exception as e:
            self.lbl_filename.config(text="Load failed")
            self.canvas.delete("all")
            messagebox.showerror("Loading Error", f"Failed to parse volume:\n{e}")

    # --- Processing ---
    def apply_image_math(self, image_multi):
        h, w, c_total = image_multi.shape
        blended = np.zeros((h, w, 3), dtype=np.float32)
        
        # Look up the true, unpadded channel count from load_czi
        orig_c = getattr(self, 'original_num_channels', c_total)
        
        # --- DYNAMIC CHANNEL MAPPING ---
        if orig_c >= 3:
            # Standard 3-channel mapping: 0=DAPI, 1=Alexa488, 2=Alexa568
            channel_colors = [
                (0, 0, 255) if self.var_ch_b.get() else (0, 0, 0), # Ch 0 -> Blue
                (0, 255, 0) if self.var_ch_g.get() else (0, 0, 0), # Ch 1 -> Green
                (255, 0, 0) if self.var_ch_r.get() else (0, 0, 0)  # Ch 2 -> Red
            ]
        else:
            # Your specific 2-channel mapping: 0=Alexa488, 1=DAPI
            channel_colors = [
                (0, 255, 0) if self.var_ch_g.get() else (0, 0, 0), # Ch 0 -> Green
                (0, 0, 255) if self.var_ch_b.get() else (0, 0, 0), # Ch 1 -> Blue
                (255, 0, 0) if self.var_ch_r.get() else (0, 0, 0)  # Ch 2 -> Red
            ]
        
        # Additive Blending (Matches ImageJ "Composite" rendering)
        for i, color in enumerate(channel_colors):
            if color == (0, 0, 0) or i >= c_total or i >= len(self.channel_baselines): 
                continue
            
            ch_data = image_multi[:, :, i]
            
            b_min = self.channel_baselines[i]['min']
            b_max = self.channel_baselines[i]['max']
            val_range = (b_max - b_min) if (b_max - b_min) > 0 else 1.0
            
            # Normalize channel based on 16-bit limits
            norm_ch = (ch_data - b_min) / val_range
            norm_ch = np.clip(norm_ch, 0.0, 1.0)
                
            blended[:, :, 0] += norm_ch * color[0] 
            blended[:, :, 1] += norm_ch * color[1] 
            blended[:, :, 2] += norm_ch * color[2] 
            
        blended = np.clip(blended, 0, 255).astype(np.uint8)
        
        # Apply global contrast and brightness
        g_contrast = self.adj_settings['contrast']
        g_brightness = self.adj_settings['brightness']
        
        if g_contrast != 1.0 or g_brightness != 0.0:
            img_float = blended.astype(np.float32) / 255.0
            img_float = (img_float * g_contrast) + g_brightness
            img_float = np.clip(img_float, 0.0, 1.0)
            blended = (img_float * 255.0).astype(np.uint8)

        return blended

    def update_preview(self, event=None):
        if self.raw_volume is None: return
        
        if self.is_merged_preview:
            try:
                z_start = max(0, min(int(self.spin_z_start.get()), self.max_z))
                z_end = max(0, min(int(self.spin_z_end.get()), self.max_z))
                if z_start > z_end: z_start, z_end = z_end, z_start
            except ValueError:
                z_start, z_end = 0, self.max_z
                
            self.lbl_z_current.config(text=f"Previewing Merge: Stacks {z_start} to {z_end}")
            stack_slice = self.raw_volume[z_start:z_end+1]
            raw_data = np.max(stack_slice, axis=0) 
        else:
            z_idx = self.scale_z.get()
            self.lbl_z_current.config(text=f"Current Stack: {z_idx}")
            raw_data = self.raw_volume[z_idx]

        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: canvas_w, canvas_h = 800, 600
        
        img_h, img_w = raw_data.shape[:2]
        if img_w == 0 or img_h == 0: return 
        
        scale = min(canvas_w / img_w, canvas_h / img_h)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))
        
        # Save display parameters for cropping math
        self.img_scale = scale
        self.img_offset_x = (canvas_w - new_w) // 2
        self.img_offset_y = (canvas_h - new_h) // 2
        
        preview_small = cv2.resize(raw_data, (new_w, new_h), interpolation=cv2.INTER_NEAREST)
        preview_edited = self.apply_image_math(preview_small)
        
        self.tk_img = ImageTk.PhotoImage(Image.fromarray(preview_edited))
        self.canvas.delete("all")
        self.canvas.create_image(canvas_w//2, canvas_h//2, anchor=tk.CENTER, image=self.tk_img)
        self.rect_id = None # Clear old bounding box state on redraw

    # --- Saving ---
    def save_image_to_disk(self):
        if self.raw_volume is None: return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".tif",
            filetypes=[("TIFF File", "*.tif *.tiff"), ("JPEG File", "*.jpg *.jpeg"), ("PNG File", "*.png")],
            title="Save Processed Image As..."
        )
        if not file_path: return
        try:
            if self.is_merged_preview:
                z_start = max(0, min(int(self.spin_z_start.get()), self.max_z))
                z_end = max(0, min(int(self.spin_z_end.get()), self.max_z))
                if z_start > z_end: z_start, z_end = z_end, z_start
                stack_slice = self.raw_volume[z_start:z_end+1]
                target_data = np.max(stack_slice, axis=0) 
            else:
                target_data = self.raw_volume[self.scale_z.get()]
            
            final_rgb = self.apply_image_math(target_data)
            
            # Use final_rgb for TIFFs, convert to BGR for standard formats
            if file_path.lower().endswith(('.tif', '.tiff')):
                tifffile.imwrite(file_path, final_rgb) 
            else:
                final_bgr = cv2.cvtColor(final_rgb, cv2.COLOR_RGB2BGR)
                cv2.imwrite(file_path, final_bgr) 
                
            messagebox.showinfo("Success", f"Image saved successfully to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save image:\n{e}")