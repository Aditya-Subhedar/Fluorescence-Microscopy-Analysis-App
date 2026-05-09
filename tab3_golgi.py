import tkinter as tk
from tkinter import ttk, filedialog
import cv2
import numpy as np
from PIL import Image, ImageTk

class GolgiTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # State variables
        self.image_paths = []
        self.current_idx = 0
        
        # Image Pipelines
        self.raw_image = None       # 8-bit original
        self.processed_image = None # After Sharpness/Contrast
        self.stencil_image = None   # After Threshold
        
        self.soma_center = None     # (x, y) tuple in RAW IMAGE coordinates
        self.pixel_size_um = 1.0
        
        # Virtual Camera Variables (Synchronized Pan/Zoom)
        self.img_scale = 1.0
        self.img_tx = 0.0 # Translation X
        self.img_ty = 0.0 # Translation Y
        self.pan_start_x = 0
        self.pan_start_y = 0
        
        self.setup_ui()
        
    def setup_ui(self):
        """Builds the split layout for Golgi Spine Analysis."""
        paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        # ==========================================
        # LEFT PANEL: Controls & Stats
        # ==========================================
        control_frame = tk.Frame(paned_window, width=320, padx=10, pady=10)
        paned_window.add(control_frame, weight=0)
        
        btn_load = tk.Button(control_frame, text="Load TIFF Image(s)", bg="#4A90E2", fg="white", font=("Arial", 10, "bold"), command=self.load_images)
        btn_load.pack(fill=tk.X, pady=(0, 10))
        
        self.lbl_file_info = tk.Label(control_frame, text="No image loaded", fg="gray")
        self.lbl_file_info.pack(anchor="w")
        
        # 1. Calibration & Sholl Parameters
        sholl_frame = tk.LabelFrame(control_frame, text="Sholl & Calibration", padx=10, pady=10)
        sholl_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(sholl_frame, text="Pixel Size:").grid(row=0, column=0, sticky="w")
        self.lbl_px_size_display = tk.Label(sholl_frame, text="Waiting...", fg="blue", font=("Arial", 9, "italic"))
        self.lbl_px_size_display.grid(row=0, column=1, pady=2, sticky="w")
        
        tk.Label(sholl_frame, text="Circle Dist. (µm):").grid(row=1, column=0, sticky="w")
        self.entry_circle_dist = tk.Entry(sholl_frame, width=8)
        self.entry_circle_dist.insert(0, "10")
        self.entry_circle_dist.grid(row=1, column=1, pady=2, sticky="w")
        
        tk.Label(sholl_frame, text="Target Range (µm):").grid(row=2, column=0, sticky="w")
        range_frame = tk.Frame(sholl_frame)
        range_frame.grid(row=2, column=1, pady=2, sticky="w")
        self.entry_range_start = tk.Entry(range_frame, width=4)
        self.entry_range_start.insert(0, "20")
        self.entry_range_start.pack(side=tk.LEFT)
        tk.Label(range_frame, text="-").pack(side=tk.LEFT)
        self.entry_range_end = tk.Entry(range_frame, width=4)
        self.entry_range_end.insert(0, "30")
        self.entry_range_end.pack(side=tk.LEFT)
        
        tk.Button(sholl_frame, text="Apply Parameters", command=self.update_previews).grid(row=3, column=0, columnspan=2, pady=5, sticky="ew")

        # 2. Image Processing & Stencil
        proc_frame = tk.LabelFrame(control_frame, text="Image Adjustments & Stencil", padx=10, pady=10)
        proc_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(proc_frame, text="Contrast:").pack(anchor="w")
        self.scale_contrast = tk.Scale(proc_frame, from_=0.5, to=3.0, resolution=0.1, orient=tk.HORIZONTAL)
        self.scale_contrast.set(1.0)
        self.scale_contrast.pack(fill=tk.X)
        self.scale_contrast.bind("<ButtonRelease-1>", lambda e: self.apply_image_math())
        
        tk.Label(proc_frame, text="Sharpness:").pack(anchor="w")
        self.scale_sharpness = tk.Scale(proc_frame, from_=0.0, to=3.0, resolution=0.1, orient=tk.HORIZONTAL)
        self.scale_sharpness.set(0.0)
        self.scale_sharpness.pack(fill=tk.X)
        self.scale_sharpness.bind("<ButtonRelease-1>", lambda e: self.apply_image_math())
        
        tk.Label(proc_frame, text="Stencil Threshold (Black/White):").pack(anchor="w")
        self.scale_thresh = tk.Scale(proc_frame, from_=0, to=255, orient=tk.HORIZONTAL)
        self.scale_thresh.set(128)
        self.scale_thresh.pack(fill=tk.X)
        self.scale_thresh.bind("<ButtonRelease-1>", lambda e: self.process_stencil())

        # 3. Real-time Stats & Export
        stats_frame = tk.LabelFrame(control_frame, text="Live Quantification", padx=10, pady=10)
        stats_frame.pack(fill=tk.X, pady=10)
        
        self.lbl_total_spines = tk.Label(stats_frame, text="Total Spines: 0", font=("Arial", 10, "bold"))
        self.lbl_total_spines.pack(anchor="w", pady=2)
        
        self.lbl_range_spines = tk.Label(stats_frame, text="Spines in Range: 0", font=("Arial", 10))
        self.lbl_range_spines.pack(anchor="w", pady=2)
        
        self.lbl_density = tk.Label(stats_frame, text="Density: 0.0", font=("Arial", 10))
        self.lbl_density.pack(anchor="w", pady=2)
        
        btn_export = tk.Button(control_frame, text="Export Session to CSV/Excel", bg="#2E7D32", fg="white", font=("Arial", 10, "bold"))
        btn_export.pack(fill=tk.X, side=tk.BOTTOM, pady=10)

        # ==========================================
        # RIGHT PANEL: Dual Image Workspaces
        # ==========================================
        workspace_frame = ttk.PanedWindow(paned_window, orient=tk.HORIZONTAL)
        paned_window.add(workspace_frame, weight=1)
        
        frame_orig = tk.LabelFrame(workspace_frame, text="Original Image (Right-Click & Drag to Pan, Scroll to Zoom)")
        workspace_frame.add(frame_orig, weight=1)
        self.canvas_orig = tk.Canvas(frame_orig, bg="#1e1e1e", cursor="crosshair")
        self.canvas_orig.pack(fill=tk.BOTH, expand=True)
        
        frame_stencil = tk.LabelFrame(workspace_frame, text="High Contrast Stencil (Synced)")
        workspace_frame.add(frame_stencil, weight=1)
        self.canvas_stencil = tk.Canvas(frame_stencil, bg="#1e1e1e")
        self.canvas_stencil.pack(fill=tk.BOTH, expand=True)
        
        self.bind_canvas_events(self.canvas_orig)
        self.bind_canvas_events(self.canvas_stencil)
        
        # Initial fit on window resize
        self.canvas_orig.bind("<Configure>", self.on_canvas_resize)

    def bind_canvas_events(self, canvas):
        """Binds Pan, Zoom, and Soma selection events to a canvas."""
        canvas.bind("<Button-1>", self.set_soma_center)
        canvas.bind("<ButtonPress-3>", self.start_pan) # Right-click to pan
        canvas.bind("<B3-Motion>", self.do_pan)
        # Scroll wheel for zoom (Windows/macOS)
        canvas.bind("<MouseWheel>", self.do_zoom)
        # Scroll wheel for Linux
        canvas.bind("<Button-4>", self.do_zoom)
        canvas.bind("<Button-5>", self.do_zoom)

    # --- Virtual Camera: Pan & Zoom Logic ---
    def start_pan(self, event):
        self.pan_start_x = event.x
        self.pan_start_y = event.y

    def do_pan(self, event):
        dx = event.x - self.pan_start_x
        dy = event.y - self.pan_start_y
        self.img_tx += dx
        self.img_ty += dy
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        self.update_previews()

    def do_zoom(self, event):
        # Determine zoom direction
        if event.num == 4 or getattr(event, 'delta', 0) > 0:
            scale_factor = 1.1 # Zoom In
        else:
            scale_factor = 0.9 # Zoom Out

        new_scale = self.img_scale * scale_factor
        
        # Math to zoom specifically towards the mouse cursor
        self.img_tx = event.x - (event.x - self.img_tx) * scale_factor
        self.img_ty = event.y - (event.y - self.img_ty) * scale_factor
        self.img_scale = new_scale
        
        self.update_previews()

    def on_canvas_resize(self, event):
        """Fires when UI loads or user resizes the window. Fits image to screen."""
        if self.raw_image is not None and self.img_scale == 1.0:
            self.fit_image_to_window()

    def fit_image_to_window(self):
        """Calculates scale and offset to fit the entire image in the canvas."""
        if self.raw_image is None: return
        canvas_w = self.canvas_orig.winfo_width()
        canvas_h = self.canvas_orig.winfo_height()
        if canvas_w < 10: return # Canvas not rendered yet
        
        img_h, img_w = self.raw_image.shape[:2]
        self.img_scale = min(canvas_w / img_w, canvas_h / img_h)
        self.img_tx = (canvas_w - (img_w * self.img_scale)) / 2
        self.img_ty = (canvas_h - (img_h * self.img_scale)) / 2
        self.update_previews()

    # --- Image Loading and Processing ---
    def load_images(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("TIFF Images", "*.tif *.tiff")])
        if not filepaths: return
        self.image_paths = list(filepaths)
        self.current_idx = 0
        self.load_current_image()
        
    def load_current_image(self):
        path = self.image_paths[self.current_idx]
        self.lbl_file_info.config(text=f"Loaded: {path.split('/')[-1]}")
        
        raw = cv2.imread(path, cv2.IMREAD_UNCHANGED)
        if raw is None: return
        if len(raw.shape) == 3: raw = cv2.cvtColor(raw, cv2.COLOR_BGR2GRAY)
            
        if raw.dtype == np.uint16:
            self.raw_image = cv2.normalize(raw, None, 0, 255, cv2.NORM_MINMAX, dtype=cv2.CV_8U)
        else:
            self.raw_image = raw.astype(np.uint8)
            
        self.soma_center = None
        self.fit_image_to_window() 
        self.apply_image_math()

    def apply_image_math(self):
        """Applies Contrast and Sharpness to the raw image."""
        if self.raw_image is None: return
        
        contrast = self.scale_contrast.get()
        sharpness = self.scale_sharpness.get()
        
        # 1. Apply Contrast
        processed = cv2.convertScaleAbs(self.raw_image, alpha=contrast, beta=0)
        
        # 2. Apply Sharpness (Unsharp Mask)
        if sharpness > 0.0:
            blurred = cv2.GaussianBlur(processed, (0, 0), 3)
            processed = cv2.addWeighted(processed, 1.0 + sharpness, blurred, -sharpness, 0)
            
        self.processed_image = processed
        self.process_stencil() # Instantly update the stencil based on new sharpness

    def process_stencil(self):
        """Applies binary thresholding to create the high-contrast stencil."""
        if self.processed_image is None: return
        thresh_val = self.scale_thresh.get()
        
        # THRESH_BINARY ensures dark neurons remain dark on a white background
        _, self.stencil_image = cv2.threshold(self.processed_image, thresh_val, 255, cv2.THRESH_BINARY)
        self.update_previews()
        
    def set_soma_center(self, event):
        """Captures mouse click and maps canvas coordinates back to raw image coordinates."""
        if self.raw_image is None: return
        raw_x = (event.x - self.img_tx) / self.img_scale
        raw_y = (event.y - self.img_ty) / self.img_scale
        self.soma_center = (int(raw_x), int(raw_y))
        self.update_previews()

    # --- High Performance Synchronized Rendering ---
    def get_render_crop(self, source_img, canvas):
        """Crops and resizes only the visible portion of the image for massive performance gains."""
        canvas_w = canvas.winfo_width()
        canvas_h = canvas.winfo_height()
        
        inv_scale = 1.0 / self.img_scale
        src_x1 = int(-self.img_tx * inv_scale)
        src_y1 = int(-self.img_ty * inv_scale)
        src_x2 = int((canvas_w - self.img_tx) * inv_scale)
        src_y2 = int((canvas_h - self.img_ty) * inv_scale)
        
        img_h, img_w = source_img.shape[:2]
        crop_x1, crop_y1 = max(0, src_x1), max(0, src_y1)
        crop_x2, crop_y2 = min(img_w, src_x2), min(img_h, src_y2)
        
        dst_w = int((crop_x2 - crop_x1) * self.img_scale)
        dst_h = int((crop_y2 - crop_y1) * self.img_scale)
        
        if dst_w > 0 and dst_h > 0:
            crop_img = source_img[crop_y1:crop_y2, crop_x1:crop_x2]
            resized = cv2.resize(crop_img, (dst_w, dst_h), interpolation=cv2.INTER_LINEAR)
            place_x = crop_x1 * self.img_scale + self.img_tx
            place_y = crop_y1 * self.img_scale + self.img_ty
            return resized, place_x, place_y
        return None, 0, 0

    def update_previews(self):
        """Renders both synced canvases dynamically."""
        if self.processed_image is None or self.stencil_image is None: return
        
        # 1. Render Left Canvas (Processed Original)
        res_orig, px_o, py_o = self.get_render_crop(self.processed_image, self.canvas_orig)
        if res_orig is not None:
            img_o_pil = Image.fromarray(res_orig)
            self.tk_img_original = ImageTk.PhotoImage(img_o_pil)
            self.canvas_orig.delete("all")
            self.canvas_orig.create_image(px_o, py_o, anchor="nw", image=self.tk_img_original)
        
        # 2. Prepare Right Canvas (Stencil + Drawings)
        stencil_color = cv2.cvtColor(self.stencil_image, cv2.COLOR_GRAY2BGR)
        
        # Draw Sholl rings in RAW coordinates (crop logic handles zoom naturally)
        if self.soma_center:
            try:
                dist_um = float(self.entry_circle_dist.get())
                radius_step_px = int(dist_um / self.pixel_size_um)
                
                if radius_step_px > 0:
                    max_dim = max(stencil_color.shape[0], stencil_color.shape[1])
                    num_circles = int(max_dim / radius_step_px) + 2
                    cx, cy = self.soma_center
                    
                    for i in range(1, num_circles):
                        r = i * radius_step_px
                        cv2.circle(stencil_color, (cx, cy), r, (255, 0, 0), 2)
                        cv2.putText(stencil_color, f"{int(i * dist_um)}um", (cx + r + 2, cy), 
                                    cv2.FONT_HERSHEY_SIMPLEX, 1.0, (255, 0, 0), 2)
                        
                    cv2.circle(stencil_color, (cx, cy), 8, (0, 0, 255), -1)
            except ValueError:
                pass 
                
        # Crop/Zoom the stencil to exactly match the original
        res_stencil, px_s, py_s = self.get_render_crop(stencil_color, self.canvas_stencil)
        if res_stencil is not None:
            img_s_pil = Image.fromarray(cv2.cvtColor(res_stencil, cv2.COLOR_BGR2RGB))
            self.tk_img_stencil = ImageTk.PhotoImage(img_s_pil)
            self.canvas_stencil.delete("all")
            self.canvas_stencil.create_image(px_s, py_s, anchor="nw", image=self.tk_img_stencil)
            
        # Update labels dynamically
        r_start, r_end = self.entry_range_start.get(), self.entry_range_end.get()
        self.lbl_range_spines.config(text=f"Spines in Range ({r_start}-{r_end}µm): 0")