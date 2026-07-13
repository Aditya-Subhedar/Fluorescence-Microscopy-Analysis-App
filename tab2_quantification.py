import os
import cv2
import json
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import tifffile
import czifile
from skimage import filters, measure
# --> Import custom widget from widgets.pyk
from widgets import ColorRangeSlider, SingleSlider
from tkinter import colorchooser
import threading
from concurrent.futures import ThreadPoolExecutor


class QuantificationTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.image_states = [] 
        self.current_index = 0
        
        self.original_image_rgb = None
        self.current_manual_add = None
        self.current_manual_remove = None
        self.current_mask = None # <-- NEW: Master binary mask containing sliders + pencil

        self.cached_hsv = None
        self.cached_gray = None
        self._update_job = None 
        
        self.auto_detect_enabled = False 
        self._ignore_sliders = False 
        
        self.draw_mode = "pencil"
        self.is_processing = False 
        self.is_drawing = False
        self.last_x = 0
        self.last_y = 0
        self.scale_x = 1.0
        self.scale_y = 1.0
        self.offset_x = 0
        self.offset_y = 0

        self.draw_mode = None          # Tracks "pencil", "eraser", or "circle"
        self.active_oval = None        # Stores raw pixel coords: (center_x, center_y, radius_x, radius_y)
        self.temp_oval_id = None       # Canvas item ID for dynamic dragging preview
        self.handle_ids = []           # Canvas item IDs for active bounding box adjusters
        self.active_handle = None      # Tracks which handle is currently dragged ("T", "B", "L", "R")

        self.mask_override_mode = False
        self.override_mask_data = None

        # --- PRESET STATE VARIABLES ---
        self.presets_file = "cytoquant_presets.json"
        self.presets_collection = {} 
        self.pinned_presets = []
        self.current_preset = None 
        
        # Load presets from file if it exists
        self.load_presets_from_file()

        # --- Caching Adjscent Images for smoother switching ----
        self.image_cache = {}  # Format: {file_path: numpy_array}
        self.cache_executor = ThreadPoolExecutor(max_workers=1)  # Dedicated background thread
        self.cache_lock = threading.Lock()

        self.setup_ui()

    def setup_ui(self):
        root_frame = tk.Frame(self)
        root_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5) # Reduced padding to maximize space
        
        # --- Control Bar ---
        control_frame = tk.Frame(root_frame, pady=5)
        control_frame.pack(fill=tk.X)
        
        # 1. Base Operations
        self.btn_select_images = tk.Button(control_frame, text="1. Select Images", command=self.load_files, font=("Arial", 9, "bold"))
        self.btn_select_images.pack(side=tk.LEFT, padx=3)
        
        self.btn_auto = tk.Button(control_frame, text="Detect: OFF", command=self.toggle_auto_detect, fg="red", font=("Arial", 9, "bold"))
        self.btn_auto.pack(side=tk.LEFT, padx=5)

        # ---> NEW: ROI Number Toggle <---
        self.show_roi_numbers = True # Default state
        self.btn_toggle_nums = tk.Button(control_frame, text="# Labels: ON", command=self.toggle_roi_numbers, fg="green", font=("Arial", 9, "bold"))
        self.btn_toggle_nums.pack(side=tk.LEFT, padx=2)
        
        # 2. Drawing Tools Frame
        tool_frame = tk.Frame(control_frame, bd=1, relief=tk.SOLID, padx=3, pady=2)
        tool_frame.pack(side=tk.LEFT, padx=4)
        tk.Label(tool_frame, text="Tools:", font=("Arial", 8)).pack(side=tk.LEFT, padx=1)
        
        self.btn_pencil = tk.Button(tool_frame, text="✏️", relief=tk.RAISED, command=lambda: self.set_draw_mode("pencil"))
        self.btn_pencil.pack(side=tk.LEFT, padx=1)

        self.btn_circle = tk.Button(tool_frame, text="⭕", width=2, command=lambda: self.set_draw_mode("circle"), font=("Arial", 9, "bold"))
        self.btn_circle.pack(side=tk.LEFT, padx=1)
        
        self.btn_eraser = tk.Button(tool_frame, text="🧹", relief=tk.RAISED, command=lambda: self.set_draw_mode("eraser"))
        self.btn_eraser.pack(side=tk.LEFT, padx=1)

        self.btn_undo = tk.Button(tool_frame, text="↩️", command=self.undo_action, font=("Arial", 9)) # Compact icon only
        self.btn_undo.pack(side=tk.LEFT, padx=1)

        self.btn_redo = tk.Button(tool_frame, text="↪️", command=self.redo_action, font=("Arial", 9)) # Compact icon only
        self.btn_redo.pack(side=tk.LEFT, padx=1)
        
        tk.Button(tool_frame, text="Clear All", command=self.clear_drawing, fg="red", font=("Arial", 9)).pack(side=tk.LEFT, padx=3)
        
        # 3. Parameters/Presets Frame (Shortened text labels to preserve space)
        preset_frame = tk.Frame(control_frame, bd=1, relief=tk.SOLID, padx=3, pady=2)
        preset_frame.pack(side=tk.LEFT, padx=4)
        tk.Label(preset_frame, text="Preset:", font=("Arial", 8)).pack(side=tk.LEFT, padx=1)
        
        tk.Button(preset_frame, text="Save Parameters", command=self.save_as_preset, font=("Arial", 9)).pack(side=tk.LEFT, padx=1)

        self.btn_apply_preset = tk.Button(preset_frame, text="Apply Preset", command=self.show_preset_dropdown, font=("Arial", 9))
        self.btn_apply_preset.pack(side=tk.LEFT, padx=1)

        # 4. ---> NEW: DATA & IMAGE EXPORTS FRAME <---
        export_frame = tk.Frame(control_frame, bd=1, relief=tk.SOLID, padx=3, pady=2)
        export_frame.pack(side=tk.RIGHT, padx=4)
        
        # Export Image Button
        tk.Button(export_frame, text="🖼️ Export Image", command=self.export_current_image_view, 
                  font=("Arial", 9, "bold"), fg="white", bg="#0288d1").pack(side=tk.LEFT, padx=1)
        
        # Export Data Button
        tk.Button(export_frame, text="📊 Export Data", command=self.export_excel, 
                  font=("Arial", 9, "bold"), fg="white", bg="#2e7d32").pack(side=tk.LEFT, padx=1)

        # 5. Mask Import / Export Frame
        mask_io_frame = tk.Frame(control_frame, bd=1, relief=tk.SOLID, padx=3, pady=2)
        mask_io_frame.pack(side=tk.RIGHT, padx=4)
        
        tk.Button(mask_io_frame, text="💾 Save Mask", command=self.save_mask_as_png, 
                  bg="#1976d2", fg="white", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=1)
        
        self.btn_apply_mask = tk.Button(
            mask_io_frame, text="📥 Apply Mask", command=self.apply_saved_mask, 
            bg="#ef6c00", fg="white", font=("Arial", 9, "bold")
        )
        self.btn_apply_mask.pack(side=tk.LEFT, padx=1)

        

        # --- SLIDER FRAME ---
        slider_frame = tk.Frame(root_frame, pady=10)
        slider_frame.pack(fill=tk.X, padx=10)

        # 1. Hue Range 
        hue_frame = tk.Frame(slider_frame)
        hue_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Label(hue_frame, text="Color Filter (Hue):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.hue_slider = ColorRangeSlider(hue_frame, width=220, height=25, slider_type="hue", abs_min=0, abs_max=179, command=self.schedule_update) 
        self.hue_slider.pack(fill=tk.X, pady=5)

        # 2. Intensity Range 
        int_frame = tk.Frame(slider_frame)
        int_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Label(int_frame, text="Intensity (0=Black, 255=White):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.int_slider = ColorRangeSlider(int_frame, width=220, height=25, slider_type="intensity", abs_min=0, abs_max=255, command=self.schedule_update)
        self.int_slider.pack(fill=tk.X, pady=5)

        # 3. Area Range
        area_frame = tk.Frame(slider_frame)
        area_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Label(area_frame, text="Area Filter (px):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.area_slider = ColorRangeSlider(area_frame, width=220, height=25, slider_type="area", abs_min=0, abs_max=1000, command=self.schedule_update)
        self.area_slider.pack(fill=tk.X, pady=5)
        
        # 4. Circularity / Split (Updated to Single Slider)
        circ_frame = tk.Frame(slider_frame)
        circ_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Updated Label to reflect the new range to the user
        tk.Label(circ_frame, text="Morphology (-100=Line, 0=All, 100=Circle):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        # Updated abs_min to -100. 
        self.circ_slider = SingleSlider(circ_frame, width=220, height=25, abs_min=-100, abs_max=100, default_value=0, command=self.schedule_update)
        self.circ_slider.pack(fill=tk.X, pady=5)

        # Canvas
        self.canvas_frame = tk.Frame(root_frame, bg="black")
        self.canvas_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)
        
        self.canvas = tk.Canvas(self.canvas_frame, bg="black", cursor="crosshair")
        self.canvas.pack(expand=True, fill=tk.BOTH)
        
        self.canvas.bind("<Button-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_motion)
        self.canvas.bind("<ButtonRelease-1>", self.stop_draw)

        # --- Trackpad / Mousewheel Bindings ---
        # 1. CTRL + Swipe (or Ctrl+Scroll) to ZOOM
        self.canvas.bind("<Control-MouseWheel>", self.on_mousewheel_zoom) 
        self.canvas.bind("<Double-Button-1>", self.on_double_click)
        
        # 2. Normal Two-Finger Swipe to PAN (Move around)
        self.canvas.bind("<MouseWheel>", self.on_trackpad_scroll_y)       # Vertical swipe
        self.canvas.bind("<Shift-MouseWheel>", self.on_trackpad_scroll_x) # Horizontal swipe
        
        # (Linux support if needed)
        self.canvas.bind("<Control-Button-4>", self.on_mousewheel_zoom)
        self.canvas.bind("<Control-Button-5>", self.on_mousewheel_zoom)
        self.canvas.bind("<Button-4>", self.on_trackpad_scroll_y)
        self.canvas.bind("<Button-5>", self.on_trackpad_scroll_y)
        
        self.zoom_factor = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0
        # -------------------------------------------------------

        # --- Nav (UPDATED WITH VARIABLE POINTERS AND STATUS LABELS) ---
        nav_frame = tk.Frame(root_frame, pady=10)
        nav_frame.pack(fill=tk.X, padx=10)
        
        # FIXED: Saved explicitly to instance variable self.btn_prev_img
        self.btn_prev_img = tk.Button(nav_frame, text="<< Prev", command=self.prev_image, font=("Arial", 10))
        self.btn_prev_img.pack(side=tk.LEFT)
        
        stats_frame = tk.Frame(nav_frame)
        stats_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        self.lbl_stats_integrated = tk.Label(stats_frame, text="", font=("Arial", 11, "bold"))
        self.lbl_stats_integrated.pack()
        
        # FIXED: Saved explicitly to instance variable self.btn_next_img
        self.btn_next_img = tk.Button(nav_frame, text="Next >>", command=self.next_image, font=("Arial", 10))
        self.btn_next_img.pack(side=tk.RIGHT)

        # -------------------------------------------------------

    # --- Loadng ---
    def load_files(self):
        """Standalone loader for Tab 2: Disconnects from Tab 1 and loads files directly."""
        filetypes = [
            ("All Supported Images", "*.tif *.tiff *.jpg *.jpeg *.jfif *.png *.czi *.JPG *.JPEG *.PNG"),
            ("JPEG Images", "*.jpg *.jpeg *.jfif *.JPG *.JPEG"),
            ("PNG Images", "*.png *.PNG"),
            ("Scientific Images", "*.tif *.tiff *.czi"),
            ("All Files", "*.*")
        ]
            
        files = filedialog.askopenfilenames(title="Select Images for Analysis", filetypes=filetypes)
        if not files: 
            return
                
        self.image_files = sorted(list(files))
        self.current_index = 0
        self.image_states = []
            
        # --- NEW: Clear the cache for the new file pool ---
        with self.cache_lock:
            self.image_cache.clear()
            
        for file_path in self.image_files:
            self.image_states.append({
                'file_path': file_path,
                'hue_min': 0, 
                'hue_max': 179,
                'int_min': 0,
                'int_max': 255,
                'area_min_pos': 30,
                'area_max_pos': 1000,
                'manual_mask_add': None, 
                'manual_mask_remove': None,
                'undo_stack': [], 
                'redo_stack': []  
            })
            
        # Trigger loading the first image
        self.load_current_image_data()
        self.update_nav_button_states()

    def get_image_from_cache(self, path):
        """Fetches an image from memory if cached; otherwise loads it synchronously."""
        with self.cache_lock:
            if path in self.image_cache:
                return self.image_cache[path]
            
        # Fallback if not cached yet
        img = self.load_raw_image_array(path)
        if img is not None:
            with self.cache_lock:
                self.image_cache[path] = img
        return img

    def manage_cache_pipeline(self):
        """Asynchronously manages memory: keeps adjacent images in RAM and drops old ones."""
        if not hasattr(self, 'image_files') or not self.image_files:
            return

        total_files = len(self.image_files)
        idx = self.current_index

        # Define targets to keep: current image, next 2 images, previous 1 image
        targets = set()
        for offset in [0, 1, 2, -1]:
            target_idx = idx + offset
            if 0 <= target_idx < total_files:
                targets.add(self.image_files[target_idx])

        # Background thread task
        def cache_worker():
            # 1. Prune stale cache entries outside our window to save memory
            with self.cache_lock:
                stale_paths = [p for p in self.image_cache if p not in targets]
                for p in stale_paths:
                    del self.image_cache[p]

            # 2. Prefetch required adjacent images into RAM
            for path in targets:
                with self.cache_lock:
                    already_cached = path in self.image_cache
                    
                if not already_cached:
                    img = self.load_raw_image_array(path)
                    if img is not None:
                        with self.cache_lock:
                            self.image_cache[path] = img

        # Dispatch worker task to the background executor thread
        self.cache_executor.submit(cache_worker)

    def load_raw_image_array(self, path):
        """Reads a 2D image from disk. If a 3D TIFF is detected, safely flattens it via MIP."""
        try:
            import tifffile
            import numpy as np
            import cv2
            from PIL import Image

            if path.lower().endswith(('.tif', '.tiff')):
                # Use tifffile to read the complete TIFF data array structure
                img = tifffile.imread(path)
                
                # --- SAFETY FALLBACK: Handle accidental 3D TIFF Inputs gracefully ---
                # Check if the array has 3+ dimensions (e.g., [Slices, Height, Width] or [Slices, Height, Width, Channels])
                if img.ndim >= 3 and img.shape[0] > 4 and img.shape[-1] != 3:
                    # Treat the first axis as the Z-stack and generate a flat 2D Projection instantly
                    print(f"Warning: 3D Multi-page TIFF detected at {os.path.basename(path)}. Flattening via MIP.")
                    img = np.max(img, axis=0)
                elif img.ndim == 3 and img.shape[0] <= 4:
                    # This is just a standard 2D image with 3 or 4 color channels, reshape to HWC if needed
                    pass
            elif path.lower().endswith('.czi'):
                # Fallback support if CZI images are selected directly
                import czifile
                img = czifile.imread(path)
                img = np.squeeze(img)
                if img.ndim == 3 and img.shape[0] < 10: 
                    img = np.transpose(img, (1, 2, 0))
                if img.ndim > 3:
                    # Flatten the raw CZI volume into a 2D maximum projection
                    img = np.max(img, axis=0) 
            else:
                # Use Pillow for standard formats (JPEG, PNG)
                pil_img = Image.open(path).convert('RGB')
                img = np.array(pil_img)
                    
            if img is None: return None
            
            # --- FIX: Only normalize raw scientific formats. Leave Tab 1 standard images alone! ---
            if img.dtype == np.uint16 or img.dtype == np.float32 or img.dtype == np.int32:
                img = cv2.normalize(img, None, 0, 255, cv2.NORM_MINMAX)
                img = np.uint8(img)
            elif img.dtype != np.uint8:
                img = np.uint8(img) # Simple type cast without crushing contrast
                
            # If loaded with an alpha channel (RGBA), drop it to preserve crisp RGB details
            if len(img.shape) == 3 and img.shape[2] == 4:
                img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
            
            # Handle grayscale images safely
            if len(img.shape) == 2:
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
                
            return img
        
        except Exception as e:
            import threading
            if threading.current_thread() == threading.main_thread():
                import tkinter.messagebox as messagebox
                messagebox.showerror("Codec Error", f"Failed reading base image data:\n{e}")
            else:
                print(f"Background Cache Thread Warning: Failed reading {path}: {e}")
            return None
        
    def get_pixel_size_um(self, file_path):
        """Extracts the physical pixel size from the 2D TIFF exported by Tab 1."""
        try:
            from PIL import Image
            with Image.open(file_path) as img:
                # 1. Try reading the high-fidelity standard TIFF metadata tags
                if hasattr(img, 'tag_v2'):
                    x_res = img.tag_v2.get(282)  # Tag 282 = XResolution
                    unit = img.tag_v2.get(296)   # Tag 296 = ResolutionUnit
                    
                    if x_res and unit:
                        # Extract the primitive values regardless of if it is an IFDRational object or a tuple
                        if hasattr(x_res, 'numerator') and hasattr(x_res, 'denominator'):
                            num, den = x_res.numerator, x_res.denominator
                        elif isinstance(x_res, (tuple, list)):
                            num, den = x_res[0], x_res[1] if len(x_res) > 1 else 1
                        else:
                            num, den = float(x_res), 1

                        if num > 0 and den > 0:
                            pixels_per_unit = num / den
                            # Unit 3 = Centimeters (1 cm = 10,000 um)
                            if unit == 3: 
                                return 10000.0 / pixels_per_unit
                            # Unit 2 = Inches (1 inch = 25,400 um)
                            elif unit == 2: 
                                return 25400.0 / pixels_per_unit
                
                # 2. Backup Fallback: Try reading ImageJ or OME-TIFF metadata strings inside ImageDescription
                if 'ImageDescription' in img.info:
                    desc = str(img.info['ImageDescription'])
                    import re
                    
                    match_ome = re.search(r'PhysicalSizeX="([0-9.]+)"', desc)
                    if match_ome:
                        return float(match_ome.group(1))
                        
                    if 'unit=micron' in desc or 'unit=um' in desc:
                        match_ij = re.search(r'spacing=([0-9.]+)', desc)
                        if match_ij: 
                            return float(match_ij.group(1))

        except Exception as e:
            print(f"Metadata extraction fell back: {e}")
            
        return None

    def load_current_image_data(self):
        """Loads the current selected image into the UI using high-speed look-ahead memory extraction."""
        if self.current_index >= len(self.image_files) or not self.image_states: return
        
        state = self.image_states[self.current_index]
        file_path = state['file_path']
        
        try:
            # --- THE FIX: Pull from fast look-ahead RAM cache instead of slow disk operations ---
            self.original_image_rgb = self.get_image_from_cache(file_path)
            
            if self.original_image_rgb is None: return

            self.zoom_factor = 1.0
            self.pan_x = 0
            self.pan_y = 0
            
            # Grab metadata
            self.pixel_size_um = self.get_pixel_size_um(file_path)

            if state.get('manual_mask_add') is None:
                state['manual_mask_add'] = np.zeros(self.original_image_rgb.shape[:2], dtype=np.uint8)
            if state.get('manual_mask_remove') is None:
                state['manual_mask_remove'] = np.zeros(self.original_image_rgb.shape[:2], dtype=np.uint8)
                
            self.current_manual_add = state['manual_mask_add']
            self.current_manual_remove = state['manual_mask_remove']
            
            img_blur = cv2.GaussianBlur(self.original_image_rgb, (3, 3), 0)
            self.cached_hsv = cv2.cvtColor(img_blur, cv2.COLOR_RGB2HSV)
            self.cached_gray = cv2.cvtColor(img_blur, cv2.COLOR_RGB2GRAY)
            
            self.auto_detect_enabled = False
            self.btn_auto.config(text="Auto Detect: OFF", fg="red")
            
            self._ignore_sliders = True 
            
            # INTENSITY DUAL SLIDER
            if state.get('int_min') is None or state.get('int_min') == 0:
                otsu_val = filters.threshold_otsu(self.cached_gray)
                state['int_min'] = int(otsu_val)
                state['int_max'] = 255 
                
            self.int_slider.set_values(state.get('int_min', 0), state.get('int_max', 255))

            # HUE SLIDER
            self.hue_slider.set_values(state.get('hue_min', 0), state.get('hue_max', 179))
            
            # AREA DUAL SLIDER
            if 'area_min_pos' not in state: state['area_min_pos'] = 30
            if 'area_max_pos' not in state: state['area_max_pos'] = 1000
            
            self.area_slider.set_values(state['area_min_pos'], state['area_max_pos'])

            # CIRCULARITY SLIDER
            if 'circ_min' not in state: state['circ_min'] = 0
            if 'circ_max' not in state: state['circ_max'] = 100
            
            self.circ_slider.set_values(state.get('circ_min', 0))

            self._ignore_sliders = False 

            self.process_image()
            
            # --- NEW: Automatically start fetching adjacent images in the background ---
            self.manage_cache_pipeline()
            
        except Exception as e:
            import traceback
            traceback.print_exc()

    # --- Mouse Events for Zooming and Scrooling ---
    def on_mousewheel_zoom(self, event):
        # 1. Determine direction
        if hasattr(event, 'num') and event.num == 4:
            scale_change = 1.05
        elif hasattr(event, 'num') and event.num == 5:
            scale_change = 0.95
        elif event.delta > 0:
            scale_change = 1.05
        else:
            scale_change = 0.95

        cx, cy = event.x, event.y

        true_x = (cx - getattr(self, 'pan_x', 0)) / getattr(self, 'zoom_factor', 1.0)
        true_y = (cy - getattr(self, 'pan_y', 0)) / getattr(self, 'zoom_factor', 1.0)

        new_zoom = getattr(self, 'zoom_factor', 1.0) * scale_change
        
        # --- THE FIX: Change minimum bound from 0.1 to 1.0 ---
        if scale_change < 1.0 and new_zoom <= 1.01:
            # If zooming out drops to or below 1.0, automatically snap and center
            self.reset_and_center_view()
            return
        
        # Enforce constraints (Hard 1.0x minimum, 25.0x maximum zoom)
        new_zoom = max(1.0, min(new_zoom, 25.0))
        self.zoom_factor = new_zoom

        self.pan_x = cx - (true_x * self.zoom_factor)
        self.pan_y = cy - (true_y * self.zoom_factor)

        # REDRAW INSTANTLY USING CACHED IMAGE
        self.fast_redraw()

    def reset_and_center_view(self, event=None):
        """Instantly resets zoom factor to 1.0x and centers the image safely inside the canvas."""
        if not hasattr(self, 'original_image_rgb') or self.original_image_rgb is None:
            return

        # Force Tkinter geometry updates to establish real canvas view space coordinates
        self.canvas.update_idletasks()

        # 1. Get true viewable viewport bounds (ignoring hidden overflow or outer frame borders)
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        # Fallback if UI layout thread hasn't assigned space parameters yet
        if canvas_w < 10 or canvas_h < 10:
            try:
                canvas_w = int(self.canvas.cget("width"))
                canvas_h = int(self.canvas.cget("height"))
            except Exception:
                canvas_w, canvas_h = 800, 600

        # 2. Extract true original image width and height dimensions
        img_h, img_w = self.original_image_rgb.shape[:2]
        
        # 3. Lock zoom factor tightly to baseline 1.0x
        self.zoom_factor = 1.0
        
        # 4. FIXED CENTERING: Apply explicit padding insulation subtraction to clear corner confusion
        # This aligns the true center point of the image matrix with the viewport midpoint
        self.pan_x = int((canvas_w - img_w) / 2.0)
        self.pan_y = int((canvas_h - img_h) / 2.0)
        
        # Prevent any single-pixel clipping drifts by checking hard absolute minimum limits
        if canvas_w <= img_w: self.pan_x = 0
        if canvas_h <= img_h: self.pan_y = 0
        
        # Trigger layout refresh loop to draw the newly aligned frame positions
        self.fast_redraw()

    def on_double_click(self, event):
        """Triggers when the user double-clicks to instantly unzoom to 1.0x and center."""
        # Clean any phantom clicks or coordinates and center layout space
        self.reset_and_center_view()

    def on_trackpad_scroll_y(self, event):
        # Explicitly check for Linux scroll buttons, otherwise use Mac/Windows delta
        if hasattr(event, 'num') and event.num in (4, 5):
            delta = 10 if event.num == 4 else -10
        else:
            delta = event.delta
            # Windows sends massive deltas (multiples of 120). Scale them down to match Mac trackpads.
            if abs(delta) >= 120: 
                delta = delta / 12 
        
        self.pan_y += delta
        self.fast_redraw()

    def on_trackpad_scroll_x(self, event):
        # X-scrolling (Shift + Scroll)
        if hasattr(event, 'num') and event.num in (4, 5):
            delta = 10 if event.num == 4 else -10
        else:
            delta = event.delta
            if abs(delta) >= 120: 
                delta = delta / 12 
        
        self.pan_x += delta
        self.fast_redraw()

    # --- Auto Detect for Segmentation of Fluorosent Regions ---
    def toggle_auto_detect(self):
        # Clear the override mode if they turn auto detect back on
        self.mask_override_mode = False
        self.override_mask_data = None
        
        # Toggle your baseline auto detect tracking boolean
        self.auto_detect_enabled = not self.auto_detect_enabled
        
        if self.auto_detect_enabled:
            self.btn_auto.config(text="Auto Detect: ON", fg="white")
        else:
            self.btn_auto.config(text="Auto Detect: OFF", fg="red")
            
        self.process_image()

    def toggle_roi_numbers(self):
        """Toggles the visibility of ROI indices on the preview canvas."""
        self.show_roi_numbers = not getattr(self, 'show_roi_numbers', True)
        
        if self.show_roi_numbers:
            self.btn_toggle_nums.config(text="# Labels: ON", fg="green")
        else:
            self.btn_toggle_nums.config(text="# Labels: OFF", fg="gray")
            
        self.schedule_update() # Redraw the canvas to apply the change

    # --- Slider Adjustments and Preview Updates ---
    def on_slider_move_continuous(self, val):
        self.schedule_update()

    def schedule_update(self):
        if self._ignore_sliders: return 
        if self._update_job is not None:
            self.after_cancel(self._update_job)
        self._update_job = self.after(50, self.update_state_and_process) 

    def update_state_and_process(self):
        # Prevent updates if we are already processing or have no image
        if getattr(self, 'is_processing', False) or getattr(self, 'original_image_rgb', None) is None: 
            return
        
        # Turn Auto Detect ON as soon as a slider is touched
        if not self.auto_detect_enabled:
            self.auto_detect_enabled = True
            self.btn_auto.config(text="Auto Detect: ON", fg="green")
        
        # ---> GRAB VALUES FROM ALL 4 SLIDERS <---
        h_min, h_max = self.hue_slider.get_values()
        int_min, int_max = self.int_slider.get_values()
        area_min_val, area_max_val = self.area_slider.get_values()
        circ_val = self.circ_slider.get_values()
            
        # Save them into the current image state
        state = self.image_states[self.current_index]
        state['hue_min'] = h_min
        state['hue_max'] = h_max
        state['int_min'] = int_min
        state['int_max'] = int_max
        state['area_min_pos'] = area_min_val
        state['area_max_pos'] = area_max_val
        state['circ_min'] = circ_val  
        
        # Trigger the visual update
        self.process_image()

    def process_image(self):
            if self.cached_hsv is None or not self.image_states: return
            self.is_processing = True
            
            state = self.image_states[self.current_index]
            file_meta = f"Image {self.current_index + 1} of {len(self.image_states)} | {state['file_path']}"
            
            overlay_rgb = self.original_image_rgb.copy()
            total_pixels = self.cached_gray.shape[0] * self.cached_gray.shape[1]
            
            slider_min = state.get('area_min_pos', 0)
            slider_max = state.get('area_max_pos', 1000)
            
            min_area_val = int(((slider_min / 1000.0) ** 4) * total_pixels) 
            max_area_val = int(((slider_max / 1000.0) ** 4) * total_pixels)
            
            state['min_area_actual'] = min_area_val
            state['max_area_actual'] = max_area_val

            # Flag to trace if we drew contours inside the conditional blocks
            contours_drawn_manually = False
            
            # -----------------------------------------------------------------
            # ---> CASE 1: THE OVERRIDE INTERCEPT (SUSPEND SLIDER RUNTIMES) <---
            # -----------------------------------------------------------------
            if getattr(self, 'mask_override_mode', False) and getattr(self, 'override_mask_data', None) is not None:
                # Baseline mask layout starts strictly from the loaded file matrix
                mask_base = self.override_mask_data.copy()
                
                # RETAIN MANUAL TOOLS: Let pencil additions and eraser modifications alter this base map
                if self.current_manual_add is not None:
                    mask_base = cv2.bitwise_or(mask_base, self.current_manual_add)
                if getattr(self, 'eraser_permanent_mask', None) is not None:
                    mask_base = cv2.bitwise_and(mask_base, cv2.bitwise_not(self.eraser_permanent_mask))
                if self.current_manual_remove is not None:
                    mask_base = cv2.bitwise_and(mask_base, cv2.bitwise_not(self.current_manual_remove))
                    
                self.current_mask = mask_base.copy()
                
                # Extract white contours directly from the modified static mask layout
                contours, _ = cv2.findContours(mask_base, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                cv2.drawContours(overlay_rgb, contours, -1, (255, 255, 255), 2)
                contours_drawn_manually = True
                
                # Static metadata message updates
                stats_meta = "Mask View Mode: Raw Image Pipeline Suspended"
                self.lbl_stats_integrated.config(text=f"{file_meta}\n{stats_meta}")
            
            # -----------------------------------------------------------------
            # ---> CASE 2: ORIGINAL ALGORITHMIC AUTO-DETECTION PROCESSING <---
            # -----------------------------------------------------------------
            elif self.auto_detect_enabled:
                h_min, h_max = state.get('hue_min', 0), state.get('hue_max', 179)
                v_min = state.get('int_min', 0)
                v_max = state.get('int_max', 255)
                
                lower_bound = np.array([h_min, 30, v_min]) 
                upper_bound = np.array([h_max, 255, v_max])
                
                mask_filtered = cv2.inRange(self.cached_hsv, lower_bound, upper_bound)
                
                # 1. Base noise cleanup
                kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
                mask_clean = cv2.morphologyEx(mask_filtered, cv2.MORPH_OPEN, kernel)
                mask_clean = cv2.morphologyEx(mask_clean, cv2.MORPH_CLOSE, kernel)
                
                # 2. DYNAMIC MORPHOLOGY (Fibers vs Cells)
                circ_val = state.get('circ_min', 0)
                if circ_val != 0:
                    k_size = int((abs(circ_val) / 100.0) * 20) * 2 + 1 
                    if k_size > 1:
                        dynamic_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k_size, k_size))
                        if circ_val > 0:
                            mask_clean = cv2.morphologyEx(mask_clean, cv2.MORPH_OPEN, dynamic_kernel)
                        else:
                            mask_clean = cv2.morphologyEx(mask_clean, cv2.MORPH_TOPHAT, dynamic_kernel)
                
                mask_visual_uint8 = np.uint8(mask_clean)
                
                if circ_val < 0:
                    bridge_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (15, 15))
                    mask_logic_uint8 = cv2.morphologyEx(mask_visual_uint8, cv2.MORPH_CLOSE, bridge_kernel)
                    ecc_threshold = (abs(circ_val) / 100.0) * 0.85 
                else:
                    mask_logic_uint8 = mask_visual_uint8.copy()
                    ecc_threshold = 0.0
                
                if not hasattr(self, 'eraser_permanent_mask') or self.eraser_permanent_mask is None:
                    self.eraser_permanent_mask = np.zeros_like(mask_visual_uint8)
                
                # Dynamic Lasso Fill for manual drawings
                manual_add_filled = self.current_manual_add.copy() if hasattr(self, 'current_manual_add') else np.zeros_like(mask_visual_uint8)
                if np.max(manual_add_filled) > 0:
                    cnts_add, _ = cv2.findContours(manual_add_filled, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    cv2.drawContours(manual_add_filled, cnts_add, -1, 255, -1)
                    
                manual_remove_filled = self.current_manual_remove.copy() if hasattr(self, 'current_manual_remove') else np.zeros_like(mask_visual_uint8)
                if np.max(manual_remove_filled) > 0:
                    cnts_rem, _ = cv2.findContours(manual_remove_filled, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    cv2.drawContours(manual_remove_filled, cnts_rem, -1, 255, -1)
                
                mask_combined_logic = cv2.bitwise_or(mask_logic_uint8, manual_add_filled)
                mask_combined_logic = cv2.bitwise_and(mask_combined_logic, cv2.bitwise_not(self.eraser_permanent_mask))
                mask_combined_logic = cv2.bitwise_and(mask_combined_logic, cv2.bitwise_not(manual_remove_filled))
                
                labeled_logic, _ = measure.label(mask_combined_logic > 0, return_num=True)
                logic_regions = measure.regionprops(labeled_logic)

                valid_labels = []
                for r in logic_regions:
                    if min_area_val <= r.area <= max_area_val:
                        if circ_val < 0:
                            if r.eccentricity >= ecc_threshold:
                                valid_labels.append(r.label)
                        else:
                            valid_labels.append(r.label)
                
                mask_approved_logic = np.isin(labeled_logic, valid_labels).astype(np.uint8) * 255
                mask_final = cv2.bitwise_and(mask_visual_uint8, mask_approved_logic)
                
                mask_final = cv2.bitwise_or(mask_final, manual_add_filled)
                mask_final = cv2.bitwise_and(mask_final, cv2.bitwise_not(self.eraser_permanent_mask))
                mask_final = cv2.bitwise_and(mask_final, cv2.bitwise_not(manual_remove_filled))

                self.current_mask = mask_final.copy()
                
                # 3. FINAL STATS & INDIVIDUAL ROI CALCULATION
                labeled_final, num_clusters = measure.label(mask_final > 0, return_num=True)
                final_regions = measure.regionprops(labeled_final, intensity_image=self.cached_gray)

                mean_intensity = np.mean([r.intensity_mean for r in final_regions]) if num_clusters > 0 else 0
                areas_total = sum([r.area for r in final_regions])
                area_percentage = (areas_total / total_pixels) * 100 if total_pixels > 0 else 0

                if 'pixel_size_um' not in state:
                    state['pixel_size_um'] = self.get_pixel_size_um(state['file_path'])
                
                pixel_size = state['pixel_size_um']
                if pixel_size is not None:
                    area_um2 = areas_total * (pixel_size ** 2)
                    area_um2_str = f" ({round(area_um2, 2)} sq \u03BCm)"
                else:
                    area_um2 = 0.0
                    area_um2_str = " (Scale Unknown)"

                state['stats'] = {
                    'area': float(areas_total),
                    'area_percentage': round(area_percentage, 2),
                    'area_um2': round(area_um2, 2), 
                    'cluster_count': num_clusters,
                    'mean_intensity': round(mean_intensity, 2)
                }

                contours, _ = cv2.findContours(mask_final, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                cv2.drawContours(overlay_rgb, contours, -1, (255, 255, 255), 2)
                contours_drawn_manually = True
                
                # =================================================================
                # ---> CLEAN INDEX-ONLY ROI RENDERING ENGINE <---
                # =================================================================
                state['roi_data'] = []
                zoom = getattr(self, 'zoom_factor', 1.0)
                
                all_centroids = np.array([r.centroid for r in final_regions]) if num_clusters > 0 else np.array([])
                
                for idx, r in enumerate(final_regions):
                    roi_id = idx + 1
                    cy, cx = r.centroid
                    cx, cy = int(cx), int(cy)
                    
                    mean_gray = r.intensity_mean
                    int_density = r.area * mean_gray
                    min_int = r.min_intensity
                    max_int = r.max_intensity
                    
                    r_area_um2 = r.area * (pixel_size ** 2) if pixel_size else None
                    roi_area_fraction = (r.area / total_pixels) * 100
                    
                    if num_clusters > 1:
                        distances = np.linalg.norm(all_centroids - np.array([r.centroid]), axis=1)
                        distances[idx] = np.inf 
                        nearest_neighbor_px = np.min(distances)
                        nearest_neighbor_um = nearest_neighbor_px * pixel_size if pixel_size else None
                    else:
                        nearest_neighbor_px = np.nan
                        nearest_neighbor_um = np.nan
                        
                    perimeter = r.perimeter
                    circularity = (4 * np.pi * r.area) / (perimeter ** 2) if perimeter > 0 else 0.0
                    circularity = min(1.0, circularity) 
                    
                    state['roi_data'].append({
                        'ROI ID': roi_id,
                        'Centroid X (px)': cx,
                        'Centroid Y (px)': cy,
                        'Mean Gray Value': round(mean_gray, 2),
                        'Integrated Density': round(int_density, 2),
                        'Min Intensity': int(min_int),
                        'Max Intensity': int(max_int),
                        'Area (px)': r.area,
                        'Area (sq um)': round(r_area_um2, 2) if r_area_um2 else 'Unknown',
                        'Area Fraction (%)': round(roi_area_fraction, 4),
                        'Perimeter (px)': round(perimeter, 2),
                        'Circularity': round(circularity, 3),
                        'Eccentricity': round(r.eccentricity, 3),
                        'Nearest Neighbor Dist (px)': round(nearest_neighbor_px, 1) if not np.isnan(nearest_neighbor_px) else 'N/A',
                        'Nearest Neighbor Dist (um)': round(nearest_neighbor_um, 2) if nearest_neighbor_um else 'N/A'
                    })
                    
                    # ---> NEW: TOGGLE CHECK <---
                    if getattr(self, 'show_roi_numbers', True):
                        # Adaptive Text Sizing (Only renders numbers)
                        if zoom < 0.8:
                            if r.area < 300: continue
                            font_scale = 0.35
                        elif zoom < 1.8:
                            if r.area < 50: continue
                            font_scale = 0.45
                        else:
                            font_scale = 0.55
                        
                        text_label = f"{roi_id}"
                        
                        # Render high-contrast clear text vectors
                        cv2.putText(overlay_rgb, text_label, (cx, cy), cv2.FONT_HERSHEY_SIMPLEX, font_scale, (0, 0, 0), 3, cv2.LINE_AA)
                        cv2.putText(overlay_rgb, text_label, (cx, cy), cv2.FONT_HERSHEY_SIMPLEX, font_scale, (255, 255, 255), 1, cv2.LINE_AA)
                # =================================================================
                
                stats_meta = f"Fluorescent Area: {round(area_percentage, 2)}%{area_um2_str} | Clusters: {num_clusters}"
                self.lbl_stats_integrated.config(text=f"{file_meta}\n{stats_meta}")
            
            # -----------------------------------------------------------------
            # ---> CASE 3: PIPELINE AND OVERRIDES ALL OFF <---
            # -----------------------------------------------------------------
            else:
                self.current_mask = None
                self.lbl_stats_integrated.config(text=f"{file_meta}\nView: Original Image (Auto Detect OFF)")

            # ---> DRAW TRANSLUCENT TEMPORARY ERASER OUTLINES WHILE DRAGGING <---
            if self.current_manual_remove is not None and np.any(self.current_manual_remove > 0):
                rem_contours, _ = cv2.findContours(self.current_manual_remove, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                red_border_overlay = overlay_rgb.copy()
                cv2.drawContours(red_border_overlay, rem_contours, -1, (255, 0, 0), 2)
                cv2.addWeighted(red_border_overlay, 0.80, overlay_rgb, 0.20, 0, overlay_rgb)

            # ---> PENCIL & CIRCLE BORDERS (Only drawn if not already rendered inside Case 1 / Case 2) <---
            if not contours_drawn_manually:
                if self.current_manual_add is not None and np.any(self.current_manual_add > 0):
                    add_contours, _ = cv2.findContours(self.current_manual_add, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    white_border_overlay = overlay_rgb.copy()
                    cv2.drawContours(white_border_overlay, add_contours, -1, (255, 255, 255), 2)
                    cv2.addWeighted(white_border_overlay, 0.80, overlay_rgb, 0.20, 0, overlay_rgb) 

            # -----------------------------------------------------------------
            # ---> TKINTER SCREEN CANVAS CALCULATIONS & POSITIONING <---
            # -----------------------------------------------------------------
            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            if canvas_w < 10: canvas_w, canvas_h = 800, 500 
            
            img_h, img_w = overlay_rgb.shape[:2]
            
            base_scale = min(canvas_w / img_w, canvas_h / img_h)
            self.base_w = max(1, int(img_w * base_scale))
            self.base_h = max(1, int(img_h * base_scale))
            
            self.scale_x = img_w / self.base_w
            self.scale_y = img_h / self.base_h
            self.offset_x = (canvas_w - self.base_w) // 2
            self.offset_y = (canvas_h - self.base_h) // 2

            # Cache array in PIL container so fast trackpad zooming is fluid
            self.base_pil_image = Image.fromarray(overlay_rgb)
            self.is_processing = False
            
            # Trigger screen paint updates
            self.fast_redraw()

    def fast_redraw(self):
        if not hasattr(self, 'base_pil_image'): return
        
        zoom = getattr(self, 'zoom_factor', 1.0)
        pan_x = getattr(self, 'pan_x', 0)
        pan_y = getattr(self, 'pan_y', 0)

        zoomed_w = max(1, int(self.base_w * zoom))
        zoomed_h = max(1, int(self.base_h * zoom))
        
        # ---> THE CRITICAL FIX: Use LANCZOS when zoom < 2.0 for maximum sharpness <---
        # NEAREST is perfect for zoomed-in inspection (prevents fuzziness at high zoom)
        # LANCZOS is perfect for normal/zoomed-out views (stops downsampling blur)
        resample_method = Image.Resampling.NEAREST if zoom >= 2.0 else Image.Resampling.LANCZOS
        zoomed_img = self.base_pil_image.resize((zoomed_w, zoomed_h), resample_method)
        
        self.tk_img = ImageTk.PhotoImage(zoomed_img)
        self.canvas.delete("all")
        
        draw_x = (self.offset_x * zoom) + pan_x
        draw_y = (self.offset_y * zoom) + pan_y
        
        # 1. Draw the actual microscope image
        self.canvas.create_image(draw_x, draw_y, anchor=tk.NW, image=self.tk_img)

        # ---------------------------------------------------------
        # 2. DRAW NATIVE TKINTER SCALE BAR (FLOATING HUD)
        # ---------------------------------------------------------
        pixel_size_um = getattr(self, 'pixel_size_um', None)
        
        # If no metadata exists, don't draw anything (safest approach for science apps)
        if pixel_size_um is None or pixel_size_um <= 0:
            return
            
        # self.scale_x is (original_width / canvas_base_width)
        # zoom is the trackpad zoom multiplier
        scale_x = getattr(self, 'scale_x', 1.0) or 1.0
        
        # Calculate how many real-world micrometers 1 screen pixel represents right now
        microns_per_screen_pixel = (scale_x / zoom) * pixel_size_um

        target_screen_pixels = 150
        real_dist_um = target_screen_pixels * microns_per_screen_pixel
        
        if real_dist_um > 0:
            import math
            magnitude = 10 ** math.floor(math.log10(real_dist_um))
            val = real_dist_um / magnitude

            if val < 2:   nice_val = 1 * magnitude
            elif val < 5: nice_val = 2 * magnitude
            else:         nice_val = 5 * magnitude

            actual_screen_pixels = int(nice_val / microns_per_screen_pixel)
            text_val = int(nice_val) if float(nice_val).is_integer() else round(nice_val, 2)
            text = f"{text_val} \u03BCm" 
            
            # Pin strictly to the Canvas viewport boundaries (bottom right corner)
            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            
            margin_x, margin_y = 30, 30
            bar_height = 8

            x1 = canvas_w - margin_x - actual_screen_pixels
            y1 = canvas_h - margin_y - bar_height
            x2 = canvas_w - margin_x
            y2 = canvas_h - margin_y

            text_x = x1 + (actual_screen_pixels / 2)
            text_y = y1 - 10

            # A. Draw high-contrast black outline/shadows
            self.canvas.create_rectangle(x1-2, y1-2, x2+2, y2+2, fill="black", outline="black")
            for dx, dy in [(-1,-1), (-1,1), (1,-1), (1,1)]:
                self.canvas.create_text(text_x+dx, text_y+dy, text=text, fill="black", font=("Arial", 12, "bold"))
                
            # B. Draw white foreground
            self.canvas.create_rectangle(x1, y1, x2, y2, fill="white", outline="white")
            self.canvas.create_text(text_x, text_y, text=text, fill="white", font=("Arial", 12, "bold"))
    
    # --- Panning ---
    def start_pan(self, event):
        self.pan_start_x = event.x
        self.pan_start_y = event.y

    def pan_motion(self, event):
        # Calculate how far the mouse has moved
        dx = event.x - self.pan_start_x
        dy = event.y - self.pan_start_y
        
        # Apply the movement to our pan variables
        self.pan_x += dx
        self.pan_y += dy
        
        # Reset the start position for the next movement tick
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        
        self.fast_redraw()
    
    # --- Drawing and Correction ---    
    def set_draw_mode(self, mode):
        """--- Drawing Mode (Pencil / Circle / Eraser) ---"""
        self.draw_mode = mode
        
        # Reset and clear any uncommitted interactive circle adjustments
        self.clear_active_oval_handles()
        if hasattr(self, 'temp_oval_id') and self.temp_oval_id:
            self.canvas.delete(self.temp_oval_id)
            self.temp_oval_id = None

        # Clean old button states down to standard style configurations
        self.btn_pencil.config(relief=tk.RAISED, bg="SystemButtonFace")
        self.btn_circle.config(relief=tk.RAISED, bg="SystemButtonFace")
        self.btn_eraser.config(relief=tk.RAISED, bg="SystemButtonFace")

        # Engage selected tool state styles and configurations
        if mode == "pencil":
            self.btn_pencil.config(relief=tk.SUNKEN, bg="lightgray")
            self.canvas.config(cursor="crosshair")
        elif mode == "eraser":
            self.btn_eraser.config(relief=tk.SUNKEN, bg="lightgray")
            self.canvas.config(cursor="circle")
        elif mode == "circle":
            self.btn_circle.config(relief=tk.SUNKEN, bg="lightgray")
            self.canvas.config(cursor="tcross")

    def save_state_for_undo(self):
        if not self.image_states: return
        state = self.image_states[self.current_index]
        add_copy = state['manual_mask_add'].copy()
        remove_copy = state['manual_mask_remove'].copy()
        state['undo_stack'].append((add_copy, remove_copy))
        if len(state['undo_stack']) > 20: state['undo_stack'].pop(0)
        state['redo_stack'].clear()

    def undo_action(self):
        if not self.image_states: return
        state = self.image_states[self.current_index]
        if not state['undo_stack']: return
        self.clear_active_oval_handles() # Dismiss active ovals before undo
        current_add = state['manual_mask_add'].copy()
        current_remove = state['manual_mask_remove'].copy()
        state['redo_stack'].append((current_add, current_remove))
        prev_add, prev_remove = state['undo_stack'].pop()
        state['manual_mask_add'] = prev_add
        state['manual_mask_remove'] = prev_remove
        self.current_manual_add = state['manual_mask_add']
        self.current_manual_remove = state['manual_mask_remove']
        self.process_image()

    def redo_action(self):
        if not self.image_states: return
        state = self.image_states[self.current_index]
        if not state['redo_stack']: return
        self.clear_active_oval_handles() # Dismiss active ovals before redo
        current_add = state['manual_mask_add'].copy()
        current_remove = state['manual_mask_remove'].copy()
        state['undo_stack'].append((current_add, current_remove))
        next_add, next_remove = state['redo_stack'].pop()
        state['manual_mask_add'] = next_add
        state['manual_mask_remove'] = next_remove
        self.current_manual_add = state['manual_mask_add']
        self.current_manual_remove = state['manual_mask_remove']
        self.process_image()

    def start_draw(self, event):
        # 1. Reverse Pan & Zoom to get the base canvas coordinate
        canvas_x = (event.x - getattr(self, 'pan_x', 0)) / getattr(self, 'zoom_factor', 1.0)
        canvas_y = (event.y - getattr(self, 'pan_y', 0)) / getattr(self, 'zoom_factor', 1.0)

        # 2. Check if user is adjusting an existing circle handle
        if self.draw_mode == "circle" and getattr(self, 'handle_ids', None):
            clicked_item = self.canvas.find_withtag("current")
            if clicked_item and clicked_item[0] in self.handle_ids:
                idx = self.handle_ids.index(clicked_item[0])
                self.active_handle = ["T", "B", "L", "R"][idx]
                return

        # 3. Handle standard pencil/eraser click actions
        if self.draw_mode in ("pencil", "eraser"):
            self.clear_active_oval_handles() # Commit previous shapes if tool changed
            self.save_state_for_undo() 
            self.is_drawing = True
            self.last_x, self.last_y = event.x, event.y
            
            orig_x = int((canvas_x - getattr(self, 'offset_x', 0)) * getattr(self, 'scale_x', 1.0))
            orig_y = int((canvas_y - getattr(self, 'offset_y', 0)) * getattr(self, 'scale_y', 1.0))
            self.draw_points_img = [(orig_x, orig_y)]
            
        # 4. Handle extra circle tool initiation click actions
        elif self.draw_mode == "circle":
            self.clear_active_oval_handles()
            self.oval_start_x = canvas_x
            self.oval_start_y = canvas_y

    def draw_motion(self, event):
        # 1. Reverse Pan & Zoom
        canvas_x = (event.x - getattr(self, 'pan_x', 0)) / getattr(self, 'zoom_factor', 1.0)
        canvas_y = (event.y - getattr(self, 'pan_y', 0)) / getattr(self, 'zoom_factor', 1.0)

        # [Keep Case A for circle adjustment exactly as it is]

        # Case B: Standard pencil or eraser dragging logic running natively
        if self.draw_mode in ("pencil", "eraser") and getattr(self, 'is_drawing', False):
            if not self.auto_detect_enabled:
                self.auto_detect_enabled = True
                self.btn_auto.config(text="Auto Detect: ON", fg="white")
                
            color = "red" if self.draw_mode == "eraser" else "white"
            self.canvas.create_line(
                self.last_x, self.last_y, event.x, event.y, 
                fill=color, width=4, capstyle=tk.ROUND, joinstyle=tk.ROUND,
                tags="temp_eraser_line" if self.draw_mode == "eraser" else ""
            )
            
            orig_x = int((canvas_x - getattr(self, 'offset_x', 0)) * getattr(self, 'scale_x', 1.0))
            orig_y = int((canvas_y - getattr(self, 'offset_y', 0)) * getattr(self, 'scale_y', 1.0))
            self.draw_points_img.append((orig_x, orig_y))
            
            # Write to active preview matrix smoothly
            if len(self.draw_points_img) > 1 and self.draw_mode == "eraser" and self.current_manual_remove is not None:
                cv2.line(self.current_manual_remove, self.draw_points_img[-2], self.draw_points_img[-1], 255, thickness=4)

            self.last_x, self.last_y = event.x, event.y
            
        # Case C: Dynamic circle initial shape stretching drag operations
        elif self.draw_mode == "circle" and hasattr(self, 'oval_start_x'):
            cx = self.oval_start_x
            cy = self.oval_start_y
            rx_val = max(1.0, abs(canvas_x - cx))
            ry_val = max(1.0, abs(canvas_y - cy))
            
            self.active_oval = (cx, cy, rx_val, ry_val)
            self.update_oval_visual_layer()

    def stop_draw(self, event):
        # Case A: Handle standard pencil/eraser mask baking loops
        if self.draw_mode in ("pencil", "eraser") and getattr(self, 'is_drawing', False):
            self.is_drawing = False
            
            if hasattr(self, 'draw_points_img') and len(self.draw_points_img) > 0:
                if self.draw_mode == "pencil":
                    if self.current_manual_add is not None:
                        if len(self.draw_points_img) > 2:
                            pts = np.array([self.draw_points_img], dtype=np.int32)
                            cv2.fillPoly(self.current_manual_add, pts, 255)
                        else:
                            dynamic_radius = max(2, int(15 / getattr(self, 'zoom_factor', 1.0)))
                            cv2.circle(self.current_manual_add, self.draw_points_img[0], radius=dynamic_radius, color=255, thickness=-1)
                
                elif self.draw_mode == "eraser":
                    if not hasattr(self, 'eraser_permanent_mask') or self.eraser_permanent_mask is None:
                        if self.current_manual_remove is not None:
                            self.eraser_permanent_mask = np.zeros_like(self.current_manual_remove)
                    
                    # Create a temporary mask to render the solid lasso fill zone
                    eraser_temp = np.zeros_like(self.current_manual_remove)
                    
                    if len(self.draw_points_img) > 2:
                        pts = np.array([self.draw_points_img], dtype=np.int32)
                        # ---> THE CRITICAL FIX: Fill the entire enclosed shape space solidly <---
                        cv2.fillPoly(eraser_temp, pts, 255)
                        if self.eraser_permanent_mask is not None:
                            cv2.fillPoly(self.eraser_permanent_mask, pts, 255)
                    else:
                        dynamic_radius = max(2, int(15 / getattr(self, 'zoom_factor', 1.0)))
                        cv2.circle(eraser_temp, self.draw_points_img[0], radius=dynamic_radius, color=255, thickness=-1)
                        if self.eraser_permanent_mask is not None:
                            cv2.circle(self.eraser_permanent_mask, self.draw_points_img[0], radius=dynamic_radius, color=255, thickness=-1)
                    
                    # 1. Erase manual pencil drawings inside this solid region
                    if self.current_manual_add is not None:
                        self.current_manual_add[eraser_temp > 0] = 0
                    
                    # 2. Clear out the Tkinter smooth vector visual paths from the interface screen
                    self.canvas.delete("temp_eraser_line")
                    
                    # 3. Reset the visual layer back to pure black
                    if self.current_manual_remove is not None:
                        self.current_manual_remove.fill(0)
                    
            self.draw_points_img = [] 
            self.process_image()  # Heavy calculations run smoothly once on button release
            
        # Case B: Circle tool mouse release
        elif self.draw_mode == "circle":
            if getattr(self, 'active_handle', None):
                self.active_handle = None 
                return
                
            if hasattr(self, 'oval_start_x') and getattr(self, 'active_oval', None):
                del self.oval_start_x
                self.render_oval_adjuster_handles()

    def clear_drawing(self):
        if not self.image_states or self.current_index >= len(self.image_states): return
        self.clear_active_oval_handles() # Dismiss active ovals before clearing
        self.save_state_for_undo() 
        state = self.image_states[self.current_index]
        if state['manual_mask_add'] is not None: state['manual_mask_add'].fill(0)
        if state['manual_mask_remove'] is not None: state['manual_mask_remove'].fill(0)
        self.process_image()

    # --- NEW INTERACTIVE CIRCLE GEOMETRY METHODS ---
    def update_oval_visual_layer(self):
        """Draws or moves the visual preview dashed ellipse vector overlay onto the canvas map."""
        if not getattr(self, 'active_oval', None): return
        cx, cy, rx, ry = self.active_oval
        
        # Translate canvas points forward through Pan/Zoom parameters onto display viewport
        x1 = (cx - rx) * self.zoom_factor + self.pan_x
        y1 = (cy - ry) * self.zoom_factor + self.pan_y
        x2 = (cx + rx) * self.zoom_factor + self.pan_x
        y2 = (cy + ry) * self.zoom_factor + self.pan_y

        if getattr(self, 'temp_oval_id', None):
            self.canvas.coords(self.temp_oval_id, x1, y1, x2, y2)
        else:
            self.temp_oval_id = self.canvas.create_oval(
                x1, y1, x2, y2, outline="yellow", width=2, dash=(4, 4)
            )

    def render_oval_adjuster_handles(self):
        """Creates 4 adjustable side knobs on Top, Bottom, Left, and Right shape bounds."""
        self.clear_active_oval_handles()
        if not getattr(self, 'active_oval', None): return
        cx, cy, rx, ry = self.active_oval
        hr = 5  # Handle handle graphics dot diameter footprint
        
        if not hasattr(self, 'handle_ids'): 
            self.handle_ids = []

        # Coordinate anchor vectors
        positions = [
            (cx, cy - ry),  # Top
            (cx, cy + ry),  # Bottom
            (cx - rx, cy),  # Left
            (cx + rx, cy)   # Right
        ]

        for px, py in positions:
            chx = px * self.zoom_factor + self.pan_x
            chy = py * self.zoom_factor + self.pan_y
            hid = self.canvas.create_oval(
                chx - hr, chy - hr, chx + hr, chy + hr, 
                fill="cyan", outline="white", width=1
            )
            self.handle_ids.append(hid)

    def clear_active_oval_handles(self):
        """Bakes adjusted ellipse permanently into masks and deletes overlay graphics nodes."""
        if getattr(self, 'handle_ids', None):
            for hid in self.handle_ids:
                self.canvas.delete(hid)
            self.handle_ids = []

        if getattr(self, 'active_oval', None):
            cx, cy, rx, ry = self.active_oval
            
            # Translate canvas pixels back to match the offset/scale variables of the underlying raw file matrix
            orig_cx = int((cx - getattr(self, 'offset_x', 0)) * getattr(self, 'scale_x', 1.0))
            orig_cy = int((cy - getattr(self, 'offset_y', 0)) * getattr(self, 'scale_y', 1.0))
            orig_rx = int(rx * getattr(self, 'scale_x', 1.0))
            orig_ry = int(ry * getattr(self, 'scale_y', 1.0))
            
            if self.current_manual_add is not None and orig_rx > 0 and orig_ry > 0:
                if not self.auto_detect_enabled:
                    self.auto_detect_enabled = True
                    self.btn_auto.config(text="Auto Detect: ON", fg="white")
                    
                self.save_state_for_undo()
                cv2.ellipse(
                    self.current_manual_add, (orig_cx, orig_cy), (orig_rx, orig_ry), 
                    0, 0, 360, 255, -1
                )
                
            self.active_oval = None
            if getattr(self, 'temp_oval_id', None):
                self.canvas.delete(self.temp_oval_id)
                self.temp_oval_id = None
                
            self.process_image()

    # --- PRESET SAVING & LOADING ---
    def load_presets_from_file(self):
        """Loads saved presets from a JSON file on startup."""
        if os.path.exists(self.presets_file):
            try:
                with open(self.presets_file, 'r') as f:
                    data = json.load(f)
                    self.presets_collection = data.get("collection", {})
                    self.pinned_presets = data.get("pinned", [])
            except Exception as e:
                print(f"Failed to load presets: {e}")

    def save_presets_to_file(self):
        """Saves current presets to a JSON file so they survive app closures."""
        try:
            with open(self.presets_file, 'w') as f:
                json.dump({
                    "collection": self.presets_collection,
                    "pinned": self.pinned_presets
                }, f, indent=4)
        except Exception as e:
            print(f"Failed to save presets: {e}")

    def show_preset_dropdown(self):
        if hasattr(self, 'dropdown_window') and self.dropdown_window.winfo_exists():
            self.dropdown_window.destroy()

        self.dropdown_window = tk.Toplevel(self)
        self.dropdown_window.wm_overrideredirect(True) # Removes window borders/title bar

        # Position it exactly under the Apply button
        x = self.btn_apply_preset.winfo_rootx()
        y = self.btn_apply_preset.winfo_rooty() + self.btn_apply_preset.winfo_height()
        self.dropdown_window.geometry(f"+{x}+{y}")

        # Create Listbox
        self.preset_listbox = tk.Listbox(self.dropdown_window, font=("Arial", 10), height=8, width=30, selectbackground="#0078D7")
        self.preset_listbox.pack(fill=tk.BOTH, expand=True)

        # Sort and Populate
        sorted_names = self.pinned_presets.copy()
        other_presets = sorted([name for name in self.presets_collection.keys() if name not in self.pinned_presets])
        sorted_names.extend(other_presets)

        if not sorted_names:
            self.preset_listbox.insert(tk.END, "No presets found.")
            self.preset_listbox.config(state=tk.DISABLED)
        else:
            for name in sorted_names:
                prefix = "★ " if name in self.pinned_presets else "   "
                suffix = " (Active)" if name == self.current_preset else ""
                self.preset_listbox.insert(tk.END, f"{prefix}{name}{suffix}")

        # Bindings for Left and Right Click
        self.preset_listbox.bind("<ButtonRelease-1>", self.on_dropdown_left_click)
        self.preset_listbox.bind("<ButtonRelease-3>", self.on_dropdown_right_click) # Windows/Mac Right Click
        self.preset_listbox.bind("<ButtonRelease-2>", self.on_dropdown_right_click) # Linux Right Click

        # Close the dropdown if the user clicks anywhere else
        self.dropdown_window.bind("<FocusOut>", lambda e: self.dropdown_window.destroy())
        self.dropdown_window.focus_set()

    def get_clean_preset_name_from_listbox(self, index):
        """Helper to strip the stars and (Active) tags from the list text."""
        item_text = self.preset_listbox.get(index)
        if item_text == "No presets found.": return None
        return item_text.replace("★ ", "").replace("   ", "").replace(" (Active)", "")

    def on_dropdown_left_click(self, event):
        idx = self.preset_listbox.nearest(event.y)
        preset_name = self.get_clean_preset_name_from_listbox(idx)
        
        if preset_name and preset_name in self.presets_collection:
            self.apply_specific_preset(preset_name)
            
        self.dropdown_window.destroy()

    def on_dropdown_right_click(self, event):
        idx = self.preset_listbox.nearest(event.y)
        self.preset_listbox.selection_clear(0, tk.END)
        self.preset_listbox.selection_set(idx)
        
        preset_name = self.get_clean_preset_name_from_listbox(idx)
        if not preset_name or preset_name not in self.presets_collection:
            return

        # Create Context Menu
        ctx_menu = tk.Menu(self.dropdown_window, tearoff=0)
        ctx_menu.add_command(label=f"Manage: {preset_name}", state=tk.DISABLED)
        ctx_menu.add_separator()
        
        is_pinned = preset_name in self.pinned_presets
        pin_label = "Unpin" if is_pinned else "Pin"
        
        ctx_menu.add_command(label=pin_label, command=lambda: self.toggle_preset_pin(preset_name))
        ctx_menu.add_command(label="Rename", command=lambda: self.rename_preset(preset_name))
        ctx_menu.add_command(label="Delete", command=lambda: self.delete_preset(preset_name), foreground="red")
        
        ctx_menu.tk_popup(event.x_root, event.y_root)

    def apply_specific_preset(self, preset_name):
        preset_values = self.presets_collection[preset_name]
        
        self.hue_slider.set_values(*preset_values['hue'])
        self.int_slider.set_values(*preset_values['intensity'])
        self.area_slider.set_values(*preset_values['area'])
        self.circ_slider.set_values(preset_values['circularity'])
        
        self.current_preset = preset_name
        self.btn_apply_preset.config(text=f"Preset: {preset_name}")
        self.schedule_update()
            
    def save_as_preset(self):
        preset_name = simpledialog.askstring("Save As Preset", "Enter a name for this preset:", parent=self)
        if not preset_name: return
        
        if preset_name in self.presets_collection:
            if not messagebox.askyesno("Overwrite?", f"A preset named '{preset_name}' already exists.\nOverwrite?", parent=self):
                return
                
        h_min, h_max = self.hue_slider.get_values()
        i_min, i_max = self.int_slider.get_values()
        a_min, a_max = self.area_slider.get_values()
        c_val = self.circ_slider.get_values()
        
        self.presets_collection[preset_name] = {
            'hue': (h_min, h_max),
            'intensity': (i_min, i_max),
            'area': (a_min, a_max),
            'circularity': c_val
        }
        self.current_preset = preset_name
        self.btn_apply_preset.config(text=f"Preset: {preset_name}")
        self.save_presets_to_file() # <--- SAVES TO FILE
        messagebox.showinfo("Success", f"Preset '{preset_name}' saved successfully.", parent=self)

    def toggle_preset_pin(self, preset_name):
        if preset_name in self.pinned_presets:
            self.pinned_presets.remove(preset_name)
        else:
            self.pinned_presets.append(preset_name)
        
        self.dropdown_window.destroy()
        self.save_presets_to_file() # <--- SAVES TO FILE

    def rename_preset(self, old_name):
        new_name = simpledialog.askstring("Rename Preset", f"Enter new name for '{old_name}':", parent=self, initialvalue=old_name)
        if not new_name or new_name == old_name: return 
        
        if new_name in self.presets_collection:
            messagebox.showerror("Error", f"A preset named '{new_name}' already exists.", parent=self)
            return
            
        self.presets_collection[new_name] = self.presets_collection.pop(old_name)
        if old_name in self.pinned_presets:
            idx = self.pinned_presets.index(old_name)
            self.pinned_presets[idx] = new_name
            
        if self.current_preset == old_name:
            self.current_preset = new_name
            self.btn_apply_preset.config(text=f"Preset: {new_name}")
            
        self.dropdown_window.destroy()
        self.save_presets_to_file() # <--- SAVES TO FILE

    def delete_preset(self, preset_name):
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete preset '{preset_name}'?", parent=self):
            return
            
        self.presets_collection.pop(preset_name)
        if preset_name in self.pinned_presets:
            self.pinned_presets.remove(preset_name)
            
        if self.current_preset == preset_name:
            self.current_preset = None
            self.btn_apply_preset.config(text="Apply Preset")
            
        self.dropdown_window.destroy()
        self.save_presets_to_file() # <--- SAVES TO FILE
    
   # --- Image Switching ---
    def next_image(self):
        self.current_mask = None
        if self.current_index < len(self.image_states) - 1:
            self.current_index += 1
            
            # ---> FIX 1: Manage background RAM thread prefetches
            if hasattr(self, 'manage_cache_pipeline'):
                self.manage_cache_pipeline()
                
            self.load_current_image_data()
            
            # ---> FIX 2: Refresh UI Button States so they remain active/clickable
            self.update_nav_button_states()

    def prev_image(self):
        self.current_mask = None
        if self.current_index > 0:
            self.current_index -= 1
            
            # ---> FIX 1: Manage background RAM thread prefetches
            if hasattr(self, 'manage_cache_pipeline'):
                self.manage_cache_pipeline()
                
            self.load_current_image_data()
            
            # ---> FIX 2: Refresh UI Button States so they remain active/clickable
            self.update_nav_button_states()

    def update_nav_button_states(self):
        """Helper to safely toggle button availability based on array limits."""
        if not hasattr(self, 'image_states') or not self.image_states:
            if hasattr(self, 'btn_prev_img'): self.btn_prev_img.config(state=tk.DISABLED)
            if hasattr(self, 'btn_next_img'): self.btn_next_img.config(state=tk.DISABLED)
            return

        if self.current_index > 0:
            self.btn_prev_img.config(state=tk.NORMAL)
        else:
            self.btn_prev_img.config(state=tk.DISABLED)

        if self.current_index < len(self.image_states) - 1:
            self.btn_next_img.config(state=tk.NORMAL)
        else:
            self.btn_next_img.config(state=tk.DISABLED)
            
        # ---> THIS HASATTR FLAG PREVENTS RUNTIME CRASHES NOW THAT THE LABEL IS GONE <---
        if hasattr(self, 'lbl_img_count') and self.lbl_img_count is not None:
            self.lbl_img_count.config(text=f"Image {self.current_index + 1} of {len(self.image_states)}")

    # --- Export Data ---
    def export_excel(self):
        import os
        import numpy as np
        import pandas as pd
        from tkinter import messagebox, filedialog

        # 1. Check if the list is empty
        if not hasattr(self, 'image_states') or not self.image_states:
            messagebox.showwarning("No Data", "There is no analyzed data to export yet!\nPlease load and process images first.")
            return
            
        final_results = []
        roi_sheets_data = {}  # Dictionary to hold ROI DataFrames for individual sheets

        for state in self.image_states:
            # We only want to export images that have actually been processed
            if 'stats' in state:
                used_manual = False
                if state.get('manual_mask_add') is not None and np.sum(state['manual_mask_add']) > 0: 
                    used_manual = True
                if state.get('manual_mask_remove') is not None and np.sum(state['manual_mask_remove']) > 0: 
                    used_manual = True

                # Format your slider ranges
                intensity_range = f"{state.get('int_min', 'N/A')}-{state.get('int_max', 'N/A')}"
                hue_range = f"{state.get('hue_min', 'N/A')}-{state.get('hue_max', 'N/A')}"
                
                # Fetch Area metrics to compute conversion factor
                total_px_area = state['stats'].get('area', 0)
                absolute_area = state['stats'].get('area_um2', 0.0)
                area_percentage = state['stats'].get('area_percentage', 0)
                
                # Fetch the pixel min/max limits
                min_px = state.get('min_area_actual', 'N/A')
                max_px = state.get('max_area_actual', 'N/A')
                
                # Convert the pixel range to square microns
                if absolute_area == 0.0 and total_px_area > 0:
                    absolute_area_val = "Scale Unknown"
                    size_range_um2 = "Scale Unknown"
                elif total_px_area > 0 and min_px != 'N/A' and max_px != 'N/A':
                    absolute_area_val = absolute_area
                    # Find how many sq microns one pixel represents
                    sq_um_per_px = absolute_area / total_px_area
                    min_um2 = round(min_px * sq_um_per_px, 2)
                    max_um2 = round(max_px * sq_um_per_px, 2)
                    size_range_um2 = f"{min_um2}-{max_um2}"
                else:
                    absolute_area_val = absolute_area if absolute_area > 0 else 0
                    size_range_um2 = "N/A"
                
                # Get just the file name
                full_path = state.get('file_path', 'Unknown')
                file_name_only = os.path.basename(full_path) if full_path != 'Unknown' else 'Unknown'
                
                # Normalize the intensity
                raw_intensity = state['stats'].get('mean_intensity', 0)
                normalized_intensity = round((raw_intensity / 255.0) * 100, 2) 

                # Microscopy-specific headers for Global Overview
                final_results.append({
                    'File Name': file_name_only,
                    'Fluorescent Area (%)': area_percentage,
                    'Absolute Area (sq \u03BCm)': absolute_area_val,
                    'Mean Fluorescence Intensity (%)': normalized_intensity,
                    'Detected Clusters (Count)': state['stats'].get('cluster_count', 0),
                    'Color Filter (Hue Range)': hue_range,
                    'Intensity Threshold (Min-Max)': intensity_range,
                    'Size Range (sq \u03BCm)': size_range_um2,
                    'Fiber Stripping (Circularity)': state.get('circ_min', 'N/A'),
                    'Manual Annotations Applied': used_manual
                })

                # Prepare individual sheet data for ROIs
                roi_records = state.get('roi_data', [])
                sheet_title = os.path.splitext(file_name_only)[0]
                
                # Excel limits sheet names to 31 characters
                if len(sheet_title) > 30:
                    sheet_title = sheet_title[:27] + "..."
                    
                if roi_records:
                    roi_sheets_data[sheet_title] = pd.DataFrame(roi_records)
                else:
                    empty_df = pd.DataFrame(columns=[
                        'ROI ID', 'Centroid X (px)', 'Centroid Y (px)', 'Mean Gray Value', 
                        'Integrated Density', 'Min Intensity', 'Max Intensity', 'Area (px)', 
                        'Area (sq um)', 'Area Fraction (%)', 'Perimeter (px)', 'Circularity', 
                        'Eccentricity', 'Nearest Neighbor Dist (px)', 'Nearest Neighbor Dist (um)'
                    ])
                    empty_df.loc[0] = ['No ROIs Detected'] + ['-'] * 14
                    roi_sheets_data[sheet_title] = empty_df
                
        # 2. Check if we found stats but the final list is still empty
        if not final_results:
            messagebox.showwarning("Incomplete Data", "Images were found, but Auto Detect hasn't been turned on to quantify them yet.")
            return

        # 3. Proceed to save
        save_path = filedialog.asksaveasfilename(
            title="Export Finalized Data",
            defaultextension=".xlsx", 
            filetypes=[("Excel Workbook", "*.xlsx"), ("CSV Document", "*.csv")]
        )
        
        if save_path:
            try:
                df_global = pd.DataFrame(final_results)
                
                # Route to the correct pandas export engine based on the file extension
                if save_path.lower().endswith('.csv'):
                    # CSV cannot hold multiple sheets or styling, so we export only the Batch Overview
                    df_global.to_csv(save_path, index=False, encoding='utf-8-sig') 
                else:
                    # Excel Engine: Multi-sheet with Aesthetic Formatting
                    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                        # Write Sheet 1: Batch Overview
                        df_global.to_excel(writer, sheet_name='Global Overview', index=False)
                        
                        # Write Sheets 2-N: Individual ROI Metrics
                        for sheet_name, df_roi in roi_sheets_data.items():
                            final_sheet_name = sheet_name
                            counter = 1
                            # Failsafe for duplicate file names resolving to the same sheet name
                            while final_sheet_name in writer.sheets and final_sheet_name != 'Global Overview':
                                final_sheet_name = f"{sheet_name[:25]}_{counter}"
                                counter += 1
                            df_roi.to_excel(writer, sheet_name=final_sheet_name, index=False)
                        
                        # --- Apply Aesthetic Styling ---
                        from openpyxl.styles import Font, PatternFill, Alignment
                        
                        workbook = writer.book
                        header_font = Font(bold=True, color="FFFFFF")
                        # Scientific standard blue header
                        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                        center_align = Alignment(horizontal="center", vertical="center")
                        
                        for ws in workbook.worksheets:
                            # Freeze the top header row
                            ws.freeze_panes = 'A2'
                            
                            for col in ws.columns:
                                max_length = 0
                                col_letter = col[0].column_letter
                                
                                for cell in col:
                                    if cell.row == 1:
                                        cell.font = header_font
                                        cell.fill = header_fill
                                        cell.alignment = center_align
                                    else:
                                        cell.alignment = center_align
                                        
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                
                                # Auto-adjust column width with a slight buffer (max capped at 40)
                                adjusted_width = (max_length + 2)
                                ws.column_dimensions[col_letter].width = min(adjusted_width, 40)
                                
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export data:\n{e}")

    # --- Binary Mask ---
    def get_active_binary_mask(self):
        """Returns the master binary mask cached during image processing."""
        if getattr(self, 'original_image_rgb', None) is None:
            return None
        return getattr(self, 'current_mask', None)

    def save_mask_as_png(self):
        """Asks for a custom color and saves contours as a transparent PNG with correct colors."""
        mask = self.get_active_binary_mask()
        
        if mask is None or np.sum(mask) == 0:
            messagebox.showwarning("No Mask Found", "There are no outlined features or pencil tracks to export.")
            return

        # 1. Open the native color selection window
        color_choice = colorchooser.askcolor(
            initialcolor="#ffffff", 
            title="Select Outline Color for Research Figure"
        )
        
        # Abort gracefully if the user cancels or exits the picker window
        if not color_choice or color_choice[0] is None:
            return

        # Extract RGB values directly from the Tkinter result tuple
        r_val = int(color_choice[0][0])
        g_val = int(color_choice[0][1])
        b_val = int(color_choice[0][2])

        # 2. Open the file save dialogue prompt
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            title="Export Colored Borders as Transparent PNG"
        )
        if not file_path:
            return

        try:
            h, w = mask.shape[:2]
            # Create a 4-channel image (RGBA), initialized to fully transparent black (0,0,0,0)
            rgba_outlines = np.zeros((h, w, 4), dtype=np.uint8)
            
            # Extract the geometric edges of your active zones
            contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
            
            # FIXED: Because Image.fromarray converts the raw matrix to RGBA,
            # OpenCV must write colors directly in (R, G, B, A) order.
            line_color = (r_val, g_val, b_val, 255) 
            thickness = 2 
            
            # Render the lines onto our transparent matrix layer
            cv2.drawContours(rgba_outlines, contours, -1, line_color, thickness)
            
            # Save using PIL to properly encode the structural alpha-channel format
            Image.fromarray(rgba_outlines).save(file_path, "PNG")
            messagebox.showinfo("Success", f"Mask borders written in your chosen color to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error Saving", f"Failed to compress PNG output:\n{str(e)}")

    def apply_saved_mask(self):
        """Loads an external transparent PNG mask, disabling auto-detect processing."""
        if self.original_image_rgb is None or not self.image_states:
            messagebox.showwarning("No Image", "Load an active image target before applying masks.")
            return

        file_path = filedialog.askopenfilename(
            filetypes=[("PNG Image", "*.png"), ("All Files", "*.*")],
            title="Select Outline Mask File"
        )
        if not file_path:
            return

        try:
            loaded_img = Image.open(file_path)
            h, w = self.original_image_rgb.shape[:2]
            loaded_img = loaded_img.resize((w, h), Image.Resampling.NEAREST)
            img_np = np.array(loaded_img)
            
            # Extract clean binary information from Alpha channel or Greyscale thresholding
            if img_np.shape[-1] == 4:
                binary_mask = (img_np[:, :, 3] > 0).astype(np.uint8) * 255
            else:
                gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
                _, binary_mask = cv2.threshold(gray, 1, 255, cv2.THRESH_BINARY)
                
            # Activating the override mechanism
            self.mask_override_mode = True
            self.override_mask_data = binary_mask.copy()
            
            # Turn off Auto-Detect visually and functionally
            self.auto_detect_enabled = False
            if hasattr(self, 'btn_auto'):
                self.btn_auto.config(text="Auto Detect: OFF", fg="yellow")
            
            # Fire the interface redraw pipeline
            self.process_image()
            messagebox.showinfo("Applied", "External mask successfully mapped. Heavy image calculations suspended.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to map external data frame:\n{str(e)}")

    def export_current_image_view(self):
        """Saves a lossless 100% resolution image copy of exactly what is active in the preview frame container."""
        if not hasattr(self, 'base_pil_image') or self.base_pil_image is None:
            messagebox.showwarning("No Image Data", "Load an image framework and map contours before attempting an export.")
            return

        # Prompt the user for an output path destination
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Lossless Image", "*.png"), ("TIFF Image", "*.tif *.tiff"), ("JPEG Compressed Image", "*.jpg *.jpeg")],
            title="Export Current Analysis View Layer As..."
        )
        if not file_path:
            return

        try:
            # Grab the underlying master NumPy array frame directly from our high-res cache container
            export_array = np.array(self.base_pil_image)

            # Check format type extension requirements
            if file_path.lower().endswith(('.tif', '.tiff')):
                import tifffile
                # Write an uncompressed high-fidelity scientific tiff array mapping profile
                tifffile.imwrite(file_path, export_array, compression='zlib')
                
            elif file_path.lower().endswith('.png'):
                # Convert the internal script RGB order back to OpenCV BGR color mapping layout
                final_bgr = cv2.cvtColor(export_array, cv2.COLOR_RGB2BGR)
                # Enforce absolute 0% compression down-sampling rules to preserve crystal clear line pixels
                cv2.imwrite(file_path, final_bgr, [int(cv2.IMWRITE_PNG_COMPRESSION), 0])
                
            else:
                final_bgr = cv2.cvtColor(export_array, cv2.COLOR_RGB2BGR)
                # Enforce absolute 100% maximum quality factor rules to bypass standard JPEG macroblock artifacts
                cv2.imwrite(file_path, final_bgr, [int(cv2.IMWRITE_JPEG_QUALITY), 100])

            messagebox.showinfo("Export Complete", f"Successfully exported current analysis image layer to:\n{os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("Export Failure", f"Failed saving rendered screen configuration to file system path grid context:\n{str(e)}")