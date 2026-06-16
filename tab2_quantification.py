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
        root_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # --- Control Bar ---
        control_frame = tk.Frame(root_frame, pady=10)
        control_frame.pack(fill=tk.X)
        
        self.btn_select_images = tk.Button(control_frame, text="1. Select Images...", command=self.load_files, font=("Arial", 10, "bold"))
        self.btn_select_images.pack(side=tk.LEFT, padx=10)
        
        self.btn_auto = tk.Button(control_frame, text="Auto Detect: OFF", command=self.toggle_auto_detect, fg="red", font=("Arial", 10, "bold"))
        self.btn_auto.pack(side=tk.LEFT, padx=(10, 10))
        
        tool_frame = tk.Frame(control_frame, bd=1, relief=tk.SOLID, padx=5, pady=2)
        tool_frame.pack(side=tk.LEFT, padx=10)
        tk.Label(tool_frame, text="Tools:", font=("Arial", 8)).pack(side=tk.LEFT)
        
        self.btn_pencil = tk.Button(tool_frame, text="✏️", relief=tk.RAISED, command=lambda: self.set_draw_mode("pencil"))
        self.btn_pencil.pack(side=tk.LEFT, padx=1)

        self.btn_circle = tk.Button(tool_frame, text="⭕", width=3, command=lambda: self.set_draw_mode("circle"), font=("Arial", 10, "bold"))
        self.btn_circle.pack(side=tk.LEFT, padx=1)
        
        self.btn_eraser = tk.Button(tool_frame, text="🧹", relief=tk.RAISED, command=lambda: self.set_draw_mode("eraser"))
        self.btn_eraser.pack(side=tk.LEFT, padx=2)

        self.btn_undo = tk.Button(tool_frame, text="↩️ Undo", command=self.undo_action)
        self.btn_undo.pack(side=tk.LEFT, padx=2)

        self.btn_redo = tk.Button(tool_frame, text="↪️ Redo", command=self.redo_action)
        self.btn_redo.pack(side=tk.LEFT, padx=2)
        
        tk.Button(control_frame, text="Clear Drawings", command=self.clear_drawing, fg="red").pack(side=tk.LEFT, padx=10)
        
        # --- PRESET BUTTONS ---
        self.btn_apply_preset = tk.Button(control_frame, text="Apply Preset", command=self.show_preset_dropdown)
        self.btn_apply_preset.pack(side=tk.LEFT, padx=2)
        
        tk.Button(control_frame, text="Save As Preset", command=self.save_as_preset).pack(side=tk.LEFT, padx=2)
        # ----------------------

        # --- NEW: MASK IMPORT / EXPORT BUTTONS ---
        tk.Button(control_frame, text="💾 Save Mask", command=self.save_mask_as_png, bg="#1976d2", fg="white", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        # ------------------------------------------

        # --- Saveing data button ---
        tk.Button(control_frame, text="Save Data to Excel/CSV", command=self.export_excel, font=("Arial", 10, "bold"), fg="white", bg="#2e7d32").pack(side=tk.RIGHT, padx=10)


        # --- SLIDER FRAME ---
        slider_frame = tk.Frame(root_frame, pady=10)
        slider_frame.pack(fill=tk.X, padx=10)

        # 1. Hue Range 
        hue_frame = tk.Frame(slider_frame)
        hue_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Label(hue_frame, text="Color Filter (Hue):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.hue_slider = ColorRangeSlider(hue_frame, width=220, height=35, slider_type="hue", abs_min=0, abs_max=179, command=self.schedule_update) 
        self.hue_slider.pack(fill=tk.X, pady=5)

        # 2. Intensity Range 
        int_frame = tk.Frame(slider_frame)
        int_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Label(int_frame, text="Intensity (0=Black, 255=White):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.int_slider = ColorRangeSlider(int_frame, width=220, height=35, slider_type="intensity", abs_min=0, abs_max=255, command=self.schedule_update)
        self.int_slider.pack(fill=tk.X, pady=5)

        # 3. Area Range
        area_frame = tk.Frame(slider_frame)
        area_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        tk.Label(area_frame, text="Area Filter (px):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.area_slider = ColorRangeSlider(area_frame, width=220, height=35, slider_type="area", abs_min=0, abs_max=1000, command=self.schedule_update)
        self.area_slider.pack(fill=tk.X, pady=5)
        
        # 4. Circularity / Split (Updated to Single Slider)
        circ_frame = tk.Frame(slider_frame)
        circ_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(circ_frame, text="Circularity (0=Line, 100=Circle):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.circ_slider = SingleSlider(circ_frame, width=220, height=35, abs_min=0, abs_max=100, command=self.schedule_update)
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

        # Nav
        nav_frame = tk.Frame(root_frame, pady=10)
        nav_frame.pack(fill=tk.X, padx=10)
        
        tk.Button(nav_frame, text="<< Prev", command=self.prev_image, font=("Arial", 10)).pack(side=tk.LEFT)
        
        stats_frame = tk.Frame(nav_frame)
        stats_frame.pack(side=tk.LEFT, expand=True)
        self.lbl_stats_integrated = tk.Label(stats_frame, text="", font=("Arial", 11, "bold"))
        self.lbl_stats_integrated.pack()

        tk.Button(nav_frame, text="Next >>", command=self.next_image, font=("Arial", 10)).pack(side=tk.RIGHT)

        # --- Global Keyboard Bindings for Tab 2 ---
        top = self.winfo_toplevel()
        
        # Left/Right Arrows for Prev/Next File Navigation
        top.bind("<Left>", lambda e: self.prev_image() if self.winfo_ismapped() else None, add="+")
        top.bind("<Right>", lambda e: self.next_image() if self.winfo_ismapped() else None, add="+")
        
        # Ctrl+O for Open/Select Images File browser
        top.bind("<Control-o>", lambda e: self.load_files() if self.winfo_ismapped() else None, add="+")
        top.bind("<Control-O>", lambda e: self.load_files() if self.winfo_ismapped() else None, add="+")

        # --- NEW: Ctrl+Z and Ctrl+Y Hotkey Listeners ---
        # Intercepts keyboard events and routes them to undo/redo systems safely if this tab is visible
        top.bind("<Control-z>", lambda e: self.undo_action() if self.winfo_ismapped() else None, add="+")
        top.bind("<Control-Z>", lambda e: self.undo_action() if self.winfo_ismapped() else None, add="+")
        top.bind("<Control-y>", lambda e: self.redo_action() if self.winfo_ismapped() else None, add="+")
        top.bind("<Control-Y>", lambda e: self.redo_action() if self.winfo_ismapped() else None, add="+")
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
            
            # Normalize 16-bit or 32-bit scientific images down to standard 8-bit array
            if img.dtype != np.uint8:
                img = np.uint8(cv2.normalize(img, None, 0, 255, cv2.NORM_MINMAX))
            
            # Handle grayscale images by converting them to standard RGB matrices
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
        if not self.image_states or self.original_image_rgb is None: return
        self.auto_detect_enabled = not self.auto_detect_enabled
        if self.auto_detect_enabled:
            self.btn_auto.config(text="Auto Detect: ON", fg="green")
        else:
            self.btn_auto.config(text="Auto Detect: OFF", fg="red")
        self.process_image()

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
        
        if self.auto_detect_enabled:
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
            
            # ---> 2. DYNAMIC FIBER STRIPPING (MORPHOLOGICAL OPENING) <---
            circ_min = state.get('circ_min', 0)
            if circ_min > 0:
                k_size = int((circ_min / 100.0) * 20) * 2 + 1 
                if k_size > 1:
                    dynamic_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k_size, k_size))
                    mask_clean = cv2.morphologyEx(mask_clean, cv2.MORPH_OPEN, dynamic_kernel)
            
            mask_final_uint8 = np.uint8(mask_clean)
            
            # ---> NEW: Initialise a permanent eraser tracking mask matching the image dimensions <---
            if not hasattr(self, 'eraser_permanent_mask') or self.eraser_permanent_mask is None:
                self.eraser_permanent_mask = np.zeros_like(mask_final_uint8)
            
            # Combine the slider thresholds and manual pencil drawings
            mask_combined = cv2.bitwise_or(mask_final_uint8, self.current_manual_add)
            
            # ---> FIXED: Subtract BOTH the persistent cuts AND the active temporary dragging strokes <---
            mask_combined = cv2.bitwise_and(mask_combined, cv2.bitwise_not(self.eraser_permanent_mask))
            mask_combined = cv2.bitwise_and(mask_combined, cv2.bitwise_not(self.current_manual_remove))
            
            labeled_mask, _ = measure.label(mask_combined > 0, return_num=True)
            regions = measure.regionprops(labeled_mask, intensity_image=self.cached_gray)

            valid_labels = []
            valid_regions = []
            
            for r in regions:
                if min_area_val <= r.area <= max_area_val:
                    valid_labels.append(r.label)
                    valid_regions.append(r)
            
            mask_filtered_area = np.isin(labeled_mask, valid_labels).astype(np.uint8) * 255

            # ---> NEW: Save the active binary matrix to the instance variable <---
            self.current_mask = mask_filtered_area.copy()
            num_clusters = len(valid_regions)
            # -------------------------------------------

            num_clusters = len(valid_regions)
            mean_intensity = np.mean([r.intensity_mean for r in valid_regions]) if num_clusters > 0 else 0
            areas_total = sum([r.area for r in valid_regions])
            area_percentage = (areas_total / total_pixels) * 100 if total_pixels > 0 else 0

            # 1. Check metadata once and cache it
            if 'pixel_size_um' not in state:
                state['pixel_size_um'] = self.get_pixel_size_um(state['file_path'])
            
            pixel_size = state['pixel_size_um']
            
            # 2. Calculate the square microns
            if pixel_size is not None:
                area_um2 = areas_total * (pixel_size ** 2)
                area_um2_str = f" ({round(area_um2, 2)} sq \u03BCm)"
            else:
                area_um2 = 0.0
                area_um2_str = " (Scale Unknown)"

            # 3. Save stats
            state['stats'] = {
                'area': float(areas_total),
                'area_percentage': round(area_percentage, 2),
                'area_um2': round(area_um2, 2), 
                'cluster_count': num_clusters,
                'mean_intensity': round(mean_intensity, 2)
            }

            contours, _ = cv2.findContours(mask_filtered_area, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            cv2.drawContours(overlay_rgb, contours, -1, (255, 255, 255), 2)
            
            # ---> 4. CRITICAL: Update the string format to include area_um2_str <---
            stats_meta = f"Fluorescent Area: {round(area_percentage, 2)}%{area_um2_str} | Clusters: {num_clusters}"
            self.lbl_stats_integrated.config(text=f"{file_meta}\n{stats_meta}")
        else:
            # ---> NEW: Reset mask tracker if Auto Detect is turned off <---
            self.current_mask = None
            self.lbl_stats_integrated.config(text=f"{file_meta}\nView: Original Image (Auto Detect OFF)")

        # ---> THE TEMPORARY RED BORDER GUIDE (WHILE DRAGGING) <---
        if self.current_manual_remove is not None and np.any(self.current_manual_remove > 0):
            rem_contours, _ = cv2.findContours(self.current_manual_remove, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            red_border_overlay = overlay_rgb.copy()
            # Draws a hollow red outline (255, 0, 0) with a thickness of 2
            cv2.drawContours(red_border_overlay, rem_contours, -1, (255, 0, 0), 2)
            cv2.addWeighted(red_border_overlay, 0.80, overlay_rgb, 0.20, 0, overlay_rgb)

        # ---> PENCIL & CIRCLE BORDERS (White Only, Transparent Inside) <---
        if self.current_manual_add is not None and np.any(self.current_manual_add > 0):
            add_contours, _ = cv2.findContours(self.current_manual_add, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            white_border_overlay = overlay_rgb.copy()
            cv2.drawContours(white_border_overlay, add_contours, -1, (255, 255, 255), 2)
            cv2.addWeighted(white_border_overlay, 0.80, overlay_rgb, 0.20, 0, overlay_rgb) 


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

        # CACHE the heavy image processing result so zooming is instant
        self.base_pil_image = Image.fromarray(overlay_rgb)
        
        self.is_processing = False
        
        # Trigger the lightweight drawing function
        self.fast_redraw()

    # --- Scale Bar (Remove) ---
    def fast_redraw(self):
        if not hasattr(self, 'base_pil_image'): return
        
        zoom = getattr(self, 'zoom_factor', 1.0)
        pan_x = getattr(self, 'pan_x', 0)
        pan_y = getattr(self, 'pan_y', 0)

        zoomed_w = max(1, int(self.base_w * zoom))
        zoomed_h = max(1, int(self.base_h * zoom))
        
        resample_method = Image.Resampling.NEAREST if zoom >= 2.0 else Image.Resampling.BILINEAR
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

        # Case A: Adjusting an existing shape using active boundary handles
        if self.draw_mode == "circle" and getattr(self, 'active_handle', None) and getattr(self, 'active_oval', None):
            cx, cy, rx_val, ry_val = self.active_oval
            if self.active_handle == "L":   rx_val = abs(cx - canvas_x)
            elif self.active_handle == "R": rx_val = abs(canvas_x - cx)
            elif self.active_handle == "T": ry_val = abs(cy - canvas_y)
            elif self.active_handle == "B": ry_val = abs(canvas_y - cy)
            
            self.active_oval = (cx, cy, rx_val, ry_val)
            self.update_oval_visual_layer()
            return

        # Case B: Standard pencil or eraser dragging logic running natively
        if self.draw_mode in ("pencil", "eraser") and getattr(self, 'is_drawing', False):
            if not self.auto_detect_enabled:
                self.auto_detect_enabled = True
                self.btn_auto.config(text="Auto Detect: ON", fg="white")
                
            color = "red" if self.draw_mode == "eraser" else "white"
            self.canvas.create_line(self.last_x, self.last_y, event.x, event.y, fill=color, width=2, capstyle=tk.ROUND)
            
            orig_x = int((canvas_x - getattr(self, 'offset_x', 0)) * getattr(self, 'scale_x', 1.0))
            orig_y = int((canvas_y - getattr(self, 'offset_y', 0)) * getattr(self, 'scale_y', 1.0))
            self.draw_points_img.append((orig_x, orig_y))
            
            # ---> NEW: Feed the eraser preview matrix dynamically as you drag <---
            if self.draw_mode == "eraser" and self.current_manual_remove is not None:
                if len(self.draw_points_img) > 1:
                    # Adjust thickness=8 if you want a wider eraser tool stroke
                    cv2.line(self.current_manual_remove, self.draw_points_img[-2], self.draw_points_img[-1], 255, thickness=8)
                    # Trigger process_image so the red outline updates live during mouse drag
                    self.process_image()

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
                    # Initialise our permanent eraser mask if it hasn't been instantiated yet
                    if not hasattr(self, 'eraser_permanent_mask') or self.eraser_permanent_mask is None:
                        if self.current_manual_remove is not None:
                            self.eraser_permanent_mask = np.zeros_like(self.current_manual_remove)
                    
                    # 1. Capture the exact shape footprint of the eraser movement points list
                    eraser_temp = np.zeros_like(self.current_manual_remove)
                    if len(self.draw_points_img) > 2:
                        pts = np.array([self.draw_points_img], dtype=np.int32)
                        cv2.fillPoly(eraser_temp, pts, 255)
                        if self.eraser_permanent_mask is not None:
                            cv2.fillPoly(self.eraser_permanent_mask, pts, 255)
                    else:
                        dynamic_radius = max(2, int(15 / getattr(self, 'zoom_factor', 1.0)))
                        cv2.circle(eraser_temp, self.draw_points_img[0], radius=dynamic_radius, color=255, thickness=-1)
                        if self.eraser_permanent_mask is not None:
                            cv2.circle(self.eraser_permanent_mask, self.draw_points_img[0], radius=dynamic_radius, color=255, thickness=-1)
                    
                    # 2. Clean out manually drawn pencil contours sitting beneath the eraser path
                    if self.current_manual_add is not None:
                        self.current_manual_add[eraser_temp > 0] = 0
                    
                    # 3. Clear the active visual layer completely so the red border instantly vanishes
                    if self.current_manual_remove is not None:
                        self.current_manual_remove.fill(0)
                    
            self.draw_points_img = [] 
            self.process_image()
            
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
            self.load_current_image_data()

    def prev_image(self):
        self.current_mask = None
        if self.current_index > 0:
            self.current_index -= 1
            self.load_current_image_data()

    # --- Export Data ---
    def export_excel(self):
        # 1. Check if the list is empty
        if not hasattr(self, 'image_states') or not self.image_states:
            messagebox.showwarning("No Data", "There is no analyzed data to export yet!\nPlease load and process images first.")
            return
            
        final_results = []
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

                # Microscopy-specific headers
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
                import pandas as pd
                df = pd.DataFrame(final_results)
                
                # Route to the correct pandas export engine based on the file extension
                if save_path.lower().endswith('.csv'):
                    df.to_csv(save_path, index=False, encoding='utf-8-sig') 
                else:
                    df.to_excel(save_path, index=False)
                    
                # Success popup removed!
                
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
        """Loads a transparent custom-colored PNG mask back into the pipeline."""
        if self.original_image_rgb is None:
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
            
            if img_np.shape[-1] == 4:
                # ROBUST CHECK: Any pixel that has opacity (Alpha > 0) is parsed as a border mask, 
                # regardless of whether the outline was saved as Red, Green, Blue, or White.
                binary_mask = (img_np[:, :, 3] > 0).astype(np.uint8) * 255
            else:
                gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
                _, binary_mask = cv2.threshold(gray, 1, 255, cv2.THRESH_BINARY)
                
            self.current_mask = binary_mask
            
            # Trigger your display pipeline to update the view and calculate stats
            self.process_image()
                
            messagebox.showinfo("Applied", "Mask outlines overlayed onto current frame matrix.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to map external data frame:\n{str(e)}")