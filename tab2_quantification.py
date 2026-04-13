import os
import cv2
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import tifffile
import czifile
from skimage import filters, measure
# --> Import your custom widget from widgets.pyk
from widgets import ColorRangeSlider, SingleSlider

class QuantificationTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.image_states = [] 
        self.current_index = 0
        
        self.original_image_rgb = None
        self.current_manual_add = None
        self.current_manual_remove = None
        
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

        self.setup_ui()

    def load_files(self):
        """Standalone loader for Tab 2: Disconnects from Tab 1 and loads files directly."""
        
        # ---> THE FIX: The Ultimate File Filter List <---
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
        
        for file_path in self.image_files:
            self.image_states.append({
                'file_path': file_path,
                'hue_min': 0, 
                'hue_max': 179,
                # ---> UPDATED: Swapped 'int_thresh' for min and max bounds <---
                'int_min': 0,
                'int_max': 255,
                # ---> UPDATED: Swapped 'min_area_slider_pos' for min and max bounds <---
                'area_min_pos': 30,
                'area_max_pos': 1000,
                'manual_mask_add': None, 
                'manual_mask_remove': None,
                'undo_stack': [], 
                'redo_stack': []  
            })
            
        # Update the UI Label
        self.lbl_file_count.config(text=f"Loaded {len(self.image_files)} images", fg="blue")
        
        # Trigger the loading of the first image in the array
        self.load_current_image_data()

    def setup_ui(self):
        root_frame = tk.Frame(self)
        root_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Control Bar
        control_frame = tk.Frame(root_frame, pady=10)
        control_frame.pack(fill=tk.X)
        
        self.btn_select_images = tk.Button(control_frame, text="1. Select Images...", command=self.load_files, font=("Arial", 10, "bold"))
        self.btn_select_images.pack(side=tk.LEFT, padx=10)
        
        self.lbl_file_count = tk.Label(control_frame, text="No images selected", fg="gray")
        self.lbl_file_count.pack(side=tk.LEFT, padx=10)
        
        self.btn_auto = tk.Button(control_frame, text="Auto Detect: OFF", command=self.toggle_auto_detect, fg="red", font=("Arial", 10, "bold"))
        self.btn_auto.pack(side=tk.LEFT, padx=(20, 10))
        
        tool_frame = tk.Frame(control_frame, bd=1, relief=tk.SOLID, padx=5, pady=2)
        tool_frame.pack(side=tk.LEFT, padx=10)
        tk.Label(tool_frame, text="Tools:", font=("Arial", 8)).pack(side=tk.LEFT)
        
        self.btn_pencil = tk.Button(tool_frame, text="✏️ Pencil", relief=tk.SUNKEN, bg="lightgray", command=lambda: self.set_draw_mode("pencil"))
        self.btn_pencil.pack(side=tk.LEFT, padx=2)
        
        self.btn_eraser = tk.Button(tool_frame, text="🧹 Eraser", relief=tk.RAISED, command=lambda: self.set_draw_mode("eraser"))
        self.btn_eraser.pack(side=tk.LEFT, padx=2)

        self.btn_undo = tk.Button(tool_frame, text="↩️ Undo", command=self.undo_action)
        self.btn_undo.pack(side=tk.LEFT, padx=2)

        self.btn_redo = tk.Button(tool_frame, text="↪️ Redo", command=self.redo_action)
        self.btn_redo.pack(side=tk.LEFT, padx=2)
        
        tk.Button(control_frame, text="Clear Drawings", command=self.clear_drawing, fg="red").pack(side=tk.LEFT, padx=10)
        tk.Button(control_frame, text="Export Finalized Data to Excel", command=self.export_excel, font=("Arial", 10, "bold"), fg="white", bg="#2e7d32").pack(side=tk.RIGHT, padx=10)

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

        # Nav
        nav_frame = tk.Frame(root_frame, pady=10)
        nav_frame.pack(fill=tk.X, padx=10)
        
        tk.Button(nav_frame, text="<< Prev", command=self.prev_image, font=("Arial", 10)).pack(side=tk.LEFT)
        
        stats_frame = tk.Frame(nav_frame)
        stats_frame.pack(side=tk.LEFT, expand=True)
        self.lbl_stats_integrated = tk.Label(stats_frame, text="", font=("Arial", 11, "bold"))
        self.lbl_stats_integrated.pack()

        tk.Button(nav_frame, text="Next >>", command=self.next_image, font=("Arial", 10)).pack(side=tk.RIGHT)
    
    def set_draw_mode(self, mode):
        self.draw_mode = mode
        if mode == "pencil":
            self.btn_pencil.config(relief=tk.SUNKEN, bg="lightgray")
            self.btn_eraser.config(relief=tk.RAISED, bg="SystemButtonFace")
            self.canvas.config(cursor="crosshair")
        else:
            self.btn_eraser.config(relief=tk.SUNKEN, bg="lightgray")
            self.btn_pencil.config(relief=tk.RAISED, bg="SystemButtonFace")
            self.canvas.config(cursor="circle") 

    def toggle_auto_detect(self):
        if not self.image_states or self.original_image_rgb is None: return
        self.auto_detect_enabled = not self.auto_detect_enabled
        if self.auto_detect_enabled:
            self.btn_auto.config(text="Auto Detect: ON", fg="green")
        else:
            self.btn_auto.config(text="Auto Detect: OFF", fg="red")
        self.process_image()

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
            # We use the minimum handle of the Circularity slider to dictate how aggressively 
            # we shrink the shape inwards to snap off fibers.
            circ_min = state.get('circ_min', 0)
            if circ_min > 0:
                # Map the 0-100 slider to a kernel size from 1 to 41 pixels
                k_size = int((circ_min / 100.0) * 20) * 2 + 1 
                if k_size > 1:
                    dynamic_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k_size, k_size))
                    # Erode to snap fibers, Dilate to restore main cell mass
                    mask_clean = cv2.morphologyEx(mask_clean, cv2.MORPH_OPEN, dynamic_kernel)
            
            mask_final_uint8 = np.uint8(mask_clean)
            
            mask_combined = cv2.bitwise_or(mask_final_uint8, self.current_manual_add)
            mask_combined = cv2.bitwise_and(mask_combined, cv2.bitwise_not(self.current_manual_remove))
            
            labeled_mask, _ = measure.label(mask_combined > 0, return_num=True)
            regions = measure.regionprops(labeled_mask, intensity_image=self.cached_gray)
            
            # ---> ONLY FILTER BY AREA NOW (Circularity math removed!) <---
            valid_labels = []
            valid_regions = []
            
            for r in regions:
                if min_area_val <= r.area <= max_area_val:
                    valid_labels.append(r.label)
                    valid_regions.append(r)
            
            mask_filtered_area = np.isin(labeled_mask, valid_labels).astype(np.uint8) * 255
            
            num_clusters = len(valid_regions)
            mean_intensity = np.mean([r.intensity_mean for r in valid_regions]) if num_clusters > 0 else 0
            areas_total = sum([r.area for r in valid_regions])
            area_percentage = (areas_total / total_pixels) * 100 if total_pixels > 0 else 0

            state['stats'] = {
                'area': float(areas_total),
                'area_percentage': round(area_percentage, 2),
                'cluster_count': num_clusters,
                'mean_intensity': round(mean_intensity, 2)
            }

            contours, _ = cv2.findContours(mask_filtered_area, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            cv2.drawContours(overlay_rgb, contours, -1, (0, 255, 0), 2)
            
            stats_meta = f"Fluorescent Area: {round(area_percentage, 2)}% | Clusters: {num_clusters}"
            self.lbl_stats_integrated.config(text=f"{file_meta}\n{stats_meta}")
        else:
            self.lbl_stats_integrated.config(text=f"{file_meta}\nView: Original Image (Auto Detect OFF)")

        if self.current_manual_remove is not None and np.any(self.current_manual_remove > 0):
            red_layer = overlay_rgb.copy()
            red_layer[self.current_manual_remove > 0] = [255, 0, 0]
            cv2.addWeighted(red_layer, 0.35, overlay_rgb, 0.65, 0, overlay_rgb) 

        if self.current_manual_add is not None and np.any(self.current_manual_add > 0):
            green_layer = overlay_rgb.copy()
            green_layer[self.current_manual_add > 0] = [0, 150, 0]
            cv2.addWeighted(green_layer, 0.35, overlay_rgb, 0.65, 0, overlay_rgb) 

        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: canvas_w, canvas_h = 800, 500 
        
        img_h, img_w = overlay_rgb.shape[:2]
        
        scale = min(canvas_w / img_w, canvas_h / img_h)
        new_w = max(1, int(img_w * scale))
        new_h = max(1, int(img_h * scale))
        
        pil_img_display = Image.fromarray(overlay_rgb)
        pil_img_display = pil_img_display.resize((new_w, new_h), Image.Resampling.LANCZOS)
        
        self.scale_x = img_w / new_w
        self.scale_y = img_h / new_h

        self.offset_x = (canvas_w - new_w) // 2
        self.offset_y = (canvas_h - new_h) // 2
        
        self.tk_img = ImageTk.PhotoImage(pil_img_display)
        self.canvas.delete("all")
        self.canvas.create_image(self.offset_x, self.offset_y, anchor=tk.NW, image=self.tk_img)
        
        self.is_processing = False
    
    def load_raw_image_array(self, path):
        """Reads the raw file from disk (TIFF, CZI, JPEG, PNG) into a numpy array."""
        try:
            if path.lower().endswith('.czi'):
                img = czifile.imread(path)
                img = np.squeeze(img)
                if img.ndim == 3 and img.shape[0] < 10: 
                    img = np.transpose(img, (1, 2, 0))
                if img.ndim > 3:
                    img = img[0, ...] 
            elif path.lower().endswith(('.tif', '.tiff')):
                try:
                    img = tifffile.imread(path)
                except Exception:
                    img = np.array(Image.open(path))
            else:
                # ---> THE FIX: Use Pillow instead of cv2 to load standard images <---
                # .convert('RGB') automatically handles Grayscale and RGBA (PNGs with transparency)
                pil_img = Image.open(path).convert('RGB')
                img = np.array(pil_img)
                    
            if img is None: return None
            
            # Normalize to 8-bit if it's a 16-bit scientific image
            if img.dtype != np.uint8:
                img = np.uint8(cv2.normalize(img, None, 0, 255, cv2.NORM_MINMAX))
            
            # Handle grayscale scientific images
            if len(img.shape) == 2:
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
                
            return img
        except Exception as e:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Codec Error", f"Failed reading base image data:\n{e}")
            return None
    
    def load_current_image_data(self):
        """Loads the current selected image into the UI."""
        if self.current_index >= len(self.image_files) or not self.image_states: return
        
        state = self.image_states[self.current_index]
        file_path = state['file_path']
        
        try:
            # ---> Loading from disk instead of memory! <---
            self.original_image_rgb = self.load_raw_image_array(file_path)
            
            if self.original_image_rgb is None: return
            
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
            
            # ---> FIXED: INTENSITY DUAL SLIDER <---
            # Keep your clever Otsu thresholding for the minimum value on first load
            if state.get('int_min') is None or state.get('int_min') == 0:
                otsu_val = filters.threshold_otsu(self.cached_gray)
                state['int_min'] = int(otsu_val)
                state['int_max'] = 255 # Default to completely white for max
                
            self.int_slider.set_values(state.get('int_min', 0), state.get('int_max', 255))

            # ---> HUE SLIDER <---
            self.hue_slider.set_values(state.get('hue_min', 0), state.get('hue_max', 179))
            
            # ---> FIXED: AREA DUAL SLIDER <---
            if 'area_min_pos' not in state: state['area_min_pos'] = 30
            if 'area_max_pos' not in state: state['area_max_pos'] = 1000
            
            self.area_slider.set_values(state['area_min_pos'], state['area_max_pos'])

            # ---> NEW: CIRCULARITY SLIDER <---
            if 'circ_min' not in state: state['circ_min'] = 0
            if 'circ_max' not in state: state['circ_max'] = 100
            
            self.circ_slider.set_values(state.get('circ_min', 0))


            self._ignore_sliders = False 

            self.process_image()
            
        except Exception as e:
            import os
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", f"Failed loading {os.path.basename(file_path)}:\n{e}")
    
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
        self.save_state_for_undo() 
        self.is_drawing = True
        self.last_x, self.last_y = event.x, event.y
        
        # Start tracking the lasso points
        orig_x = int((event.x - self.offset_x) * self.scale_x)
        orig_y = int((event.y - self.offset_y) * self.scale_y)
        self.draw_points_img = [(orig_x, orig_y)]

    def draw_motion(self, event):
        if self.is_drawing:
            if not self.auto_detect_enabled:
                self.auto_detect_enabled = True
                self.btn_auto.config(text="Auto Detect: ON", fg="green")
                
            color = "green" if self.draw_mode == "pencil" else "red"
            
            # Draw a thin visual outline on the UI canvas so you can see your lasso
            self.canvas.create_line(self.last_x, self.last_y, event.x, event.y, fill=color, width=2, capstyle=tk.ROUND)
            
            # Record the coordinate for the polygon fill
            orig_x = int((event.x - self.offset_x) * self.scale_x)
            orig_y = int((event.y - self.offset_y) * self.scale_y)
            self.draw_points_img.append((orig_x, orig_y))
                
            self.last_x, self.last_y = event.x, event.y

    def stop_draw(self, event):
        self.is_drawing = False
        
        # Determine which mask we are editing
        mask = self.current_manual_add if self.draw_mode == "pencil" else self.current_manual_remove
        
        if mask is not None and hasattr(self, 'draw_points_img') and len(self.draw_points_img) > 0:
            if len(self.draw_points_img) > 2:
                # If you drew a shape, connect the end to the start and fill the polygon
                pts = np.array([self.draw_points_img], dtype=np.int32)
                cv2.fillPoly(mask, pts, 255)
            else:
                # Fallback: If you just do a single mouse click without dragging, draw a small dot
                cv2.circle(mask, self.draw_points_img[0], radius=15, color=255, thickness=-1)
                
        self.draw_points_img = [] # Reset for the next drawing action
        self.process_image()

    def clear_drawing(self):
        if not self.image_states or self.current_index >= len(self.image_states): return
        self.save_state_for_undo() 
        state = self.image_states[self.current_index]
        if state['manual_mask_add'] is not None: state['manual_mask_add'].fill(0)
        if state['manual_mask_remove'] is not None: state['manual_mask_remove'].fill(0)
        self.process_image()

    def next_image(self):
        if self.current_index < len(self.image_states) - 1:
            self.current_index += 1
            self.load_current_image_data()

    def prev_image(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.load_current_image_data()

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
                area_range = f"{state.get('min_area_actual', 'N/A')}-{state.get('max_area_actual', 'N/A')}"
                
                # Get just the file name
                full_path = state.get('file_path', 'Unknown')
                file_name_only = os.path.basename(full_path) if full_path != 'Unknown' else 'Unknown'
                
                # Normalize the intensity
                raw_intensity = state['stats'].get('mean_intensity', 0)
                normalized_intensity = round((raw_intensity / 255.0) * 100, 2) 

                # Keep only the Normalized Area
                normalized_area = state['stats'].get('area_percentage', 0)

                final_results.append({
                    'File Name': file_name_only,
                    'Normalized Area (%)': normalized_area,
                    'Normalized Intensity (% of White)': normalized_intensity,
                    'Total Clusters': state['stats'].get('cluster_count', 0),
                    'Hue Range Used': f"{state.get('hue_min', 'N/A')}-{state.get('hue_max', 'N/A')}",
                    'Intensity Range Used': intensity_range,
                    'Area Threshold (px)': area_range,
                    'Fiber Stripping (Circularity)': state.get('circ_min', 'N/A'),
                    'Manual Edit Used': used_manual
                })
                
        # 2. Check if we found stats but the final list is still empty
        if not final_results:
            messagebox.showwarning("Incomplete Data", "Images were found, but Auto Detect hasn't been turned on to quantify them yet.")
            return

        # 3. Proceed to save (ONLY ONCE!)
        save_path = filedialog.asksaveasfilename(
            title="Export to Excel",
            defaultextension=".xlsx", 
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if save_path:
            try:
                # pandas must be imported at the top of the file!
                df = pd.DataFrame(final_results)
                df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Data finalized and saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to save to Excel:\n{e}")