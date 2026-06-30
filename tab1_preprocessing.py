import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cv2
import numpy as np
from PIL import Image, ImageTk
import os
import czifile
import tifffile
import platform

class PreProcessingTab(ttk.Frame):
    def __init__(self, parent, main_app):
        # 1. Initialize the parent ttk.Frame cleanly with ONLY the parent widget
        super().__init__(parent)
        
        # 2. Save custom application references safely
        self.main_app = main_app
        self.os_type = platform.system()
        
        # 3. Data Volumes
        self.original_raw_volume = None # Keeps the uncropped backup
        self.raw_volume = None      
        self.original_filename = ""
        self.max_z = 0
        self.is_merged_preview = False
        self.channel_baselines = [] 
        
        # 4. 16-bit Adjustments (Global Only)
        self.adj_settings = {'contrast': 1.0, 'brightness': 0.0}
        
        # 5. Cropping and Image View Variables
        self.rect_id = None
        self.start_x = None
        self.start_y = None
        self.current_rect = None
        
        # 6. View State: Coordinates of the image top-left corner relative to canvas
        self.img_offset_x = 0
        self.img_offset_y = 0
        self.img_scale = 1.0  # Dynamic zoom level factor
        
        # 7. Panning State: Mouse positioning caches for dragging actions
        self.pan_start_x = 0
        self.pan_start_y = 0
        
        # 8. Image References (Prevents Tkinter garbage collection)
        self.current_pil_image = None
        self.tk_img = None  # Matches self.tk_img used in update_preview / redraw_image

        # 9. UI Orchestration
        self.setup_ui()
        if hasattr(self, 'maximize_window'):
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
        tk.Button(action_frame, text="1. Load Microscopy Images", command=self.load_files, font=("Arial", 11, "bold"), bg="#4a90e2", fg="white").pack(fill=tk.X)
        
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

        tk.Button(action_frame, text="2. Save Processed Image As...", command=self.save_image_to_disk, font=("Arial", 11, "bold"), bg="#2e7d32", fg="white").pack(fill=tk.X, pady=5)

        # Z-Navigation
        nav_frame = tk.LabelFrame(control_frame, text="Stack Preview Navigation", padx=5, pady=2)
        nav_frame.pack(fill=tk.X, pady=2)
        
        self.lbl_z_current = tk.Label(nav_frame, text="Current Stack: 0")
        self.lbl_z_current.pack()
        
            # Link the slider to our interceptor function
        self.scale_z = tk.Scale(nav_frame, from_=0, to=0, orient=tk.HORIZONTAL, showvalue=0, command=self.on_z_slider_move)
        self.scale_z.pack(fill=tk.X)

        # Merge Range (Z-Projection) 
        proj_frame = tk.LabelFrame(control_frame, text="Z-Projection Merge Range", padx=5, pady=2)
        proj_frame.pack(fill=tk.X, pady=2)
        
        tk.Label(proj_frame, text="Start:").grid(row=0, column=0, sticky="e", pady=2)
        self.spin_z_start = tk.Spinbox(proj_frame, from_=0, to=0, width=4, command=self.update_preview)
        self.spin_z_start.grid(row=0, column=1, pady=2, padx=2)
        
        tk.Label(proj_frame, text="End:").grid(row=1, column=0, sticky="e", pady=2)
        self.spin_z_end = tk.Spinbox(proj_frame, from_=0, to=0, width=4, command=self.update_preview)
        self.spin_z_end.grid(row=1, column=1, pady=2, padx=2)

        # --- NEW: Z-Projection Method Dropdown ---
        tk.Label(proj_frame, text="Method:").grid(row=2, column=0, sticky="e", pady=2)
        self.combo_proj_method = ttk.Combobox(proj_frame, values=["Max IP", "Mean IP", "Sum IP", "Min IP"], state="readonly", width=8)
        self.combo_proj_method.grid(row=2, column=1, pady=2, padx=2)
        self.combo_proj_method.set("Max IP")
        self.combo_proj_method.bind("<<ComboboxSelected>>", self.update_preview)
        
        # Button takes up 3 rows now on the right side
        self.btn_preview_merge = tk.Button(proj_frame, text="👁 Preview\nMerge", command=self.toggle_merge_preview)
        self.btn_preview_merge.grid(row=0, column=2, rowspan=3, padx=(5, 0), sticky="nsew")
        proj_frame.grid_columnconfigure(2, weight=1)

        # ---------------------------------------------------------
        # Channel Visibility & Pseudo-Color (Dynamic Containers)
        # ---------------------------------------------------------
        self.chan_frame = tk.LabelFrame(control_frame, text="Channel Visibility & Pseudo-Color", padx=10, pady=5)
        self.chan_frame.pack(fill=tk.X, pady=2)
        
        # Dictionaries to hold N-channel variables dynamically
        self.channel_vars = {}    # Maps name -> tk.BooleanVar()
        self.channel_colors = {}  # Maps name -> (R, G, B) tuple
        self.channel_widgets = [] # Tracks UI rows so we can clear them on new image load
        
        # ---------------------------------------------------------
        # Channel Adjustments (Compact)
        # ---------------------------------------------------------
        adj_frame = tk.LabelFrame(control_frame, text="Channel Adjustments", padx=10, pady=5)
        adj_frame.pack(fill=tk.X, pady=2)

        self.var_apply_adjustments = tk.BooleanVar(value=True)
        tk.Checkbutton(adj_frame, text="Apply Adjustments & Range Expansion", variable=self.var_apply_adjustments, command=self.update_preview).pack(fill=tk.X, pady=(0,5))
        
        # Data dictionary for dynamic sliders
        self.adj_data = {} 
        self.active_adj_channel = tk.StringVar()
        self._is_updating_ui = False 
        
        # Dropdown (Values will be populated dynamically)
        self.combo_channel = ttk.Combobox(adj_frame, textvariable=self.active_adj_channel, state="readonly")
        self.combo_channel.pack(fill=tk.X, pady=(0, 5))
        self.combo_channel.bind("<<ComboboxSelected>>", self.on_adj_channel_change)
        
        # Sub-frame for sliders
        slider_frame = tk.Frame(adj_frame)
        slider_frame.pack(fill=tk.X)
        slider_frame.columnconfigure(1, weight=1)
        slider_frame.columnconfigure(3, weight=1)
        
        # Contrast Slider
        tk.Label(slider_frame, text="C:", font=("Arial", 9, "bold"), fg="gray").grid(row=0, column=0, sticky="e")
        self.scale_contrast = tk.Scale(slider_frame, from_=0.0, to=5.0, resolution=0.05, orient=tk.HORIZONTAL, 
                                       command=self.on_shared_slider_move, width=10, sliderlength=15)
        self.scale_contrast.grid(row=0, column=1, sticky="ew", padx=(0, 5))
        
        # Brightness Slider
        tk.Label(slider_frame, text="B:", font=("Arial", 9, "bold"), fg="gray").grid(row=0, column=2, sticky="e")
        self.scale_brightness = tk.Scale(slider_frame, from_=-1.0, to=1.0, resolution=0.05, orient=tk.HORIZONTAL, 
                                         command=self.on_shared_slider_move, width=10, sliderlength=15)
        self.scale_brightness.grid(row=0, column=3, sticky="ew", padx=(0, 2))
        
        # Call this once during init to set up defaults (optional, can wait until image load)
        self.build_dynamic_channels([
            {"name": "Red (Alexa 568)", "hex": "#FF0000", "rgb": (255, 0, 0)},
            {"name": "Green (Alexa 488)", "hex": "#00FF00", "rgb": (0, 255, 0)},
            {"name": "Blue (DAPI)", "hex": "#0000FF", "rgb": (0, 0, 255)}
        ])

        # ---------------------------------------------------------
        # --- Scale Bar Overlay ---
        # ---------------------------------------------------------
        scale_frame = tk.LabelFrame(control_frame, text="Scale Bar Overlay", padx=10, pady=5)
        scale_frame.pack(fill=tk.X, pady=2)
        
        self.var_show_scalebar = tk.BooleanVar(value=True)
        tk.Checkbutton(scale_frame, text="Show Scale Bar", variable=self.var_show_scalebar, 
                       command=self.update_preview).pack(anchor="w")

        # The new More Options button
        tk.Button(scale_frame, text="More Options", command=self.open_scale_bar_options).pack(fill=tk.X, pady=5)

        # --- HIDDEN VARIABLES FOR BACKGROUND CALIBRATION ---
        # Kept in a hidden frame so the metadata extractor and draw_scale_bar() 
        # can still operate perfectly without cluttering the main UI.
        self.hidden_sb_frame = tk.Frame(self)
        
        self.entry_pixel_size = tk.Entry(self.hidden_sb_frame)
        self.entry_pixel_size.insert(0, "0.5") 
        
        self.entry_sb_width = tk.Entry(self.hidden_sb_frame)
        self.entry_sb_width.insert(0, "100") 
        
        self.spin_sb_thick = tk.Spinbox(self.hidden_sb_frame, from_=1, to=100)
        self.spin_sb_thick.delete(0, tk.END)
        self.spin_sb_thick.insert(0, "2")
        
        self.spin_sb_font = tk.Spinbox(self.hidden_sb_frame, from_=0.1, to=5.0, increment=0.1)
        self.spin_sb_font.delete(0, tk.END)
        self.spin_sb_font.insert(0, "0.5")
        
        self.combo_sb_color = ttk.Combobox(self.hidden_sb_frame, values=["White", "Black", "Red", "Green", "Blue", "Yellow"])
        self.combo_sb_color.set("White")
        
        # ---> NEW: Hidden combobox to store position <---
        self.combo_sb_position = ttk.Combobox(self.hidden_sb_frame, values=["Bottom Right", "Bottom Left", "Top Right", "Top Left"])
        self.combo_sb_position.set("Bottom Left")
        # ---------------------------------------------------

        # Right Panel: Canvas
        self.canvas_frame = tk.Frame(self, bg="black")
        self.canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.canvas = tk.Canvas(self.canvas_frame, bg="black", cursor="crosshair")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Mouse Bindings for Cropping (Left Click)
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_release)

        # TRACKPAD PINCH-TO-ZOOM BINDINGS (Windows / Linux / macOS)
        self.canvas.bind("<MouseWheel>", self.on_zoom)                     # Mouse wheels & macOS Pinch
        self.canvas.bind("<Control-MouseWheel>", self.on_zoom)             # Windows Trackpad Pinch Gesture
        self.canvas.bind("<Button-4>", self.on_zoom)                       # Linux Scroll Up
        self.canvas.bind("<Button-5>", self.on_zoom)                       # Linux Scroll Down

        # TRACKPAD PANNING BINDINGS (Right-Click Drag)
        self.canvas.bind("<ButtonPress-3>", self.on_pan_start)             # Right-Click Press
        self.canvas.bind("<B3-Motion>", self.on_pan_drag)                   # Right-Click Drag Move
        self.canvas.bind("<ButtonRelease-3>", self.on_pan_end)             # Restore cursor on release

        # NEW: TWO-FINGER TRACKPAD PANNING BINDINGS (Windows / macOS)
        self.canvas.bind("<MouseWheel>", self.on_trackpad_pan, add="+")       # Vertical Trackpad Pan
        self.canvas.bind("<Shift-MouseWheel>", self.on_trackpad_pan)          # Horizontal Trackpad Pan
        
        # Double click left mouse button (or double tap trackpad) to instantly reset view
        self.canvas.bind("<Double-Button-1>", self.reset_view_layout)       # Center Image to the Window with Double Click


        # # --- Keyboard bindings for Left/Right arrows ---
        # top = self.winfo_toplevel()
        # top.bind("<Left>", lambda e: self.prev_image() if self.btn_prev_img['state'] == tk.NORMAL else None)
        # top.bind("<Right>", lambda e: self.next_image() if self.btn_next_img['state'] == tk.NORMAL else None)

# --- Preview zoom and pan ---
    def on_zoom(self, event):
        """Handles zoom gestures, preventing zooming out past full-window fit."""
        """Processes zooming only if Control key is held, otherwise passes event to trackpad pan."""
        # 0x0004 represents the Control key state flag in Tkinter
        if not (event.state & 0x0004):
            # No control key? This is a normal trackpad pan gesture!
            self.on_trackpad_pan(event)
            return
        # --- THE FIX: Look for the active raw matrix slice instead of the missing PIL image ---
        if not hasattr(self, 'active_raw_slice') or self.active_raw_slice is None:
            return

        # Determine zoom direction step vector
        if event.num == 4:
            zoom_factor = 1.1
        elif event.num == 5:
            zoom_factor = 0.9
        elif event.delta != 0:
            zoom_factor = 1.1 if event.delta > 0 else 0.9
        else:
            return

        # Calculate the base scale needed to fit the image perfectly on screen
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: 
            canvas_w, canvas_h = 800, 600

        # --- THE FIX: Get original size from the raw matrix shape [Height, Width] ---
        orig_h, orig_w = self.active_raw_slice.shape[:2]
        fit_scale = min(canvas_w / orig_w, canvas_h / orig_h)
        
        # Calculate what the new scale would be
        new_scale = self.img_scale * zoom_factor

        # RECENTER TRIGGER: If zooming out drops below or matches the window fit scale
        if zoom_factor < 1.0 and new_scale <= (fit_scale + 0.01):
            # Snap back to centered overview layout
            self.img_scale = fit_scale
            self.zoom_scale = fit_scale  # Keep synced with master state register
            self.scale_x = fit_scale
            self.img_offset_x = (canvas_w - int(orig_w * fit_scale)) // 2
            self.img_offset_y = (canvas_h - int(orig_h * fit_scale)) // 2
            self.img_x = self.img_offset_x
            self.img_y = self.img_offset_y
        else:
            # Enforce hard constraints (No zooming out past fit_scale, max 15x zoom)
            if new_scale < fit_scale:
                new_scale = fit_scale
            if new_scale > 15.0:
                return

            # Keep your standard focal point adjustment when zooming in
            canvas_x = self.canvas.canvasx(event.x)
            canvas_y = self.canvas.canvasy(event.y)

            # Re-adjust factor to account for any boundary capping
            actual_factor = new_scale / self.img_scale
            self.img_offset_x = canvas_x - (canvas_x - self.img_offset_x) * actual_factor
            self.img_offset_y = canvas_y - (canvas_y - self.img_offset_y) * actual_factor
            
            self.img_scale = new_scale
            # Keep master state variables explicitly synced
            # (Inside on_zoom...)
            self.zoom_scale = new_scale
            self.img_x = self.img_offset_x
            self.img_y = self.img_offset_y

        self.redraw_image() 

    def reset_view_layout(self, event=None):
        """Instantly resets zoom factor and centers the image on the canvas."""
        # --- THE FIX: Match structural verification properties ---
        if not hasattr(self, 'active_raw_slice') or self.active_raw_slice is None:
            return
            
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: 
            canvas_w, canvas_h = 800, 600
            
        # --- THE FIX: Extract from array shape ---
        orig_h, orig_w = self.active_raw_slice.shape[:2]
        fit_scale = min(canvas_w / orig_w, canvas_h / orig_h)
        
        self.img_scale = fit_scale
        self.zoom_scale = fit_scale
        self.scale_x = fit_scale
        self.img_offset_x = (canvas_w - int(orig_w * fit_scale)) // 2
        self.img_offset_y = (canvas_h - int(orig_h * fit_scale)) // 2
        # (Inside reset_view_layout...)
        self.img_x = self.img_offset_x
        self.img_y = self.img_offset_y
        
        self.update_preview()  # Changed from self.redraw_image()

    def on_pan_start(self, event):
        """Initializes coordinates for drag-panning and changes cursor shape."""
        self.pan_start_x = event.x
        self.pan_start_y = event.y
        self.canvas.config(cursor="fleur")  # Visual anchor update for panning

    def on_pan_drag(self, event):
        """Updates image offsets dynamically during a drag movement."""
        dx = event.x - self.pan_start_x
        dy = event.y - self.pan_start_y

        self.img_offset_x += dx
        self.img_offset_y += dy
        
        self.img_x = self.img_offset_x
        self.img_y = self.img_offset_y

        self.pan_start_x = event.x
        self.pan_start_y = event.y

        self.redraw_image()  

    def on_pan_end(self, event):
        """Restores the standard crop crosshair cursor when right-click pan drag ends."""
        self.canvas.config(cursor="crosshair")

    def on_trackpad_pan(self, event):
        """Enables native two-finger trackpad swipe-to-pan for Windows and macOS."""
        if not hasattr(self, 'current_pil_image') or self.current_pil_image is None:
            return

        # Skip panning if Control modifier is active to prevent conflicting jitter
        if event.state & 0x0004:  
            return
            
        if hasattr(event, 'delta') and event.delta != 0:
            pan_step = 20 if event.delta > 0 else -20
            
            # Check if Shift key is pressed (Horizontal Scroll)
            if event.state & 0x0001:  
                self.img_offset_x += pan_step
            else:
                # Vertical Scroll
                if hasattr(self, 'os_type') and self.os_type == "Darwin":
                    self.img_offset_y += event.delta
                else:
                    self.img_offset_y += pan_step

            # (Inside on_trackpad_pan...)
            self.img_x = self.img_offset_x
            self.img_y = self.img_offset_y
            
            self.redraw_image()  

    def redraw_image(self):
        """Resamples the base image and draws it at the active scale and offsets."""
        if not hasattr(self, 'current_pil_image') or self.current_pil_image is None:
            return

        self.canvas.delete("all")

        # Keep everything explicitly tied together to support tracking dependencies
        self.zoom_scale = self.img_scale
        self.scale_x = self.img_scale
        self.img_x = self.img_offset_x
        self.img_y = self.img_offset_y

        # Calculate new dynamic size vectors
        orig_w, orig_h = self.current_pil_image.size
        new_w = max(1, int(orig_w * self.img_scale))
        new_h = max(1, int(orig_h * self.img_scale))

        # Use NEAREST resampling for smooth, lag-free rendering during trackpad movements
        resized_img = self.current_pil_image.resize((new_w, new_h), Image.Resampling.NEAREST)
        
        # Saves the reference to self.tk_img to prevent canvas garbage collection bugs
        self.tk_img = ImageTk.PhotoImage(resized_img)

        # Place image on canvas using our dynamic offset registers
        self.canvas.create_image(self.img_offset_x, self.img_offset_y, anchor=tk.NW, image=self.tk_img)
        self.rect_id = None

        # Redraw crop rectangle if it exists, converting image space to zoomed canvas space
        if hasattr(self, 'current_rect') and self.current_rect:
            x1, y1, x2, y2 = self.current_rect
            cx1 = x1 * self.img_scale + self.img_offset_x
            cy1 = y1 * self.img_scale + self.img_offset_y
            cx2 = x2 * self.img_scale + self.img_offset_x
            cy2 = y2 * self.img_scale + self.img_offset_y
            self.rect_id = self.canvas.create_rectangle(cx1, cy1, cx2, cy2, outline="red", width=2)

        # Render scale bar overlay layer dynamically
        if hasattr(self, 'draw_scale_bar'):
            self.draw_scale_bar()

    # --- Channel Adjustments ---
    def on_adj_channel_change(self, event=None):
        """Updates the sliders to reflect the saved values for the newly selected channel."""
        channel = self.active_adj_channel.get()
        vals = self.adj_data[channel]
        
        # Temporarily block the slider command so it doesn't trigger an image redraw
        # while we are just visually moving the sliders to match the saved data
        self._is_updating_ui = True
        self.scale_contrast.set(vals["c"])
        self.scale_brightness.set(vals["b"])
        self._is_updating_ui = False

    def on_shared_slider_move(self, val=None):
        """Saves the current slider values to the active channel and triggers an image update."""
        if self._is_updating_ui: 
            return
            
        channel = self.active_adj_channel.get()
        self.adj_data[channel]["c"] = self.scale_contrast.get()
        self.adj_data[channel]["b"] = self.scale_brightness.get()
        
        # Call your existing image update function!
        # (Change this to self.update_preview() if that's what Tab 1 uses)
        self.on_slider_move()

    def build_dynamic_channels(self, channel_list):
        """
        Dynamically rebuilds the UI for N channels. 
        Expects a list of dicts: [{'name': 'DAPI', 'hex': '#0000FF', 'rgb': (0,0,255)}, ...]
        """
        # 1. Clear existing UI rows
        for widget in self.channel_widgets:
            widget.destroy()
        self.channel_widgets.clear()
        
        self.channel_vars.clear()
        self.channel_colors.clear()
        self.adj_data.clear()
        
        channel_names = []

        # 2. Build new UI rows
        for ch in channel_list:
            ch_name = ch["name"]
            hex_col = ch["hex"]
            
            # Setup State Variables
            self.channel_vars[ch_name] = tk.BooleanVar(value=True)
            self.channel_colors[ch_name] = ch["rgb"]
            self.adj_data[ch_name] = {"c": 1.0, "b": 0.0}
            channel_names.append(ch_name)
            
            # Create the UI Row (Modify this if your create_channel_row returns different things)
            # We append the parent frame/widgets to self.channel_widgets so we can destroy them later
            row_frame = tk.Frame(self.chan_frame)
            row_frame.pack(fill=tk.X, pady=2)
            self.channel_widgets.append(row_frame)
            
            # Visibility Checkbox
            chk = tk.Checkbutton(row_frame, text=ch_name, variable=self.channel_vars[ch_name], command=self.update_preview)
            chk.pack(side=tk.LEFT)
            
            # Color Button (Assuming you want to change it later)
            btn_color = tk.Button(row_frame, bg=hex_col, width=2)
            btn_color.pack(side=tk.RIGHT, padx=5)
            
        # 3. Update Adjustments Combobox
        self.combo_channel.config(values=channel_names)
        if channel_names:
            self.combo_channel.current(0)
            self.on_adj_channel_change()

    # --- Pseudo Coloring Color Picker ---
    def pick_color(self, channel_id):
        """Opens a color picker and updates the specific channel's color block."""
        from tkinter import colorchooser
        
        initial = {"R": self.color_r, "G": self.color_g, "B": self.color_b}[channel_id]
        color_result = colorchooser.askcolor(title=f"Select Color for Channel {channel_id}", color=initial)
        
        if color_result[0] is not None:
            rgb_tuple = tuple(int(c) for c in color_result[0])
            hex_color = color_result[1]
            
            # Only update the button background now!
            if channel_id == "R":
                self.color_r = rgb_tuple
                self.btn_color_r.config(bg=hex_color)
            elif channel_id == "G":
                self.color_g = rgb_tuple
                self.btn_color_g.config(bg=hex_color)
            elif channel_id == "B":
                self.color_b = rgb_tuple
                self.btn_color_b.config(bg=hex_color)
                
            self.update_preview()

    def create_channel_row(self, parent, text, var, channel_id, default_hex):
        """Helper to build a clean row with Checkbox + Color Button."""
        import tkinter as tk
        row = tk.Frame(parent)
        row.pack(fill=tk.X, pady=2)
        
        cb = tk.Checkbutton(row, variable=var, command=self.update_preview)
        cb.pack(side=tk.LEFT)
        
        # Text remains standard black
        lbl = tk.Label(row, text=text, fg="black")
        lbl.pack(side=tk.LEFT)
        
        # The color block does all the visual communication
        btn = tk.Button(row, bg=default_hex, width=3, relief="raised", cursor="hand2", 
                        command=lambda: self.pick_color(channel_id))
        btn.pack(side=tk.RIGHT, padx=5)
        
        return btn, lbl

    # --- Scale bar functions ---
    def open_scale_bar_options(self):
        """Opens a pop-up window for user-friendly scale bar customization."""
        opts = tk.Toplevel(self)
        opts.title("Scale Bar Settings")
        # Increased height from 220 to 260 to fit the new row
        opts.geometry("280x260") 
        opts.attributes('-topmost', True) # Keeps the pop-up above the main app
        opts.resizable(False, False)
        
        # Helper function to generate clean rows in the pop-up
        def make_row(parent, label_text, widget_class, **kwargs):
            frame = tk.Frame(parent)
            frame.pack(fill=tk.X, padx=15, pady=5)
            tk.Label(frame, text=label_text).pack(side=tk.LEFT)
            widget = widget_class(frame, **kwargs)
            widget.pack(side=tk.RIGHT)
            return widget

        # Create the UI inputs and populate them with the current hidden values
        ent_width = make_row(opts, "Length (\u03BCm):", tk.Entry, width=12)
        ent_width.insert(0, self.entry_sb_width.get())
        
        spn_thick = make_row(opts, "Thickness (px):", tk.Spinbox, from_=1, to=100, width=10)
        spn_thick.delete(0, tk.END); spn_thick.insert(0, self.spin_sb_thick.get())
        
        spn_font = make_row(opts, "Font Scale:", tk.Spinbox, from_=0.1, to=5.0, increment=0.1, width=10)
        spn_font.delete(0, tk.END); spn_font.insert(0, self.spin_sb_font.get())
        
        cmb_color = make_row(opts, "Color:", ttk.Combobox, values=["White", "Black", "Red", "Green", "Blue", "Yellow"], width=10)
        cmb_color.set(self.combo_sb_color.get())
        
        # ---> NEW: Position Dropdown <---
        cmb_position = make_row(opts, "Position:", ttk.Combobox, values=["Bottom Right", "Bottom Left", "Top Right", "Top Left"], state="readonly", width=12)
        cmb_position.set(self.combo_sb_position.get())
        
        def apply_options():
            # 1. Save pop-up values back to our hidden persistent widgets
            self.entry_sb_width.delete(0, tk.END)
            self.entry_sb_width.insert(0, ent_width.get())
            
            self.spin_sb_thick.delete(0, tk.END)
            self.spin_sb_thick.insert(0, spn_thick.get())
            
            self.spin_sb_font.delete(0, tk.END)
            self.spin_sb_font.insert(0, spn_font.get())
            
            self.combo_sb_color.set(cmb_color.get())
            
            # ---> NEW: Save Position to hidden variable <---
            self.combo_sb_position.set(cmb_position.get())
            
            # 2. Trigger the redraw and close the pop-up
            self.update_preview()
            opts.destroy()
            
        tk.Button(opts, text="Apply & Close", command=apply_options, bg="#4CAF50", fg="white").pack(pady=15)

    def draw_scale_bar(self):
        """Draws a floating Tkinter scale bar strictly calibrated to user inputs."""
        # 1. Clear any existing scale bar
        self.canvas.delete("scalebar")
        
        # 2. Check if user toggled it off
        if not getattr(self, 'var_show_scalebar', None) or not self.var_show_scalebar.get():
            return
            
        # 3. Safely get all manual User Inputs
        try:
            pixel_size_um = float(self.entry_pixel_size.get())
            user_width_um = float(self.entry_sb_width.get())
            thickness = int(self.spin_sb_thick.get())
            font_size = int(float(self.spin_sb_font.get()) * 14) # Convert to Tkinter font scale
            color_name = self.combo_sb_color.get()
            
            # ---> NEW: Fetch the position safely <---
            position = "Bottom Left" # Default fallback
            if hasattr(self, 'combo_sb_position'):
                position = self.combo_sb_position.get()
                
        except (ValueError, AttributeError):
            return # Abort if UI inputs are empty or invalid

        if pixel_size_um <= 0 or user_width_um <= 0:
            return
            
        # 4. --- THE IMAGEJ CALIBRATION MATH ---
        # Calculate how many raw image pixels equal the user's requested physical width
        raw_pixels = user_width_um / pixel_size_um
        
        # Adjust for the Tkinter canvas shrinking/expanding the image on your monitor
        scale_x = getattr(self, 'scale_x', 1.0) or 1.0
        actual_screen_pixels = raw_pixels * scale_x
        
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        # Prevent math errors if canvas hasn't fully rendered yet
        if canvas_w <= 1 or canvas_h <= 1:
            return

        margin_x, margin_y = 30, 30

        # ---> NEW: Position Logic <---
        # Determine X coordinates
        if "Left" in position:
            x1 = margin_x
            x2 = margin_x + actual_screen_pixels
        else: # Right
            x1 = canvas_w - margin_x - actual_screen_pixels
            x2 = canvas_w - margin_x

        # Determine Y coordinates
        text_offset = font_size + 5 # Dynamic spacing based on font size
        if "Top" in position:
            y1 = margin_y + text_offset
            y2 = y1 + thickness
        else: # Bottom
            y1 = canvas_h - margin_y - thickness
            y2 = canvas_h - margin_y

        text_x = x1 + (actual_screen_pixels / 2)
        text_y = y1 - (font_size / 2) - 4

        # Format text to remove .0 if it's a whole number (e.g., 50.0 -> 50)
        text_val = int(user_width_um) if float(user_width_um).is_integer() else round(user_width_um, 2)
        text = f"{text_val} \u03BCm" 

        # Map UI color dropdown to Tkinter hex/color names
        color_map = {
            "White": "white", "Black": "black", "Red": "red", 
            "Green": "#00FF00", "Blue": "blue", "Yellow": "yellow"
        }
        bar_color = color_map.get(color_name, "white")
        outline_color = "black" if bar_color in ["white", "yellow", "#00FF00"] else "white"

        # Draw Shadow/Outline (for visibility against light/dark backgrounds)
        self.canvas.create_rectangle(x1-2, y1-2, x2+2, y2+2, fill=outline_color, outline=outline_color, tags="scalebar")
        for dx, dy in [(-1,-1), (-1,1), (1,-1), (1,1)]:
            self.canvas.create_text(text_x+dx, text_y+dy, text=text, fill=outline_color, font=("Arial", font_size, "bold"), tags="scalebar")
            
        # Draw Foreground
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=bar_color, outline=bar_color, tags="scalebar")
        self.canvas.create_text(text_x, text_y, text=text, fill=bar_color, font=("Arial", font_size, "bold"), tags="scalebar")

    def stamp_scale_bar_for_export(self, image_rgb):
        """Burns a physical scale bar into the numpy image array using OpenCV at a specified position."""
        try:
            # 1. ABORT IF SCALE BAR IS UNCHECKED
            if not getattr(self, 'var_show_scalebar', None) or not self.var_show_scalebar.get():
                return image_rgb 
                
            # 2. FETCH ALL USER SETTINGS FROM UI
            try:
                pixel_size_um = float(self.entry_pixel_size.get())
                user_width_um = float(self.entry_sb_width.get())
                bar_thickness = int(self.spin_sb_thick.get())
                font_scale = float(self.spin_sb_font.get())
                color_name = self.combo_sb_color.get()
                position = self.combo_sb_position.get()
            except (ValueError, AttributeError):
                return image_rgb 

            if pixel_size_um <= 0 or user_width_um <= 0: 
                return image_rgb
                
            import cv2
            import numpy as np

            img_h, img_w = image_rgb.shape[:2]
            
            # 3. CALCULATE EXACT PIXEL LENGTH (No more 15% auto-math)
            bar_length_px = int(user_width_um / pixel_size_um)
            margin = int(max(10, img_w * 0.02)) 
            
            # Format text (Note: OpenCV doesn't render the Greek µ symbol well, so we use 'um')
            text_val = int(user_width_um) if float(user_width_um).is_integer() else round(user_width_um, 2)
            text = f"{text_val} um"
            
            # Setup Font
            font = cv2.FONT_HERSHEY_SIMPLEX
            cv_font_scale = font_scale * 2.0 # Scaled up slightly to match Tkinter proportions
            text_thickness = max(1, int(cv_font_scale * 2))
            (text_w, text_h), _ = cv2.getTextSize(text, font, cv_font_scale, text_thickness)
            
            # 4. COLOR MAPPING (RGB format)
            color_map = {
                "White": (255, 255, 255), "Black": (0, 0, 0), "Red": (255, 0, 0), 
                "Green": (0, 255, 0), "Blue": (0, 0, 255), "Yellow": (255, 255, 0)
            }
            bar_color = color_map.get(color_name, (255, 255, 255))
            outline_color = (0, 0, 0) if color_name in ["White", "Yellow", "Green"] else (255, 255, 255)

            # 5. DETERMINE X/Y COORDINATES
            if "Left" in position:
                bar_x = margin
            else: # Right
                bar_x = img_w - margin - bar_length_px
                
            if "Top" in position:
                bar_y = margin + int(text_h * 1.5)
            else: # Bottom
                bar_y = img_h - margin - bar_thickness

            start_point = (bar_x, bar_y)
            end_point = (bar_x + bar_length_px, bar_y + bar_thickness)
            
            # Draw Outline/Shadow for bar
            cv2.rectangle(image_rgb, (bar_x-1, bar_y-1), (end_point[0]+1, end_point[1]+1), outline_color, -1)
            # Draw Foreground bar
            cv2.rectangle(image_rgb, start_point, end_point, bar_color, -1) 
            
            # Center the text over the bar dynamically
            text_x = bar_x + (bar_length_px // 2) - (text_w // 2)
            text_y = bar_y - int(text_h * 0.5)
            
            # Draw Text shadow/outline
            cv2.putText(image_rgb, text, (text_x, text_y), font, cv_font_scale, outline_color, text_thickness + 2, cv2.LINE_AA)
            # Draw Text foreground
            cv2.putText(image_rgb, text, (text_x, text_y), font, cv_font_scale, bar_color, text_thickness, cv2.LINE_AA)
            
            return image_rgb
            
        except Exception as e:
            print(f"Failed to burn scale bar: {e}")
            return image_rgb

    # --- Mouse Events ---
    def on_mouse_press(self, event):
        # Record starting point in absolute canvas space
        self.start_canvas_x = event.x
        self.start_canvas_y = event.y
        
        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.delete(self.rect_id)
            
        self.rect_id = self.canvas.create_rectangle(
            self.start_canvas_x, self.start_canvas_y, 
            self.start_canvas_x, self.start_canvas_y, 
            outline="yellow", dash=(4, 4), width=2
        )

    def on_mouse_drag(self, event):
        if hasattr(self, 'rect_id') and self.rect_id:
            self.canvas.coords(self.rect_id, self.start_canvas_x, self.start_canvas_y, event.x, event.y)

    def on_mouse_release(self, event):
        if hasattr(self, 'rect_id') and self.rect_id:
            # Convert canvas display coordinates back to the underlying RAW file pixel indices
            raw_x1 = (self.start_canvas_x - self.img_x) / self.zoom_scale
            raw_y1 = (self.start_canvas_y - self.img_y) / self.zoom_scale
            raw_x2 = (event.x - self.img_x) / self.zoom_scale
            raw_y2 = (event.y - self.img_y) / self.zoom_scale
            
            # Keep values within true ordered bounds
            x1, x2 = min(raw_x1, raw_x2), max(raw_x1, raw_x2)
            y1, y2 = min(raw_y1, raw_y2), max(raw_y1, raw_y2)
            
            self.current_rect = (x1, y1, x2, y2)

    # --- Adjustment Logic ---
    def on_slider_move(self, event=None):
        """Triggers a visual update when sliders are moved."""
        self.update_preview()

    def toggle_merge_preview(self):
        self.is_merged_preview = not getattr(self, 'is_merged_preview', False)
        
        if self.is_merged_preview:
            self.btn_preview_merge.config(text="Back to\nSingle Stack")
        else:
            self.btn_preview_merge.config(text="👁 Preview\nMerge")
            
        self.update_preview()

    # --- Loading ---
    def load_files(self):
        """Opens a file dialog, unpacks multi-series files (like LIF), and initializes loading."""
        file_paths = filedialog.askopenfilenames(
            title="Select Microscopy Images (CZI, LIF, TIFF)", 
            filetypes=[("Microscopy Files", "*.czi *.lif *.tif *.tiff")]
        )
        if not file_paths: return

        self.lbl_filename.config(text="Scanning containers...")
        self.update()

        # Instead of just storing paths, we store dictionaries with the series index
        self.loaded_files = []
        
        for path in file_paths:
            if path.lower().endswith('.lif'):
                try:
                    from readlif.reader import LifFile
                    lif = LifFile(path)
                    # Add every series inside the LIF as its own "image" in the playlist
                    for i in range(lif.num_images):
                        img_name = lif.get_image(i).name
                        display_name = f"{os.path.basename(path)} [{img_name}]"
                        self.loaded_files.append({'path': path, 'series': i, 'name': display_name})
                except Exception as e:
                    print(f"Failed to scan LIF {path}: {e}")
            else:
                # Standard files default to series 0
                self.loaded_files.append({'path': path, 'series': 0, 'name': os.path.basename(path)})

        if not self.loaded_files:
            self.lbl_filename.config(text="No valid files loaded.")
            return

        self.current_file_index = 0
        self.load_image_from_index()

    def prev_image(self):
        if hasattr(self, 'current_file_index') and self.current_file_index > 0:
            self.current_file_index -= 1
            self.load_image_from_index()

    def next_image(self):
        if hasattr(self, 'current_file_index') and self.current_file_index < len(self.loaded_files) - 1:
            self.current_file_index += 1
            self.load_image_from_index()

    def map_channels_from_xml(self, channels_metadata):
        """Maps raw indices to R, G, B based on emission wavelengths or Zeiss Color tags."""
        self.czi_channel_map = {'R': None, 'G': None, 'B': None}
        
        for idx, ch in enumerate(channels_metadata):
            if idx >= self.original_num_channels: break # Safety limit
            
            # Grab data using the keys we defined in our new extractor
            wave_str = ch.get('Wavelength', 'N/A')
            color_hex = ch.get('Color', 'Unknown').upper()
            
            mapped = False

            # 1. Try mapping by Wavelength first (Your original logic)
            if wave_str != 'N/A':
                try:
                    wave = float(wave_str)
                    # Convert to nanometers if saved in meters
                    if 0 < wave < 1.0: 
                        wave *= 1e9 
                    
                    # Strict wavelength boundaries
                    if wave < 480: 
                        self.czi_channel_map['B'] = idx
                        mapped = True
                    elif 480 <= wave < 550: 
                        self.czi_channel_map['G'] = idx
                        mapped = True
                    elif wave >= 550: 
                        self.czi_channel_map['R'] = idx
                        mapped = True
                except ValueError:
                    pass

            # 2. Smart Failsafe: Use Zeiss Hex Color if Wavelength failed
            # Format is usually #AARRGGBB
            if not mapped and color_hex.startswith('#') and len(color_hex) >= 9:
                try:
                    r_val = int(color_hex[3:5], 16)
                    g_val = int(color_hex[5:7], 16)
                    b_val = int(color_hex[7:9], 16)
                    
                    # Map based on the dominant color in the hex code
                    if b_val > r_val and b_val > g_val:
                        self.czi_channel_map['B'] = idx
                    elif g_val > r_val and g_val > b_val:
                        self.czi_channel_map['G'] = idx
                    elif r_val > g_val and r_val > b_val:
                        self.czi_channel_map['R'] = idx
                except ValueError:
                    pass

        # 3. Final Failsafe: Sequential fill if both wave and color are missing
        mapped_indices = [v for v in self.czi_channel_map.values() if v is not None]
        unmapped_indices = [i for i in range(self.original_num_channels) if i not in mapped_indices]
        
        for color_key in ['R', 'G', 'B']:
            if self.czi_channel_map[color_key] is None and unmapped_indices:
                self.czi_channel_map[color_key] = unmapped_indices.pop(0)
                
        return self.czi_channel_map

    def stack_rgb_image(self, img):
        """Builds a strict (Z, Y, X, 3) RGB array for PIL processing."""
        # Create empty volume with exactly 3 channels
        sorted_img = np.zeros((*img.shape[:-1], 3), dtype=img.dtype)
        
        r_idx = self.czi_channel_map.get('R')
        g_idx = self.czi_channel_map.get('G')
        b_idx = self.czi_channel_map.get('B')
        
        # Standard RGB Mapping
        # Slot 0 = Red, Slot 1 = Green, Slot 2 = Blue
        if r_idx is not None and r_idx < img.shape[-1]: 
            sorted_img[..., 0] = img[..., r_idx]
            
        if g_idx is not None and g_idx < img.shape[-1]: 
            sorted_img[..., 1] = img[..., g_idx]
            
        if b_idx is not None and b_idx < img.shape[-1]: 
            sorted_img[..., 2] = img[..., b_idx]
        
        return sorted_img
    
    def get_lif_pixel_size_um(self, file_path, series_idx=0):
        """Extracts the physical X-axis pixel size in micrometers from a specific LIF series."""
        try:
            from readlif.reader import LifFile
            lif = LifFile(file_path)
            lif_img = lif.get_image(series_idx) # Target the specific series!
            
            x_scale = lif_img.scale[0]
            if x_scale:
                # readlif sometimes returns pixels/um instead of um/pixel
                if x_scale > 10.0:
                    return round(1.0 / x_scale, 4)
                return round(x_scale, 4)
        except Exception as e:
            print(f"Warning: LIF scale extraction failed. {e}")
        return None
    
    def get_czi_pixel_size_um(self, file_path):
        """Extracts the physical X-axis pixel size in micrometers from a CZI file."""
        try:
            from pylibCZIrw import czi as pyczi
            with pyczi.open_czi(file_path) as czidoc:
                metadata_dict = czidoc.metadata
                
                def find_distances(data):
                    found = []
                    if isinstance(data, dict):
                        for key, value in data.items():
                            if key == 'Distance':
                                if isinstance(value, list):
                                    found.extend(value)
                                else:
                                    found.append(value)
                            else:
                                found.extend(find_distances(value))
                    elif isinstance(data, list):
                        for item in data:
                            found.extend(find_distances(item))
                    return found
                
                raw_distances = find_distances(metadata_dict)
                for dist in raw_distances:
                    axis_id = dist.get("@Id") or dist.get("Id")
                    val_str = dist.get("Value")
                    
                    # We just need the X axis to determine the 2D scale
                    if axis_id and str(axis_id).upper() == 'X' and val_str:
                        val_meters = float(val_str)
                        return round(val_meters * 1e6, 4) # Convert to um
                        
        except Exception as e:
            print(f"Warning: Scale extraction failed. {e}")
        return None

    def load_image_from_index(self):
        """Loads the current file, normalizes it to a standard ZYXC matrix, and initializes UI."""
        if not hasattr(self, 'loaded_files') or not self.loaded_files: return

        # --- EXTRACT THE SPECIFIC ITEM DATA ---
        current_item = self.loaded_files[self.current_file_index]
        file_path = current_item['path']
        series_idx = current_item['series']
        display_name = current_item['name']
        
        # --- Update UI ---
        total = len(self.loaded_files)
        current = self.current_file_index + 1
        self.lbl_img_count.config(text=f"{current} / {total}")
        self.btn_prev_img.config(state=tk.NORMAL if self.current_file_index > 0 else tk.DISABLED)
        self.btn_next_img.config(state=tk.NORMAL if self.current_file_index < total - 1 else tk.DISABLED)
        self.lbl_filename.config(text="Loading...")
        self.canvas.delete("all")
        self.update() 

        try:
            self.original_filename = display_name
            self.czi_channel_map = {'R': 0, 'G': 1, 'B': 2} # Default fallback map
            
            # =========================================================
            # 1. FORMAT-SPECIFIC EXTRACTION (Translating to ZYXC Numpy)
            # =========================================================
            if file_path.lower().endswith('.czi'):
                import czifile
                import numpy as np

            elif file_path.lower().endswith('.lif'):
                from readlif.reader import LifFile
                import numpy as np
                
                lif = LifFile(file_path)
                lif_img = lif.get_image(series_idx) 
                
                z_slices = lif_img.info.get('z', 1)
                c_channels = lif_img.info.get('channels', 1)
                dims = lif_img.info.get('dims', (1, 1, 1, 1, 1))
                x_dim, y_dim = dims[0], dims[1]
                
                img = np.zeros((z_slices, y_dim, x_dim, c_channels), dtype=np.uint16)
                
                for z in range(z_slices):
                    for c in range(c_channels):
                        frame_pil = lif_img.get_frame(z=z, t=0, c=c)
                        img[z, :, :, c] = np.array(frame_pil)
                        
                self.original_num_channels = c_channels
                
                # --- EXTRACT EXACT CHANNEL DATA FOR THE DYNAMIC UI ---
                self._temp_extracted_channels = []
                try:
                    root = lif.xml_root
                    img_name = lif_img.name
                    
                    for element in root.iter("Element"):
                        if element.get("Name") == img_name:
                            channel_nodes = element.findall(".//ChannelDescription")
                            for c_idx, ch_node in enumerate(channel_nodes):
                                if c_idx >= c_channels: break
                                
                                ch_info = {}
                                c_name = ch_node.get("Name") or ch_node.get("LUTName")
                                if c_name: ch_info["name"] = c_name
                                
                                raw_color_val = ch_node.get("Color")
                                if raw_color_val:
                                    try:
                                        int_color = int(raw_color_val)
                                        if int_color < 0: int_color = (1 << 32) + int_color
                                        # Decode Leica's 32-bit signed integer color to RGB
                                        r = (int_color >> 16) & 0xFF
                                        g = (int_color >> 8) & 0xFF
                                        b = int_color & 0xFF
                                        ch_info["rgb"] = (r, g, b)
                                        ch_info["hex"] = f"#{r:02X}{g:02X}{b:02X}"
                                    except ValueError:
                                        pass
                                self._temp_extracted_channels.append(ch_info)
                except Exception as meta_err:
                    print(f"Warning: LIF XML channel mapping failed: {meta_err}")

            else:
                import tifffile
                import numpy as np
                
                # Load the raw TIFF array, ignoring multi-file OME linking if siblings are missing
                img_raw = tifffile.imread(file_path, is_ome=False)
                
                # Ensure float32 or uint16 type checking as per your pipeline
                if img_raw.dtype == np.uint8:
                    img_raw = img_raw.astype(np.uint16) * 257  # scale to 16-bit range roughly if needed, or keep raw
                
                # --- ROBUST 4D RESHAPING FOR TIFFs ---
                ndim = img_raw.ndim
                shape = img_raw.shape
                
                if ndim == 2:
                    # Case 1: Single 2D plane grayscale -> shape (Y, X)
                    # Convert to (1, Y, X, 1) -> 1 Z-slice, 1 Channel
                    img = img_raw[np.newaxis, :, :, np.newaxis]
                    
                elif ndim == 3:
                    # Case 2: Could be standard RGB (Y, X, 3) or a Grayscale Z-stack (Z, Y, X)
                    # Look at the last dimension to guess if it's channels
                    if shape[2] in [3, 4]:  # Standard RGB or RGBA plane
                        # Convert (Y, X, C) -> (1, Y, X, C)
                        img = img_raw[np.newaxis, :, :, :]
                    else:
                        # Convert Grayscale Z-stack (Z, Y, X) -> (Z, Y, X, 1)
                        img = img_raw[:, :, :, np.newaxis]
                        
                elif ndim == 4:
                    # Case 3: Already multi-channel Z-stack or time series -> shape (Z, Y, X, C) or (C, Z, Y, X)
                    # Standard microscopy outputs usually order multi-page tiffs as (Z, Y, X, C) or (Z, C, Y, X)
                    # If the last dimension is very large (like height/width), channels might be at axis 1 or 0.
                    if shape[3] > 20 and shape[1] <= 10:  # Appears to be (Z, C, Y, X)
                        img = np.transpose(img_raw, (0, 2, 3, 1))
                    elif shape[0] <= 10 and shape[3] > 20: # Appears to be (C, Z, Y, X)
                        img = np.transpose(img_raw, (1, 2, 3, 0))
                    else:
                        img = img_raw
                        
                else:
                    # Fallback for weirdly ordered high-dimensional hyperstacks
                    img = img_raw
                    while img.ndim < 4:
                        img = img[:, :, :, np.newaxis]
                    if img.ndim > 4:
                        img = img[0, :, :, :, 0] # Squash extra dimensions safely
                
                self.original_num_channels = img.shape[3]

            # =========================================================
            # 2. UNIVERSAL PIPELINE (Format-Agnostic from here down)
            # =========================================================
            
            c_total = img.shape[3]
            dynamic_channel_list = []
            
            # Default UI Palette (Reordered to standard microscopy: Blue, Green, Red)
            palette = [
                ((0, 0, 255), "#0000FF"),   # 0: Blue 
                ((0, 255, 0), "#00FF00"),   # 1: Green
                ((255, 0, 0), "#FF0000"),   # 2: Red
                ((0, 255, 255), "#00FFFF"), # 3: Cyan
                ((255, 0, 255), "#FF00FF"), # 4: Magenta
                ((255, 255, 0), "#FFFF00"), # 5: Yellow
                ((255, 255, 255), "#FFFFFF")# 6: White
            ]

            # --- SMART SINGLE-CHANNEL OVERRIDE ---
            # If the image is just 1 channel, default to Green instead of Blue
            if c_total == 1:
                palette[0] = ((0, 255, 0), "#00FF00") 
                # Note: If you ever prefer raw data to look Grayscale, 
                # just change the line above to: ((255, 255, 255), "#FFFFFF")

            # Generate the exact list of UI channels to match the matrix dimension
            for c in range(c_total):
                name = f"Channel {c+1}"
                rgb, hex_str = palette[c % len(palette)]
                
                # Apply specific metadata if we extracted it during loading
                if hasattr(self, '_temp_extracted_channels') and c < len(self._temp_extracted_channels):
                    ch_meta = self._temp_extracted_channels[c]
                    name = ch_meta.get("name", name)
                    
                    # If hardware provided an exact color, use it
                    if "rgb" in ch_meta:
                        rgb = ch_meta["rgb"]
                        hex_str = ch_meta["hex"]
                    else:
                        # SMART FALLBACK: If we have a name but no hardware color, guess it!
                        lower_name = name.lower()
                        if any(k in lower_name for k in ["blue", "dapi", "hoechst"]):
                            rgb, hex_str = (0, 0, 255), "#0000FF"
                        elif any(k in lower_name for k in ["green", "fitc", "gfp", "alexa 488"]):
                            rgb, hex_str = (0, 255, 0), "#00FF00"
                        elif any(k in lower_name for k in ["red", "tritc", "cy3", "cy5", "alexa 5"]):
                            rgb, hex_str = (255, 0, 0), "#FF0000"
                    
                dynamic_channel_list.append({
                    "name": name,
                    "rgb": rgb,
                    "hex": hex_str
                })

            # ---> TRIGGER THE UI TO REBUILD ITSELF FOR N CHANNELS <---
            self.build_dynamic_channels(dynamic_channel_list)
            
            # Clean up temp memory so it doesn't bleed into the next image!
            if hasattr(self, '_temp_extracted_channels'):
                del self._temp_extracted_channels

            # Initialize rendering matrices
            self.original_raw_volume = img.astype(np.float32)
            self.raw_volume = self.original_raw_volume
            
            self.max_z = int(img.shape[0] - 1)
            mid_z = int(self.max_z // 2)
            
            # Fast Cache: Spatial downsampling to calculate percentiles instantly
            self.z_percentiles = {}
            for z in range(self.max_z + 1):
                self.z_percentiles[z] = []
                for c in range(c_total):  
                    ch_data_sampled = self.raw_volume[z, ::4, ::4, c]
                    p_min, p_max = np.percentile(ch_data_sampled, (1.0, 99.9))
                    val_range = float(p_max - p_min) if (p_max - p_min) > 0 else 1.0
                    self.z_percentiles[z].append((p_min, val_range))

            # Reset view interactions
            self.zoom_scale = 1.0
            self.img_scale = 1.0
            self.img_x, self.img_y = 0, 0
            self.img_offset_x, self.img_offset_y = 0, 0
            self._initialized_view = False

            # UI Sliders
            self.scale_z.config(to=self.max_z)
            self.scale_z.set(mid_z)
            
            # Scale Bar Calibration
            pixel_size = None
            if file_path.lower().endswith('.czi'):
                pixel_size = self.get_czi_pixel_size_um(file_path)
            elif file_path.lower().endswith('.lif'):
                pixel_size = self.get_lif_pixel_size_um(file_path, series_idx) # PASS SPECIFIC SERIES INDEX!
                
            if pixel_size and hasattr(self, 'entry_pixel_size'):
                self.entry_pixel_size.delete(0, 'end')
                self.entry_pixel_size.insert(0, str(pixel_size))
            
            self.lbl_filename.config(text=self.original_filename)
            self.update_preview()
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.lbl_filename.config(text="Load failed")
            self.canvas.delete("all")

    # --- Processing ---
    def apply_image_math(self, image_multi, current_z=None):
        """Processes contrast and brightness dynamically for N channels."""
        import numpy as np
        
        # --- THE FIX: Restore missing channel dimension squashed by cv2.resize ---
        if image_multi.ndim == 2:
            image_multi = image_multi[:, :, np.newaxis]
            
        h, w, c_total = image_multi.shape
        
        # If no Z index is provided, default to the slider's current position
        if current_z is None and hasattr(self, 'scale_z'):
            current_z = int(float(self.scale_z.get()))
        elif current_z is None:
            current_z = 0

        var_adj = getattr(self, 'var_apply_adjustments', None)
        apply_adj = var_adj.get() if var_adj is not None else True
        
        # Get the names of the channels currently built in the UI
        channel_names = list(self.channel_vars.keys())
        processed_channels = []

        # ... (keep the rest of your apply_image_math loop exactly the same) ...
        for i, ch_name in enumerate(channel_names):
            if i >= c_total: 
                processed_channels.append(np.zeros((h, w), dtype=np.float32))
                continue
            
            is_visible = self.channel_vars[ch_name].get()
            contrast = self.adj_data[ch_name]["c"]
            brightness = self.adj_data[ch_name]["b"]

            ch_data = image_multi[:, :, i].astype(np.float32)

            # --- Raw display bypass ---
            if not apply_adj:
                if not is_visible:
                    processed_channels.append(np.zeros((h, w), dtype=np.float32))
                else:
                    norm_ch = ch_data / 65535.0
                    norm_ch = np.clip(norm_ch, 0.0, 1.0)
                    processed_channels.append(norm_ch)
                continue
            
            # --- Invisible bypass ---
            if not is_visible:
                processed_channels.append(np.zeros((h, w), dtype=np.float32))
                continue
            
            # Pull from cache if available
            if hasattr(self, 'z_percentiles') and current_z in self.z_percentiles and i < len(self.z_percentiles[current_z]):
                p_min, val_range = self.z_percentiles[current_z][i]
            else:
                p_min, p_max = np.percentile(ch_data, (1.0, 99.9))
                val_range = (p_max - p_min) if (p_max - p_min) > 0 else 1.0
            
            norm_ch = (ch_data - p_min) / val_range
            norm_ch = np.clip(norm_ch, 0.0, 1.0)
            
            norm_ch = (norm_ch * contrast) + brightness
            norm_ch = np.clip(norm_ch, 0.0, 1.0)
            
            processed_channels.append(norm_ch)

        # Pass the entire list of processed channels natively!
        return self.apply_pseudo_colors(processed_channels)

    def apply_pseudo_colors(self, processed_channels):
        """Blends an arbitrary number of normalized arrays dynamically using UI colors."""
        import numpy as np
        
        # 1. Stack the dynamic list of N channels into a single (H, W, N) matrix
        stacked_channels = np.stack(processed_channels, axis=-1)
        
        # 2. Build transformation matrix directly from our dynamic color dictionary
        channel_names = list(self.channel_vars.keys())
        colors = [self.channel_colors[name] for name in channel_names[:len(processed_channels)]]
        
        color_matrix = np.array(colors, dtype=np.float32)
        
        if color_matrix.max() <= 1.0 and color_matrix.max() > 0:
            color_matrix = color_matrix * 255.0

        # 3. Fast matrix dot product: (H, W, N) matrix * (N, 3) color weights -> (H, W, 3) RGB Image
        blended = np.dot(stacked_channels, color_matrix)
            
        # 4. Safe clip and cast to 8-bit image array
        return np.clip(blended, 0, 255).astype(np.uint8)

    def change_z_slice(self, direction):
        """Moves the Z-stack index up or down by 1 slice."""
        if self.raw_volume is None or self.is_merged_preview:
            return  # Skip navigation if no file is loaded or if previewing a merge
            
        current_val = self.scale_z.get()
        # Move up (+1) or down (-1)
        new_val = current_val + direction
        
        # Check boundary limits based on self.max_z
        max_z_limit = getattr(self, 'max_z', len(self.raw_volume) - 1)
        
        if 0 <= new_val <= max_z_limit:
            self.scale_z.set(new_val)      # Visually moves the slider layout handle
            self.update_preview()          # Reloads data, changes text, updates canvas

    def on_z_slider_move(self, val=None):
        """Interrupts the Z-slider to break out of merge mode if it's active."""
        if getattr(self, 'is_merged_preview', False):
            # Turn off merge mode
            self.is_merged_preview = False
            # Reset the button text 
            self.btn_preview_merge.config(text="👁 Preview\nMerge")
            
        # Continue updating the canvas with the new Z-slice
        self.update_preview()

    def update_preview(self, event=None):
        """Schedules a preview update, debouncing rapid consecutive calls to eliminate lag."""
        if self.raw_volume is None: 
            return
            
        # Cancel any pending preview updates that haven't executed yet
        if hasattr(self, '_preview_after_id') and self._preview_after_id:
            self.canvas.after_cancel(self._preview_after_id)
            
        # Schedule execution 15ms out to eliminate slider drag lagging
        self._preview_after_id = self.canvas.after(15, self._execute_update_preview)

    def _execute_update_preview(self):
        """The actual heavy preview updater, optimized for strict cache hits and PIL image generation."""
        self._preview_after_id = None
        
        # 1. Pull the raw matrix data frame out of RAM cache
        if self.is_merged_preview:
            try:
                z_start = max(0, min(int(self.spin_z_start.get()), self.max_z))
                z_end = max(0, min(int(self.spin_z_end.get()), self.max_z))
                if z_start > z_end: 
                    z_start, z_end = z_end, z_start
            except ValueError:
                z_start, z_end = 0, self.max_z
                
            self.lbl_z_current.config(text=f"Previewing Merge: Stacks {z_start} to {z_end}")
            stack_slice = self.raw_volume[z_start:z_end+1]
            
            # Map dropdown string selections safely
            combo_method = getattr(self, 'combo_proj_method', None)
            method = combo_method.get() if combo_method is not None else "Max IP"
            
            if method in ["Maximum", "Max IP"]:
                self.active_raw_slice = np.max(stack_slice, axis=0)
            elif method in ["Mean", "Mean IP"]:
                self.active_raw_slice = np.mean(stack_slice, axis=0).astype(stack_slice.dtype)
            elif method in ["Sum", "Sum IP"]:
                self.active_raw_slice = np.clip(np.sum(stack_slice, axis=0, dtype=np.float32), 0, 65535).astype(stack_slice.dtype)
            elif method in ["Minimum", "Min IP"]:
                self.active_raw_slice = np.min(stack_slice, axis=0)
                
            self.current_z_idx = int((z_start + z_end) // 2)  
        else:
            self.current_z_idx = int(float(self.scale_z.get()))
            self.lbl_z_current.config(text=f"Current Stack: {self.current_z_idx}")
            self.active_raw_slice = self.raw_volume[self.current_z_idx]

        img_h, img_w = self.active_raw_slice.shape[:2]
        if img_w == 0 or img_h == 0: 
            return 

        # 2. Setup initial scale layout matrices on first pass
        if not hasattr(self, '_initialized_view') or not self._initialized_view:
            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            if canvas_w < 10: 
                canvas_w, canvas_h = 800, 600
            
            fit_scale = min(canvas_w / img_w, canvas_h / img_h)
            self.img_scale = fit_scale
            self.zoom_scale = fit_scale
            self.scale_x = fit_scale
            
            self.img_offset_x = (canvas_w - int(img_w * fit_scale)) // 2  
            self.img_offset_y = (canvas_h - int(img_h * fit_scale)) // 2  
            self.img_x = self.img_offset_x
            self.img_y = self.img_offset_y
            self._initialized_view = True

        # 3. FIX: Process matrix parameters and generate self.current_pil_image for panning
        view_image_uint8 = self.apply_image_math(self.active_raw_slice, self.current_z_idx)
        self.current_pil_image = Image.fromarray(view_image_uint8)

        # 4. Fire renderer loop
        self.redraw_image()

    def redraw_image(self):
        """Intelligently processes only the visible window viewport at high resolution."""
        if not hasattr(self, 'active_raw_slice') or self.active_raw_slice is None:
            return

        import numpy as np
        from PIL import Image, ImageTk
        import tkinter as tk
        import cv2

        self.canvas.delete("all")

        # Maintain global register synchronization definitions
        self.img_scale = self.zoom_scale
        self.scale_x = self.zoom_scale
        self.img_offset_x = self.img_x
        self.img_offset_y = self.img_y

        orig_h, orig_w = self.active_raw_slice.shape[:2]
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10: 
            canvas_w, canvas_h = 800, 600

        # 1. Calculate bounding coordinates of what part of the image is inside the canvas
        src_x1 = max(0, int(-self.img_x / self.zoom_scale))
        src_y1 = max(0, int(-self.img_y / self.zoom_scale))
        src_x2 = min(orig_w, int((canvas_w - self.img_x) / self.zoom_scale) + 1)
        src_y2 = min(orig_h, int((canvas_h - self.img_y) / self.zoom_scale) + 1)

        # Calculate where on the Tkinter canvas the cropped piece should actually start drawing
        dest_x = max(self.img_x, self.img_x + int(src_x1 * self.zoom_scale))
        dest_y = max(self.img_y, self.img_y + int(src_y1 * self.zoom_scale))

        # Calculate width and height of the sub-view on screen
        view_w = max(1, int((src_x2 - src_x1) * self.zoom_scale))
        view_h = max(1, int((src_y2 - src_y1) * self.zoom_scale))

        # --- HYBRID RENDERING STRATEGY ---
        if self.zoom_scale > 1.0:
            # STRATEGY A (Zoomed In): Extract full-resolution cropped region first, then process math
            cropped_raw = self.active_raw_slice[src_y1:src_y2, src_x1:src_x2]
            processed_sub_region = self.apply_image_math(cropped_raw, self.current_z_idx)
            
            # Resize processed chunk to match canvas footprint size
            if processed_sub_region.dtype != np.uint8:
                processed_sub_region = (processed_sub_region * 255.0).astype(np.uint8)
                
            view_image = cv2.resize(processed_sub_region, (view_w, view_h), interpolation=cv2.INTER_NEAREST)
        else:
            # STRATEGY B (Zoomed Out / Overview): Downsample first to remain fluid, then process math
            cropped_raw = self.active_raw_slice[src_y1:src_y2, src_x1:src_x2]
            downsampled_raw = cv2.resize(cropped_raw, (view_w, view_h), interpolation=cv2.INTER_NEAREST)
            
            view_image = self.apply_image_math(downsampled_raw, self.current_z_idx)
            if view_image.dtype != np.uint8:
                view_image = (view_image * 255.0).astype(np.uint8)

        # 2. Render processed matrix data slice to screen
        view_pil_image = Image.fromarray(view_image)
        self.tk_img = ImageTk.PhotoImage(view_pil_image)
        self.canvas.create_image(dest_x, dest_y, anchor=tk.NW, image=self.tk_img)
        self.rect_id = None 

        # 3. Re-render cropping selection bounding boxes if actively drawn
        if hasattr(self, 'current_rect') and self.current_rect:
            x1, y1, x2, y2 = self.current_rect
            cx1 = x1 * self.zoom_scale + self.img_x
            cy1 = y1 * self.zoom_scale + self.img_y
            cx2 = x2 * self.zoom_scale + self.img_x
            cy2 = y2 * self.zoom_scale + self.img_y
            self.rect_id = self.canvas.create_rectangle(cx1, cy1, cx2, cy2, outline="red", width=2)

        # 4. Render scale bar overlay layer dynamically
        if hasattr(self, 'draw_scale_bar'):
            self.draw_scale_bar()

    # --- Saving ---
    def save_image_to_disk(self):
        if self.raw_volume is None: return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".tif",
            filetypes=[("TIFF File", "*.tif *.tiff"), ("PNG File", "*.png"), ("JPEG File", "*.jpg *.jpeg")],
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
                export_z = int((z_start + z_end) // 2)
            else:
                export_z = self.scale_z.get()
                target_data = self.raw_volume[export_z]
            
            final_rgb = self.apply_image_math(target_data, current_z=export_z)
            final_rgb = self.stamp_scale_bar_for_export(final_rgb)
            
            # Ensure data layout is continuous in memory for optimal writing performance
            final_rgb = np.ascontiguousarray(final_rgb)
            
            if file_path.lower().endswith(('.tif', '.tiff')):
                try:
                    pixel_size_um = float(self.entry_pixel_size.get())
                except (ValueError, AttributeError):
                    pixel_size_um = 0 
                
                if pixel_size_um > 0:
                    # Convert to a clean rational tuple to prevent metadata truncation bugs
                    # 1 cm = 10,000 um. We store resolution as pixels per 10,000 units.
                    resolution_val = 10000.0 / pixel_size_um
                    
                    # Convert to rational fraction (numerator, denominator) for the TIFF standard
                    res_fraction = float(resolution_val).as_integer_ratio()
                    
                    tifffile.imwrite(
                        file_path, 
                        final_rgb,
                        resolution=res_fraction,
                        resolutionunit=3,  # 3 = CENTIMETER
                        metadata={'unit': 'um'},  # Standard for ImageJ/Fiji compatibility
                        compression='zlib'  # Enforce lossless compression to preserve sharpness
                    )
                else:
                    tifffile.imwrite(file_path, final_rgb, compression='zlib')
                    
            elif file_path.lower().endswith('.png'):
                final_bgr = cv2.cvtColor(final_rgb, cv2.COLOR_RGB2BGR)
                # Enforce maximum, lossless PNG compression to maintain high-res detail
                cv2.imwrite(file_path, final_bgr, [int(cv2.IMWRITE_PNG_COMPRESSION), 0])
                
            else:
                final_bgr = cv2.cvtColor(final_rgb, cv2.COLOR_RGB2BGR)
                # If saving as JPEG, maximize quality to 100% to minimize artifacts
                cv2.imwrite(file_path, final_bgr, [int(cv2.IMWRITE_JPEG_QUALITY), 100])
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Save Error", f"Failed to save image:\n{e}")
