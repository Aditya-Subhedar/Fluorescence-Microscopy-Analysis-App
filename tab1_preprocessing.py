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
        tk.Button(action_frame, text="1. Load CZI Images", command=self.load_czi, font=("Arial", 11, "bold"), bg="#4a90e2", fg="white").pack(fill=tk.X)
        
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
        
            # Button takes up 2 rows on the right side
        self.btn_preview_merge = tk.Button(proj_frame, text="👁 Preview\nMerge", command=self.toggle_merge_preview)
        self.btn_preview_merge.grid(row=0, column=2, rowspan=2, padx=(5, 0), sticky="nsew")
        proj_frame.grid_columnconfigure(2, weight=1) # Makes the button expand to fill the right side

        # Cropping Tool
        crop_frame = tk.LabelFrame(control_frame, text="Cropping Tool", padx=10, pady=5)
        crop_frame.pack(fill=tk.X, pady=2)
        tk.Label(crop_frame, text="Draw a rectangle on the image to crop.", fg="gray").pack(anchor="w", pady=(0,2))
        tk.Button(crop_frame, text="✂ Crop to Selection", command=self.apply_crop, bg="#f39c12", fg="white").pack(fill=tk.X, pady=2)
        tk.Button(crop_frame, text="🔄 Reset Original Image", command=self.reset_crop).pack(fill=tk.X, pady=2)

        # ---------------------------------------------------------
        # Channel Visibility
        # ---------------------------------------------------------
        chan_frame = tk.LabelFrame(control_frame, text="Channel Visibility & Pseudo-Color", padx=10, pady=5)
        chan_frame.pack(fill=tk.X, pady=2)
        
        self.var_ch_r = tk.BooleanVar(value=True)
        self.var_ch_g = tk.BooleanVar(value=True)
        self.var_ch_b = tk.BooleanVar(value=True)

        # Variables to store the actual RGB tuples for rendering/blending
        self.color_r = (255, 0, 0)   # Default Pure Red
        self.color_g = (0, 255, 0)   # Default Pure Green
        self.color_b = (0, 0, 255)   # Default Pure Blue

        # Build the three rows and store the UI elements as class attributes 
        # so self.pick_color() can access and update them later.
        self.btn_color_r, self.lbl_ch_r = self.create_channel_row(chan_frame, "Alexa Fluor 568 (Red)", self.var_ch_r, "R", "#FF0000")
        self.btn_color_g, self.lbl_ch_g = self.create_channel_row(chan_frame, "Alexa Fluor 488 (Green)", self.var_ch_g, "G", "#00FF00")
        self.btn_color_b, self.lbl_ch_b = self.create_channel_row(chan_frame, "DAPI (Blue)", self.var_ch_b, "B", "#0000FF")

        # ---------------------------------------------------------
        # Channel Adjustments (Compact)
        # ---------------------------------------------------------
        adj_frame = tk.LabelFrame(control_frame, text="Channel Adjustments", padx=10, pady=5)
        adj_frame.pack(fill=tk.X, pady=2)
        
        # Data dictionary to remember the user's settings for each channel
        self.adj_data = {
            "Red (Alexa 568)": {"c": 1.0, "b": 0.0},
            "Green (Alexa 488)": {"c": 1.0, "b": 0.0},
            "Blue (DAPI)": {"c": 1.0, "b": 0.0}
        }
        self.active_adj_channel = tk.StringVar(value="Red (Alexa 568)")
        self._is_updating_ui = False # Flag to prevent feedback loops
        
        # Dropdown to select the channel (Uses PACK)
        self.combo_channel = ttk.Combobox(adj_frame, textvariable=self.active_adj_channel, 
                                          values=list(self.adj_data.keys()), state="readonly")
        self.combo_channel.pack(fill=tk.X, pady=(0, 5))
        self.combo_channel.bind("<<ComboboxSelected>>", self.on_adj_channel_change)
        
        # Sub-frame for sliders (Uses PACK inside adj_frame)
        slider_frame = tk.Frame(adj_frame)
        slider_frame.pack(fill=tk.X)
        
        # Configure columns so both sliders get equal horizontal space
        slider_frame.columnconfigure(1, weight=1)
        slider_frame.columnconfigure(3, weight=1)
        
        # Contrast Slider (C:) - Side by side!
        tk.Label(slider_frame, text="C:", font=("Arial", 9, "bold"), fg="gray").grid(row=0, column=0, sticky="e")
        self.scale_contrast = tk.Scale(slider_frame, from_=0.1, to=5.0, resolution=0.1, orient=tk.HORIZONTAL, 
                                       command=self.on_shared_slider_move, 
                                       width=10, sliderlength=15) # Made thinner & less chunky
        self.scale_contrast.grid(row=0, column=1, sticky="ew", padx=(0, 5))
        
        # Brightness Slider (B:) - Side by side!
        tk.Label(slider_frame, text="B:", font=("Arial", 9, "bold"), fg="gray").grid(row=0, column=2, sticky="e")
        self.scale_brightness = tk.Scale(slider_frame, from_=-1.0, to=1.0, resolution=0.05, orient=tk.HORIZONTAL, 
                                         command=self.on_shared_slider_move, 
                                         width=10, sliderlength=15) # Made thinner & less chunky
        self.scale_brightness.grid(row=0, column=3, sticky="ew", padx=(0, 2))
        
        # Initialize sliders
        self.on_adj_channel_change()

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
        
        # Mouse Bindings for Cropping
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_release)

        # --- NEW: Keyboard bindings for Left/Right arrows ---
        top = self.winfo_toplevel()
        top.bind("<Left>", lambda e: self.prev_image() if self.btn_prev_img['state'] == tk.NORMAL else None)
        top.bind("<Right>", lambda e: self.next_image() if self.btn_next_img['state'] == tk.NORMAL else None)

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
        if not hasattr(self, 'loaded_files') or not self.loaded_files: return

        file_path = self.loaded_files[self.current_file_index]
        
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
            self.original_filename = os.path.basename(file_path)
            # Default fallback map in case metadata is missing entirely
            self.czi_channel_map = {'R': 0, 'G': 1, 'B': 2}
            
            # --- 1. LOAD RAW DATA & NORMALIZE DIMENSIONS ---
            if file_path.lower().endswith('.czi'):
                import czifile
                import numpy as np
                
                with czifile.CziFile(file_path) as czi:
                    img = czi.asarray()
                    raw_axes = list(czi.axes) # e.g., ['B', 'C', 'Z', 'Y', 'X', '0']
                    
                    # Track original number of channels before manipulating shape
                    c_index = raw_axes.index('C') if 'C' in raw_axes else -1
                    self.original_num_channels = img.shape[c_index] if c_index != -1 else 1

                    # Remove dimensions of size 1
                    squeeze_indices = [i for i, dim in enumerate(img.shape) if dim == 1]
                    img = np.squeeze(img)
                    
                    # Update axes list to match the squeezed image
                    current_axes = [ax for i, ax in enumerate(raw_axes) if i not in squeeze_indices]
                    
                    # Force array into (Z, Y, X, C) order
                    target_axes = ['Z', 'Y', 'X', 'C']
                    
                    # If an axis is missing (e.g., no Z stack), we add a fake axis
                    for ax in target_axes:
                        if ax not in current_axes:
                            img = img[..., np.newaxis]
                            current_axes.append(ax)
                            
                    # Now dynamically move axes to match (Z, Y, X, C)
                    source_indices = [current_axes.index(ax) for ax in target_axes]
                    target_indices = [0, 1, 2, 3]
                    img = np.moveaxis(img, source_indices, target_indices)

                    # --- 2. EXTRACT METADATA (Using pylibCZIrw Dict Logic) ---
                    try:
                        from pylibCZIrw import czi as pyczi
                        with pyczi.open_czi(file_path) as czidoc:
                            metadata_dict = czidoc.metadata
                            
                            def find_channels(data):
                                found = []
                                if isinstance(data, dict):
                                    for key, value in data.items():
                                        if key == 'Channel':
                                            if isinstance(value, list):
                                                found.extend(value)
                                            else:
                                                found.append(value)
                                        else:
                                            found.extend(find_channels(value))
                                elif isinstance(data, list):
                                    for item in data:
                                        found.extend(find_channels(item))
                                return found
                            
                            raw_channels = find_channels(metadata_dict)
                            unique_channels = []
                            seen_names = set()
                            
                            for ch in raw_channels:
                                c_name = ch.get("@Name") or ch.get("Name") or "Unknown"
                                c_color = ch.get("Color") or "Unknown"
                                wavelength = ch.get("EmissionWavelength") or "N/A"
                                
                                if c_name not in seen_names:
                                    unique_channels.append({
                                        "Name": c_name,
                                        "Color": c_color,
                                        "Wavelength": wavelength
                                    })
                                    seen_names.add(c_name)
                                else:
                                    if wavelength != "N/A":
                                        for existing_ch in unique_channels:
                                            if existing_ch["Name"] == c_name and existing_ch["Wavelength"] == "N/A":
                                                existing_ch["Wavelength"] = wavelength

                            # Pass our clean list of dictionaries to the new mapping function!
                            self.map_channels_from_xml(unique_channels)
                            
                    except Exception as meta_err:
                        print(f"Warning: Metadata extraction failed. Using fallback map. Error: {meta_err}")

            else:
                import tifffile
                import numpy as np
                img = tifffile.imread(file_path)
                # Tiff fallback (Assume it's already ZYXC or guess shape)
                img = np.squeeze(img)
                if img.ndim == 3: img = img[np.newaxis, ...]
                self.original_num_channels = img.shape[-1]
                self.czi_channel_map = {'R': 0, 'G': 1, 'B': 2} # Default TIFF fallback

            # --- 3. CONVERT TO RGB ---
            # CALL FUNCTION 2: Stack the final image using the map
            img = self.stack_rgb_image(img)

            # --- 4. INITIALIZE RENDERING VARIABLES ---
            self.original_raw_volume = img.astype(np.float32)
            self.raw_volume = self.original_raw_volume
            
            self.max_z = int(img.shape[0] - 1)
            mid_z = int(self.max_z // 2)
            
            # Auto-calculate display ranges for contrast
            self.channel_baselines = []
            z_step = max(1, img.shape[0] // 5) 
            for c in range(3): 
                sub_vol = self.raw_volume[::z_step, ::4, ::4, c]
                valid_pixels = sub_vol[sub_vol > 0]
                if len(valid_pixels) > 0:
                    p_min, p_max = np.percentile(valid_pixels, (0.1, 99.9))
                else:
                    p_min, p_max = 0, 1
                if p_max <= p_min: p_max = p_min + 1
                self.channel_baselines.append({'min': float(p_min), 'max': float(p_max)})

            # Update UI Sliders
            self.scale_z.config(to=self.max_z)
            self.scale_z.set(mid_z)
            
            # --- 5. AUTO-CALIBRATE SCALE BAR ---
            if file_path.lower().endswith('.czi'):
                pixel_size = self.get_czi_pixel_size_um(file_path)
                if pixel_size and hasattr(self, 'entry_pixel_size'):
                    # Clear the old manual entry and insert the exact hardware calibration
                    self.entry_pixel_size.delete(0, 'end')
                    self.entry_pixel_size.insert(0, str(pixel_size))
            
            self.lbl_filename.config(text=self.original_filename)
            self.canvas.after(100, self.update_preview)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.lbl_filename.config(text="Load failed")
            self.canvas.delete("all")

    # --- Processing ---
    def apply_image_math(self, image_multi):
        """Processes contrast and brightness for each channel, then sends to blender."""
        import numpy as np
        h, w, c_total = image_multi.shape
        
        # --- READ FROM COMPACT UI DICTIONARY ---
        channel_params = [
            (self.adj_data["Red (Alexa 568)"]["c"], self.adj_data["Red (Alexa 568)"]["b"], self.var_ch_r.get()),
            (self.adj_data["Green (Alexa 488)"]["c"], self.adj_data["Green (Alexa 488)"]["b"], self.var_ch_g.get()),
            (self.adj_data["Blue (DAPI)"]["c"], self.adj_data["Blue (DAPI)"]["b"], self.var_ch_b.get())
        ]
        
        processed_channels = []

        for i, (contrast, brightness, is_visible) in enumerate(channel_params):
            # If channel is hidden or doesn't exist, pass a blank black array
            if not is_visible or i >= c_total or i >= len(self.channel_baselines): 
                processed_channels.append(np.zeros((h, w), dtype=np.float32))
                continue
            
            ch_data = image_multi[:, :, i].astype(np.float32)
            
            b_min = self.channel_baselines[i]['min']
            b_max = self.channel_baselines[i]['max']
            val_range = (b_max - b_min) if (b_max - b_min) > 0 else 1.0
            
            # 1. Normalize channel based on 16-bit limits
            norm_ch = (ch_data - b_min) / val_range
            norm_ch = np.clip(norm_ch, 0.0, 1.0)
            
            # 2. Apply individual Contrast and Brightness
            norm_ch = (norm_ch * contrast) + brightness
            norm_ch = np.clip(norm_ch, 0.0, 1.0)
            
            processed_channels.append(norm_ch)

        # 3. Send the fully processed grayscale arrays to the pseudo-color blender
        return self.apply_pseudo_colors(processed_channels[0], processed_channels[1], processed_channels[2])

    def apply_pseudo_colors(self, norm_r, norm_g, norm_b):
        """Blends 3 normalized grayscale arrays (0.0 to 1.0) using the chosen UI colors."""
        import numpy as np
        
        h, w = norm_r.shape
        blended = np.zeros((h, w, 3), dtype=np.float32)
        
        # Link the grayscale arrays to their dynamically chosen UI colors
        layers = [
            (norm_r, self.color_r),
            (norm_g, self.color_g),
            (norm_b, self.color_b)
        ]
        
        # Additive Blending (Matches ImageJ "Composite" rendering)
        for norm_ch, color in layers:
            # color[0]=Red, color[1]=Green, color[2]=Blue
            blended[:, :, 0] += norm_ch * color[0] 
            blended[:, :, 1] += norm_ch * color[1] 
            blended[:, :, 2] += norm_ch * color[2] 
            
        return np.clip(blended, 0, 255).astype(np.uint8)
    
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
        self.scale_x = scale  # <-- NEW: Feeds the exact preview zoom into draw_scale_bar!
        
        self.img_offset_x = (canvas_w - new_w) // 2
        self.img_offset_y = (canvas_h - new_h) // 2
        
        preview_small = cv2.resize(raw_data, (new_w, new_h), interpolation=cv2.INTER_NEAREST)
        preview_edited = self.apply_image_math(preview_small)
        
        self.tk_img = ImageTk.PhotoImage(Image.fromarray(preview_edited))
        self.canvas.delete("all")
        self.canvas.create_image(canvas_w//2, canvas_h//2, anchor=tk.CENTER, image=self.tk_img)
        self.rect_id = None # Clear old bounding box state on redraw

        # Draw scale bar on top
        self.draw_scale_bar()

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
            
            # 1. Calculate the RGB image
            final_rgb = self.apply_image_math(target_data)
            
            # 2. Burn the scale bar into the image pixels
            final_rgb = self.stamp_scale_bar_for_export(final_rgb)
            
            # 3. Save to disk (TIFF retains RGB, others convert to BGR for OpenCV)
            if file_path.lower().endswith(('.tif', '.tiff')):
                # Fetch the hidden pixel size
                try:
                    pixel_size_um = float(self.entry_pixel_size.get())
                except (ValueError, AttributeError):
                    pixel_size_um = 0 # Fallback if empty or invalid
                
                # If we have a valid pixel size, save with spatial metadata
                if pixel_size_um > 0:
                    # TIFF standard uses pixels per centimeter. 1 cm = 10,000 um.
                    pixels_per_cm = 10000 / pixel_size_um
                    
                    tifffile.imwrite(
                        file_path, 
                        final_rgb,
                        resolution=(pixels_per_cm, pixels_per_cm),
                        resolutionunit=3, # 3 = CENTIMETER in TIFF standard
                        metadata={'unit': 'um'} # Ensures ImageJ/Fiji reads the unit correctly
                    )
                else:
                    tifffile.imwrite(file_path, final_rgb) # Save without metadata
            else:
                final_bgr = cv2.cvtColor(final_rgb, cv2.COLOR_RGB2BGR)
                cv2.imwrite(file_path, final_bgr) 
                
            # Removed the annoying Success messagebox so you can export in peace!
            
        except Exception as e:
            # We keep the error popup, because if it fails, you definitely want to know
            messagebox.showerror("Save Error", f"Failed to save image:\n{e}")