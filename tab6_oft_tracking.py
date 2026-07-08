import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cv2
import numpy as np
import math
import csv
from PIL import Image, ImageTk

class OFTTrackingTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # ================= UI BEAUTIFICATION =================
        self.style = ttk.Style()
        
        # Configure specific styles for this tab to avoid breaking other tabs
        self.style.configure("OFT.TFrame", background="#f8fafc")
        self.style.configure("OFT_Card.TFrame", background="#ffffff")
        self.style.configure("OFT.TLabelframe", background="#ffffff", bordercolor="#e2e8f0", borderwidth=1)
        self.style.configure("OFT.TLabelframe.Label", font=("Segoe UI", 11, "bold"), background="#ffffff", foreground="#0f172a")
        self.style.configure("OFT_Header.TLabel", font=("Segoe UI", 11, "bold"), background="#ffffff", foreground="#334155")
        self.style.configure("OFT_Instruct.TLabel", font=("Segoe UI", 9, "italic"), background="#ffffff", foreground="#64748b")
        self.style.configure("OFT_MetricTitle.TLabel", font=("Segoe UI", 10), background="#ffffff", foreground="#475569")
        self.style.configure("OFT_MetricValue.TLabel", font=("Segoe UI", 14, "bold"), background="#ffffff")
        self.style.configure("OFT.TButton", font=("Segoe UI", 9), padding=4)
        
        # Apply base background to the main tab
        self.config(style="OFT.TFrame")
        
        # Video & Playback State
        self.video_path = None
        self.cap = None
        self.current_frame = None
        self.video_playing = False
        
        # Dimensions & Viewports (Restored for full width layout)
        self.orig_w = 0
        self.orig_h = 0
        self.canvas_w = 640       
        self.canvas_h = 360       # 16:9 Aspect Ratio
        self.map_size = 400       
        self.map_display_size = 360
        
        # Calibration State 
        self.calibrating = False
        self.active_drag_pt = None
        self.calib_pts = {
            'TL': [220, 120], 'TR': [420, 120],
            'BR': [460, 300], 'BL': [180, 300]
        }
        self.homography_matrix = None
        self.scale_factor = 1.0
        
        # Bounding Box Selection State
        self.placement_mode = False
        self.box_start = None
        self.box_end = None
        self.track_window = None 
        
        # Kinematics & Tracking 
        self.smoothed_cx = None
        self.smoothed_cy = None
        self.base_inertia = 0.1
        self.velocity_deadzone_cm = 2.0  # Lowered default threshold for better responsiveness
        self.playback_speed = 1.0
        self.bg_subtractor = None
        self.tracking_active = False
        
        # Quantitative Metrics
        self.total_distance_cm = 0.0
        self.center_frames = 0
        self.fps = 30.0
        self.path_coordinates = []
        self.odometer_anchor_cm = None
        self.trajectory_canvas = None
        
        self.create_widgets()
        self.setup_keybindings()

    def setup_keybindings(self):
        # 1. Global window bindings
        top_level = self.winfo_toplevel()
        top_level.bind("<space>", self.handle_spacebar)
        top_level.bind("<Left>", self.handle_left_arrow)
        top_level.bind("<Right>", self.handle_right_arrow)
        top_level.bind("<Control-o>", self.handle_ctrl_o)
        top_level.bind("<Control-O>", self.handle_ctrl_o)
        
        # 2. Intercept spacebar on all buttons so they don't execute natively when focused
        buttons_to_intercept = [
            self.btn_browse, self.btn_play_pause, self.btn_calibrate, 
            self.btn_confirm_calib, self.btn_init_placement, self.btn_track, 
            self.btn_export_data, self.btn_export_map, self.btn_reset_data
        ]
        for btn in buttons_to_intercept:
            btn.bind("<space>", self.handle_spacebar)

    def handle_spacebar(self, event):
        focused_widget = self.focus_get()
        
        # Safe zone: If the user is actively typing in the calibration entry, let them type spaces
        if isinstance(focused_widget, tk.Entry):
            return None 
            
        # If a video is loaded, toggle playback
        if self.cap is not None:
            self.toggle_video_playback()
            
        # CRITICAL: Returning "break" intercepts the event and stops Tkinter 
        # from passing the spacebar press down to click any focused button.
        return "break"

    def handle_ctrl_o(self, event):
        focused_widget = self.focus_get()
        if isinstance(focused_widget, tk.Entry):
            return
            
        self.load_video()
        return "break"

    def handle_left_arrow(self, event):
        focused_widget = self.focus_get()
        if isinstance(focused_widget, tk.Entry): return
        self.seek_video(-int(self.fps * 2)) # Jump back 2 seconds
        return "break"

    def handle_right_arrow(self, event):
        focused_widget = self.focus_get()
        if isinstance(focused_widget, tk.Entry): return
        self.seek_video(int(self.fps * 2)) # Jump forward 2 seconds
        return "break"

    def create_widgets(self):
        # Configure the grid to use 100% of the horizontal and vertical space
        self.rowconfigure(0, weight=1) 
        self.columnconfigure(0, weight=1)
        
        main_container = ttk.Frame(self, style="OFT.TFrame")
        # Grid it at 0,0 and let it stick to all edges (nsew)
        main_container.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        self.lbl_instruct = ttk.Label(main_container, text="Step 1: Load a video file to begin.", font=("Segoe UI", 12, "bold"), foreground="#2c3e50")
        self.lbl_instruct.pack(fill="x", pady=(0, 10))
        
        # ================= TOP PANEL: VIEWPORTS =================
        display_panel = ttk.Frame(main_container, style="OFT.TFrame")
        display_panel.pack(fill="both", expand=True, pady=(0, 20))
        
        self.frame_vid = ttk.LabelFrame(display_panel, text=" Monitor Viewport ", style="OFT.TLabelframe", padding=15)
        self.frame_vid.pack(side="left", padx=(0, 10), fill="both", expand=True)
        
        # Add a subtle background to the canvas area
        canvas_bg_frame = tk.Frame(self.frame_vid, bg="#000000", bd=0)
        canvas_bg_frame.pack(anchor="center", expand=True)
        
        self.canvas_video = tk.Canvas(canvas_bg_frame, width=self.canvas_w, height=self.canvas_h, bg="#1e1e1e", highlightthickness=0)
        self.canvas_video.pack(padx=2, pady=2)
        
        self.canvas_video.bind("<ButtonPress-1>", self.on_canvas_click)
        self.canvas_video.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas_video.bind("<ButtonRelease-1>", self.on_canvas_release)
        
        frame_path = ttk.LabelFrame(display_panel, text=" Reconstructed Trajectory ", style="OFT.TLabelframe", padding=15)
        frame_path.pack(side="right", padx=(10, 0), fill="both", expand=True)
        
        self.lbl_path = ttk.Label(frame_path, background="#ffffff", relief="flat")
        self.lbl_path.pack(anchor="center", expand=True)

        # ================= BOTTOM PANEL: CONTROLS =================
        # Reduced padding from 20 to 10 to save vertical space
        control_panel = ttk.LabelFrame(main_container, text=" OFT Inertial Tracker Controls ", style="OFT.TLabelframe", padding=10)
        control_panel.pack(fill="x")
        
        for i in range(4):
            control_panel.columnconfigure(i, weight=1)
            
        # -- Col 0: Video Source --
        col0 = ttk.Frame(control_panel, style="OFT_Card.TFrame")
        col0.grid(row=0, column=0, sticky="nsew", padx=5)
        ttk.Label(col0, text="1. Video Source", style="OFT_Header.TLabel").pack(anchor="w", pady=(0,4))
        
        # Grouped Load & Play buttons horizontally
        btn_frame_0 = ttk.Frame(col0, style="OFT_Card.TFrame")
        btn_frame_0.pack(fill="x", pady=(0, 5))
        self.btn_browse = ttk.Button(btn_frame_0, text="Load Video", style="OFT.TButton", command=self.load_video)
        self.btn_browse.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.btn_play_pause = ttk.Button(btn_frame_0, text="Play/Pause", style="OFT.TButton", command=self.toggle_video_playback, state="disabled")
        self.btn_play_pause.pack(side="left", fill="x", expand=True, padx=(2, 0))
        
        ttk.Label(col0, text="Playback Speed:", style="OFT_MetricTitle.TLabel").pack(anchor="w")
        self.scale_speed = ttk.Scale(col0, from_=0.5, to=5.0, orient="horizontal", command=self.update_speed)
        self.scale_speed.pack(fill="x")
        self.lbl_speed_val = ttk.Label(col0, text="1.0x", style="OFT_Instruct.TLabel")
        self.lbl_speed_val.pack(anchor="e")
        self.scale_speed.set(1.0)
        
        # -- Col 1: Calibration --
        col1 = ttk.Frame(control_panel, style="OFT_Card.TFrame")
        col1.grid(row=0, column=1, sticky="nsew", padx=5)
        ttk.Label(col1, text="2. Arena Calibration", style="OFT_Header.TLabel").pack(anchor="w", pady=(0,2))
        
        # Add target selection variable
        if not hasattr(self, 'calib_mode'):
            self.calib_mode = tk.StringVar(value="Outer")

        # Mode Selection Radiobuttons
        mode_frame = ttk.Frame(col1, style="OFT_Card.TFrame")
        mode_frame.pack(fill="x", pady=(0,2))
        ttk.Radiobutton(mode_frame, text="Drag Outer", variable=self.calib_mode, value="Outer", command=self.draw_draggable_poly).pack(side="left", expand=True)
        ttk.Radiobutton(mode_frame, text="Drag Inner", variable=self.calib_mode, value="Inner", command=self.draw_draggable_poly).pack(side="left", expand=True)
        
        # Outer Dimension (Bound to real-time updates)
        frame_outer = ttk.Frame(col1, style="OFT_Card.TFrame")
        frame_outer.pack(fill="x", pady=(0,2))
        ttk.Label(frame_outer, text="Outer Arena (cm):", background="#ffffff").pack(side="left")
        self.entry_outer_size = ttk.Entry(frame_outer, width=6, font=("Segoe UI", 10))
        self.entry_outer_size.insert(0, "96.0")
        self.entry_outer_size.pack(side="right")
        self.entry_outer_size.bind("<KeyRelease>", lambda e: self.draw_draggable_poly())

        # Inner Dimension (Bound to real-time updates)
        frame_inner = ttk.Frame(col1, style="OFT_Card.TFrame")
        frame_inner.pack(fill="x", pady=(0,4))
        ttk.Label(frame_inner, text="Inner Zone (cm):", background="#ffffff").pack(side="left")
        self.entry_inner_size = ttk.Entry(frame_inner, width=6, font=("Segoe UI", 10))
        self.entry_inner_size.insert(0, "48.0")
        self.entry_inner_size.pack(side="right")
        self.entry_inner_size.bind("<KeyRelease>", lambda e: self.draw_draggable_poly())
        
        # Grouped Calibration buttons horizontally
        btn_frame_1 = ttk.Frame(col1, style="OFT_Card.TFrame")
        btn_frame_1.pack(fill="x", pady=(0, 5))
        self.btn_calibrate = ttk.Button(btn_frame_1, text="Handles", style="OFT.TButton", command=self.toggle_calibration, state="disabled")
        self.btn_calibrate.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.btn_confirm_calib = ttk.Button(btn_frame_1, text="Lock", style="OFT.TButton", command=self.confirm_calibration, state="disabled")
        self.btn_confirm_calib.pack(side="left", fill="x", expand=True, padx=(2, 0))
        
        # -- Col 2: Tracking --
        col2 = ttk.Frame(control_panel, style="OFT_Card.TFrame")
        col2.grid(row=0, column=2, sticky="nsew", padx=5)
        ttk.Label(col2, text="3. Target Initialization", style="OFT_Header.TLabel").pack(anchor="w", pady=(0,4))
        
        # Grouped Box & Track buttons horizontally
        btn_frame_2 = ttk.Frame(col2, style="OFT_Card.TFrame")
        btn_frame_2.pack(fill="x", pady=(0, 5))
        self.btn_init_placement = ttk.Button(btn_frame_2, text="Draw Box", style="OFT.TButton", command=self.enable_placement_mode, state="disabled")
        self.btn_init_placement.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.btn_track = ttk.Button(btn_frame_2, text="Track", style="OFT.TButton", command=self.start_tracking, state="disabled")
        self.btn_track.pack(side="left", fill="x", expand=True, padx=(2, 0))
        
        # Tightened vertical spacing for sliders
        ttk.Label(col2, text="Inertia Alpha (Smoothing):", style="OFT_MetricTitle.TLabel").pack(anchor="w")
        self.scale_inertia = ttk.Scale(col2, from_=0.01, to=1.0, orient="horizontal", command=self.update_inertia)
        self.scale_inertia.pack(fill="x")
        self.lbl_inertia_val = ttk.Label(col2, text="0.10", style="OFT_Instruct.TLabel")
        self.lbl_inertia_val.pack(anchor="e", pady=(0,2))
        self.scale_inertia.set(0.1)
        
        ttk.Label(col2, text="Min Movement Deadzone:", style="OFT_MetricTitle.TLabel").pack(anchor="w")
        self.scale_deadzone = ttk.Scale(col2, from_=0.0, to=30.0, orient="horizontal", command=self.update_deadzone)
        self.scale_deadzone.pack(fill="x")
        self.lbl_deadzone_val = ttk.Label(col2, text="5.0 cm", style="OFT_Instruct.TLabel")
        self.lbl_deadzone_val.pack(anchor="e")
        self.scale_deadzone.set(5.0)
        
        # -- Col 3: Metrics & Export --
        col3 = ttk.Frame(control_panel, style="OFT_Card.TFrame")
        col3.grid(row=0, column=3, sticky="nsew", padx=5)
        ttk.Label(col3, text="4. Live Metrics", style="OFT_Header.TLabel").pack(anchor="w", pady=(0,4))
        
        dist_frame = ttk.Frame(col3, style="OFT_Card.TFrame")
        dist_frame.pack(fill="x", pady=(0, 2))
        ttk.Label(dist_frame, text="Distance:", style="OFT_MetricTitle.TLabel").pack(side="left")
        self.lbl_distance = ttk.Label(dist_frame, text="0.00 cm", style="OFT_MetricValue.TLabel", foreground="#d35400")
        self.lbl_distance.pack(side="right")
        
        time_frame = ttk.Frame(col3, style="OFT_Card.TFrame")
        time_frame.pack(fill="x", pady=(0, 6))
        ttk.Label(time_frame, text="Center Time:", style="OFT_MetricTitle.TLabel").pack(side="left")
        self.lbl_center_time = ttk.Label(time_frame, text="0.00 s", style="OFT_MetricValue.TLabel", foreground="#2980b9")
        self.lbl_center_time.pack(side="right")
        
        frame_export = ttk.Frame(col3, style="OFT_Card.TFrame")
        frame_export.pack(fill="x", pady=(0,5))
        self.btn_export_data = ttk.Button(frame_export, text="Save CSV", style="OFT.TButton", command=self.export_data, state="disabled")
        self.btn_export_data.pack(side="left", fill="x", expand=True, padx=(0,2))
        self.btn_export_map = ttk.Button(frame_export, text="Save Map", style="OFT.TButton", command=self.export_map, state="disabled")
        self.btn_export_map.pack(side="left", fill="x", expand=True, padx=(2,0))
        
        self.btn_reset_data = ttk.Button(col3, text="Reset Trial", style="OFT.TButton", command=self.reset_trial_data, state="disabled")
        self.btn_reset_data.pack(fill="x")

    def update_speed(self, val):
        self.playback_speed = float(val)
        if hasattr(self, 'lbl_speed_val'):
            self.lbl_speed_val.config(text=f"{self.playback_speed:.1f}x")

    def update_inertia(self, val):
        self.base_inertia = float(val)
        if hasattr(self, 'lbl_inertia_val'):
            self.lbl_inertia_val.config(text=f"{self.base_inertia:.2f}")

    def update_deadzone(self, val):
        self.velocity_deadzone_cm = float(val)
        if hasattr(self, 'lbl_deadzone_val'):
            self.lbl_deadzone_val.config(text=f"{self.velocity_deadzone_cm:.1f} cm")

    def load_video(self):
        self.video_path = filedialog.askopenfilename(filetypes=[("Video File", "*.mp4 *.avi *.mov")])
        if not self.video_path: return
            
        self.cap = cv2.VideoCapture(self.video_path)
        self.fps = self.cap.get(cv2.CAP_PROP_FPS) or 30.0
        self.orig_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        self.orig_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        
        if self.orig_w > 0:
            self.canvas_h = int(self.canvas_w * (self.orig_h / self.orig_w))
            self.canvas_video.config(width=self.canvas_w, height=self.canvas_h)
        
        self.reset_trial_data()
        self.lbl_instruct.config(text="Step 2: Define tracking bounds.")
        
        success, frame = self.cap.read()
        if success:
            self.current_frame = frame.copy()
            self.update_video_canvas(frame)
            self.btn_play_pause.config(state="normal")
            self.btn_calibrate.config(state="normal")
            self.btn_reset_data.config(state="normal")
            self.draw_empty_grid_map()
            self.video_loop()

    def reset_trial_data(self):
        self.total_distance_cm = 0.0
        self.center_frames = 0
        self.path_coordinates = []
        self.odometer_anchor_cm = None
        self.lbl_distance.config(text="0.00 cm")
        self.lbl_center_time.config(text="0.00 s")
        self.draw_empty_grid_map()

    def toggle_video_playback(self):
        self.video_playing = not self.video_playing
        
    def seek_video(self, frames_to_move):
        if self.cap is None: return
        
        # Pause tracking if actively seeking to prevent massive distance spikes
        if self.tracking_active:
            self.tracking_active = False
            self.video_playing = False
            self.lbl_instruct.config(text="Tracking paused due to seek. Please redraw box.", foreground="#e74c3c")
        
        current_pos = self.cap.get(cv2.CAP_PROP_POS_FRAMES)
        new_pos = max(0, current_pos + frames_to_move)
        
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_pos)
        success, frame = self.cap.read()
        if success:
            self.current_frame = frame.copy()
            self.update_video_canvas(frame)

    def video_loop(self):
        if self.cap is not None:
            if self.video_playing:
                success, frame = self.cap.read()
                if success:
                    self.current_frame = frame.copy()
                    if self.tracking_active:
                        self.execute_silhouette_tracker(frame)
                    else:
                        self.update_video_canvas(frame)
                else:
                    self.video_playing = False
                    self.tracking_active = False
            else:
                if self.current_frame is not None:
                    display_frame = self.current_frame.copy()
                    if self.track_window and not self.placement_mode:
                        x, y, w, h = self.track_window
                        cv2.rectangle(display_frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    self.update_video_canvas(display_frame)
                    
        # Calculate dynamic delay: 1000ms / FPS / PlaybackSpeed
        base_fps = self.fps if self.fps > 0 else 30.0
        loop_delay_ms = max(1, int((1000.0 / base_fps) / self.playback_speed))
        
        self.after(loop_delay_ms, self.video_loop)

    def enable_placement_mode(self):
        if self.video_playing: self.video_playing = False
        self.placement_mode = True
        self.track_window = None
        self.odometer_anchor_cm = None 
        self.lbl_instruct.config(text="Click & drag over the rat.", foreground="#2980b9")

    def on_canvas_click(self, event):
        sx, sy = self.orig_w / self.canvas_w, self.orig_h / self.canvas_h
        if self.placement_mode:
            self.box_start = (int(event.x * sx), int(event.y * sy))
            return
        if self.calibrating:
            self.active_drag_pt = None
            for key, pt in self.calib_pts.items():
                if math.hypot(event.x - pt[0], event.y - pt[1]) < 15:
                    self.active_drag_pt = key
                    break

    def on_canvas_drag(self, event):
        sx, sy = self.orig_w / self.canvas_w, self.orig_h / self.canvas_h
        if self.placement_mode and self.box_start:
            self.box_end = (int(event.x * sx), int(event.y * sy))
            preview = self.current_frame.copy()
            cv2.rectangle(preview, self.box_start, self.box_end, (255, 255, 0), 2)
            self.update_video_canvas(preview)
            return
        if self.calibrating and self.active_drag_pt:
            self.calib_pts[self.active_drag_pt] = [event.x, event.y]
            self.draw_draggable_poly()

    def on_canvas_release(self, event):
        if self.placement_mode and self.box_start and self.box_end:
            x1, y1 = self.box_start; x2, y2 = self.box_end
            x, y = min(x1, x2), min(y1, y2)
            w, h = abs(x1 - x2), abs(y1 - y2)
            
            if w > 10 and h > 10:
                self.track_window = (x, y, w, h)
                self.smoothed_cx = x + w // 2
                self.smoothed_cy = y + h // 2
                self.placement_mode = False
                self.lbl_instruct.config(text="Target registered! Ready to track.", foreground="#27ae60")
                self.btn_track.config(state="normal")
            return
        self.active_drag_pt = None

    def toggle_calibration(self):
        self.calibrating = True
        self.btn_confirm_calib.config(state="normal")
        
        # Initialize points dynamically at 20% inset from the canvas edges
        if not hasattr(self, 'calib_pts') or not self.calib_pts:
            pad_x = int(self.canvas_w * 0.2)
            pad_y = int(self.canvas_h * 0.2)
            
            self.calib_pts = {
                'TL': (pad_x, pad_y),                                 # Top-Left
                'TR': (self.canvas_w - pad_x, pad_y),                 # Top-Right
                'BR': (self.canvas_w - pad_x, self.canvas_h - pad_y), # Bottom-Right
                'BL': (pad_x, self.canvas_h - pad_y)                  # Bottom-Left
            }
            
        self.draw_draggable_poly()

    def draw_draggable_poly(self, event=None):
        self.canvas_video.delete("calib")
        if not getattr(self, 'calibrating', False): return
        
        pts = [self.calib_pts['TL'], self.calib_pts['TR'], self.calib_pts['BR'], self.calib_pts['BL']]
        mode = self.calib_mode.get()
        
        try:
            out_len = float(self.entry_outer_size.get())
            in_len = float(self.entry_inner_size.get())
            
            # Prevent math errors
            if in_len >= out_len or out_len <= 0 or in_len <= 0: return
            
            src_pts = np.array(pts, dtype=np.float32)
            offset = (out_len - in_len) / 2.0
            
            if mode == "Outer":
                # --- HANDLES CONTROL THE OUTER BOX ---
                self.canvas_video.create_polygon(pts[0][0], pts[0][1], pts[1][0], pts[1][1], 
                                                 pts[2][0], pts[2][1], pts[3][0], pts[3][1], 
                                                 outline="#00ffff", fill="", width=2, tags="calib")
                
                # Project Inner Box Inwards
                dst_pts = np.array([[0, 0], [out_len, 0], [out_len, out_len], [0, out_len]], dtype=np.float32)
                inv_matrix = np.linalg.inv(cv2.getPerspectiveTransform(src_pts, dst_pts))
                
                inner_dst_pts = np.array([[
                    [offset, offset], [out_len - offset, offset],
                    [out_len - offset, out_len - offset], [offset, out_len - offset]
                ]], dtype=np.float32)
                
                inner_src_pts = cv2.perspectiveTransform(inner_dst_pts, inv_matrix)[0]
                self.canvas_video.create_polygon(
                    inner_src_pts[0][0], inner_src_pts[0][1], inner_src_pts[1][0], inner_src_pts[1][1], 
                    inner_src_pts[2][0], inner_src_pts[2][1], inner_src_pts[3][0], inner_src_pts[3][1], 
                    outline="#ff00ff", fill="", width=2, dash=(4, 4), tags="calib")

            else:
                # --- HANDLES CONTROL THE INNER BOX ---
                self.canvas_video.create_polygon(pts[0][0], pts[0][1], pts[1][0], pts[1][1], 
                                                 pts[2][0], pts[2][1], pts[3][0], pts[3][1], 
                                                 outline="#ff00ff", fill="", width=2, dash=(4, 4), tags="calib")
                
                # Project Outer Box Outwards (Negative Offsets)
                dst_pts = np.array([[0, 0], [in_len, 0], [in_len, in_len], [0, in_len]], dtype=np.float32)
                inv_matrix = np.linalg.inv(cv2.getPerspectiveTransform(src_pts, dst_pts))
                
                outer_dst_pts = np.array([[
                    [-offset, -offset], [in_len + offset, -offset],
                    [in_len + offset, in_len + offset], [-offset, in_len + offset]
                ]], dtype=np.float32)
                
                outer_src_pts = cv2.perspectiveTransform(outer_dst_pts, inv_matrix)[0]
                self.canvas_video.create_polygon(
                    outer_src_pts[0][0], outer_src_pts[0][1], outer_src_pts[1][0], outer_src_pts[1][1], 
                    outer_src_pts[2][0], outer_src_pts[2][1], outer_src_pts[3][0], outer_src_pts[3][1], 
                    outline="#00ffff", fill="", width=2, tags="calib")
                    
            # Always draw the yellow draggable handles on top
            for key, pt in self.calib_pts.items():
                self.canvas_video.create_oval(pt[0]-6, pt[1]-6, pt[0]+6, pt[1]+6, fill="#f1c40f", outline="black", tags=("calib", key))
                
        except ValueError:
            pass # Fails silently if user deletes all text in the entry box

    def confirm_calibration(self):
        try:
            out_len = float(self.entry_outer_size.get())
            in_len = float(self.entry_inner_size.get())
            
            if in_len >= out_len or out_len <= 0 or in_len <= 0:
                return
                
            # Internal reference size for our coordinate mapping grid
            map_w = self.map_size # 400
            
            # Calculate dynamic scale factor (Real-world cm per internal pixel)
            self.scale_factor = out_len / map_w
            
            # Dynamically compute pixel boundaries for the center zone tracking limits
            cm_offset = (out_len - in_len) / 2.0
            self.inner_min_px = cm_offset / self.scale_factor
            self.inner_max_px = map_w - self.inner_min_px
            
            # Extract source coordinates from our current yellow dragging handles
            pts = [self.calib_pts['TL'], self.calib_pts['TR'], self.calib_pts['BR'], self.calib_pts['BL']]
            src_pts = np.array(pts, dtype=np.float32)
            
            mode = self.calib_mode.get()
            if mode == "Outer":
                # Handles represent outer edge -> Map handles to total map width
                dst_pts = np.array([
                    [0, 0],
                    [map_w, 0],
                    [map_w, map_w],
                    [0, map_w]
                ], dtype=np.float32)
            else:
                # Handles represent inner edge -> Map handles to calculated inner limits
                dst_pts = np.array([
                    [self.inner_min_px, self.inner_min_px],
                    [self.inner_max_px, self.inner_min_px],
                    [self.inner_max_px, self.inner_max_px],
                    [self.inner_min_px, self.inner_max_px]
                ], dtype=np.float32)
                
            # Compute the final transformation matrix
            self.homography_matrix = cv2.getPerspectiveTransform(src_pts, dst_pts)
            
            # Shut down calibration overlays
            self.calibrating = False
            self.canvas_video.delete("calib")
            
            # Toggle control button availability
            self.btn_confirm_calib.config(state="disabled")
            self.btn_init_placement.config(state="normal")
            
            if hasattr(self, 'lbl_instruct'):
                self.lbl_instruct.config(text="Step 3: Click 'Draw Box' and select the target animal.")
                
            # Re-draw the trajectory canvas grid to show true box proportions
            self.draw_empty_grid_map()
            
        except ValueError:
            pass

    def start_tracking(self):
        if self.track_window is None: return
        
        self.bg_subtractor = cv2.createBackgroundSubtractorMOG2(history=300, varThreshold=24, detectShadows=True)
        for _ in range(10):
            self.bg_subtractor.apply(self.current_frame)
            
        self.btn_export_data.config(state="normal")
        self.btn_export_map.config(state="normal")
        self.lbl_instruct.config(text="Tracking Live.", foreground="#e67e22")
        
        self.tracking_active = True
        self.video_playing = True

    def execute_silhouette_tracker(self, frame):
        if self.smoothed_cx is None or self.smoothed_cy is None: return
        
        # Scale inertia based on speed so bounding box catches up during fast forward
        # Cap at 1.0 to prevent jumping past the tracking point
        dynamic_inertia = min(1.0, self.base_inertia * self.playback_speed)
        
        fg_mask = self.bg_subtractor.apply(frame)
        fg_mask[fg_mask == 127] = 0
        _, fg_mask = cv2.threshold(fg_mask, 200, 255, cv2.THRESH_BINARY)
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
        fg_mask = cv2.morphologyEx(fg_mask, cv2.MORPH_CLOSE, kernel)
        
        h, w = fg_mask.shape
        r = 120
        x_min, x_max = max(0, int(self.smoothed_cx - r)), min(w, int(self.smoothed_cx + r))
        y_min, y_max = max(0, int(self.smoothed_cy - r)), min(h, int(self.smoothed_cy + r))
        
        local_mask = fg_mask[y_min:y_max, x_min:x_max]
        if local_mask.size == 0: return
        
        cnts, _ = cv2.findContours(local_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        best_contour = None
        max_area = 0
        for c in cnts:
            area = cv2.contourArea(c)
            if area > max_area and area > 180:
                max_area = area
                best_contour = c
                
        if best_contour is not None:
            M = cv2.moments(best_contour)
            if M["m00"] != 0:
                raw_cx = (M["m10"] / M["m00"]) + x_min
                raw_cy = (M["m01"] / M["m00"]) + y_min
                
                # Apply dynamic inertia scaling
                self.smoothed_cx = dynamic_inertia * raw_cx + (1 - dynamic_inertia) * self.smoothed_cx
                self.smoothed_cy = dynamic_inertia * raw_cy + (1 - dynamic_inertia) * self.smoothed_cy
                
                bw, bh = self.track_window[2], self.track_window[3]
                self.track_window = (int(self.smoothed_cx - bw//2), int(self.smoothed_cy - bh//2), bw, bh)
                
                cv2.rectangle(frame, (self.track_window[0], self.track_window[1]), 
                              (self.track_window[0]+bw, self.track_window[1]+bh), (0, 255, 255), 2)
                cv2.circle(frame, (int(self.smoothed_cx), int(self.smoothed_cy)), 6, (0, 0, 255), -1)
                
                self.map_metrics_to_arena((self.smoothed_cx, self.smoothed_cy))
        else:
            cv2.rectangle(frame, (self.track_window[0], self.track_window[1]), 
                          (self.track_window[0]+self.track_window[2], self.track_window[1]+self.track_window[3]), (255, 0, 0), 2)
            cv2.circle(frame, (int(self.smoothed_cx), int(self.smoothed_cy)), 6, (0, 0, 255), -1)
                
        self.update_video_canvas(frame)
        self.render_trajectory_view()

    def map_metrics_to_arena(self, target_pt):
        if not hasattr(self, 'homography_matrix') or self.homography_matrix is None: 
            return
            
        # --- FIX: Scale original video coordinates to UI Canvas coordinates ---
        if hasattr(self, 'orig_w') and self.orig_w > 0:
            scale_x = self.canvas_w / self.orig_w
            scale_y = self.canvas_h / self.orig_h
        else:
            scale_x, scale_y = 1.0, 1.0
            
        canvas_x = target_pt[0] * scale_x
        canvas_y = target_pt[1] * scale_y
        
        # 1. Transform Scaled Coordinates to Map Coordinates
        pt_arr = np.array([[[float(canvas_x), float(canvas_y)]]], dtype=np.float32)
        warped = cv2.perspectiveTransform(pt_arr, self.homography_matrix)
        px, py = warped[0][0][0], warped[0][0][1]
        
        # Calculate real-world position (cm)
        pos_cm = (px * self.scale_factor, py * self.scale_factor)
        
        # --- VISUAL PATH: Always update for a perfectly smooth line ---
        if not hasattr(self, 'path_coordinates'): 
            self.path_coordinates = []
            
        if 0 <= px < self.map_size and 0 <= py < self.map_size:
            if not self.path_coordinates or self.path_coordinates[-1] != (int(px), int(py)):
                self.path_coordinates.append((int(px), int(py)))
                
        # --- DISTANCE ODOMETER: Strictly apply the deadzone noise filter ---
        if getattr(self, 'odometer_anchor_cm', None) is None:
            self.odometer_anchor_cm = pos_cm
            self.total_distance_cm = 0.0
        else:
            step_distance = math.hypot(pos_cm[0] - self.odometer_anchor_cm[0], pos_cm[1] - self.odometer_anchor_cm[1])
            if step_distance >= self.velocity_deadzone_cm:
                self.total_distance_cm += step_distance
                self.lbl_distance.config(text=f"{self.total_distance_cm:.2f} cm")
                self.odometer_anchor_cm = pos_cm
                
        # --- CENTER TIME LOGIC (Dynamically Scaled Boundaries) ---
        min_px = getattr(self, 'inner_min_px', 100)
        max_px = getattr(self, 'inner_max_px', 300)
        
        if min_px <= px <= max_px and min_px <= py <= max_px:
            if not hasattr(self, 'center_frames'): self.center_frames = 0
            self.center_frames += 1
            if hasattr(self, 'fps') and self.fps > 0:
                self.lbl_center_time.config(text=f"{(self.center_frames / self.fps):.2f} s")
            
        # --- DRAWING: Paint the trajectory onto the canvas ---
        if len(self.path_coordinates) > 1:
            # BGR format: Blue line to match the example image trajectory
            cv2.line(self.trajectory_canvas, self.path_coordinates[-2], self.path_coordinates[-1], (200, 50, 50), 2)


    def draw_empty_grid_map(self):
        self.trajectory_canvas = np.zeros((self.map_size, self.map_size, 3), dtype=np.uint8)
        self.trajectory_canvas.fill(250) 
        step = self.map_size // 4
        
        # Draw base grid
        for i in range(1, 4):
            cv2.line(self.trajectory_canvas, (i * step, 0), (i * step, self.map_size), (220, 220, 220), 1)
            cv2.line(self.trajectory_canvas, (0, i * step), (self.map_size, i * step), (220, 220, 220), 1)
            
        # Draw dynamic center zone box (Fixed to Pink in OpenCV BGR format: 180 Blue, 180 Green, 255 Red)
        min_px = int(getattr(self, 'inner_min_px', 100))
        max_px = int(getattr(self, 'inner_max_px', 300))
        cv2.rectangle(self.trajectory_canvas, (min_px, min_px), (max_px, max_px), (180, 180, 255), 2) 
        
        self.display_image_on_label(self.lbl_path, self.trajectory_canvas, self.map_display_size, self.map_display_size)

    def render_trajectory_view(self):
        temp_map = self.trajectory_canvas.copy()
        
        # Draw the red dot at the rat's current location
        if hasattr(self, 'path_coordinates') and self.path_coordinates:
            cv2.circle(temp_map, self.path_coordinates[-1], 6, (0, 0, 255), -1) 
            
        self.display_image_on_label(self.lbl_path, temp_map, self.map_display_size, self.map_display_size)

    def update_video_canvas(self, opencv_frame):
        resized = cv2.resize(opencv_frame, (self.canvas_w, self.canvas_h))
        rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)
        img_tk = ImageTk.PhotoImage(Image.fromarray(rgb))
        self.canvas_video.create_image(0, 0, anchor="nw", image=img_tk, tags="video")
        self.canvas_video.image = img_tk 
        if self.calibrating: self.draw_draggable_poly()

    def display_image_on_label(self, label_widget, opencv_frame, w, h):
        resized = cv2.resize(opencv_frame, (w, h))
        rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)
        img_tk = ImageTk.PhotoImage(Image.fromarray(rgb))
        label_widget.config(image=img_tk)
        label_widget.image = img_tk

    def export_data(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Excel CSV File", "*.csv")])
        if filename:
            try:
                with open(filename, mode='w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Metric", "Value"])
                    writer.writerow(["Total Distance (cm)", round(self.total_distance_cm, 2)])
                    writer.writerow(["Center Zone Time (s)", round(self.center_frames / self.fps, 2)])
                messagebox.showinfo("Export Successful", f"Data saved to:\n{filename}")
            except Exception as e:
                messagebox.showerror("Export Failed", f"Could not save file:\n{e}")

    def export_map(self):
        filename = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg")])
        if filename and self.trajectory_canvas is not None:
            cv2.imwrite(filename, self.trajectory_canvas)
            messagebox.showinfo("Export Successful", f"Map image saved to:\n{filename}")