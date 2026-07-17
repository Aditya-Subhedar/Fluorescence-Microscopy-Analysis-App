import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import numpy as np
import pandas as pd
import itertools
import scipy.stats as stats

# ==========================================
# CRITICAL WINDOWS HIGH-DPI FIX
# ==========================================
if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

class GraphCreationTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#f8f9fa")
        
        self.graph_zoom = 1.0
        self.graph_pan_x = 0.0
        self.graph_pan_y = 0.0
        self.is_panning = False
        
        self.data_df = None
        self.file_mappings = {} 
        
        self.setup_ui()

    def setup_ui(self):
        self.pack(fill=tk.BOTH, expand=True)

        self.paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.left_container = tk.Frame(self.paned_window, width=400, bg="#f0f0f0")
        self.paned_window.add(self.left_container, weight=0)

        self.canvas_left = tk.Canvas(self.left_container, bg="#f0f0f0", highlightthickness=0)
        self.scrollbar_left = ttk.Scrollbar(self.left_container, orient="vertical", command=self.canvas_left.yview)
        
        self.scrollable_frame = tk.Frame(self.canvas_left, bg="#f0f0f0")
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas_left.configure(scrollregion=self.canvas_left.bbox("all"))
        )

        self.canvas_window = self.canvas_left.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas_left.bind('<Configure>', self._on_canvas_configure)

        self.canvas_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_left.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas_left.configure(yscrollcommand=self.scrollbar_left.set)

        self._build_control_widgets(self.scrollable_frame)
        
        self.bind_left_scroll_recursive(self.left_container)

        self.right_panel = tk.Frame(self.paned_window, bg="white")
        self.paned_window.add(self.right_panel, weight=1)

        self.figure = Figure(figsize=(6, 4), dpi=100)
        self.figure.patch.set_facecolor('#ffffff')
        self.ax = self.figure.add_subplot(111)
        
        self.mpl_canvas = FigureCanvasTkAgg(self.figure, master=self.right_panel)
        self.canvas = self.mpl_canvas.get_tk_widget()
        
        self.figure.canvas.mpl_connect('scroll_event', self.on_zoom_layout)
        self.figure.canvas.mpl_connect('button_press_event', self.on_pan_start)
        self.figure.canvas.mpl_connect('motion_notify_event', self.on_pan_move)
        self.figure.canvas.mpl_connect('button_release_event', self.on_pan_end)
        
        self.toolbar_frame = tk.Frame(self.right_panel, bg="#f0f0f0")
        self.toolbar_frame.pack(side=tk.TOP, fill=tk.X)
        self.toolbar = NavigationToolbar2Tk(self.mpl_canvas, self.toolbar_frame)
        self.toolbar.update()
        
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.generate_plot()

    def bind_left_scroll_recursive(self, widget):
        widget.bind("<MouseWheel>", self._on_left_panel_scroll)
        widget.bind("<Button-4>", self._on_left_panel_scroll)
        widget.bind("<Button-5>", self._on_left_panel_scroll)
        for child in widget.winfo_children():
            self.bind_left_scroll_recursive(child)

    def _on_left_panel_scroll(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas_left.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas_left.yview_scroll(1, "units")

    def _on_canvas_configure(self, event):
        self.canvas_left.itemconfig(self.canvas_window, width=event.width)

    def on_zoom_layout(self, event):
        if event.step > 0:
            self.graph_zoom *= 1.15
        else:
            self.graph_zoom /= 1.15
        self.graph_zoom = max(0.15, min(self.graph_zoom, 6.0))
        self.apply_layout_transform()

    def on_pan_start(self, event):
        if event.button == 1:
            self.is_panning = True
            self.start_pixel_x = event.x
            self.start_pixel_y = event.y
            self.orig_pan_x = self.graph_pan_x
            self.orig_pan_y = self.graph_pan_y

    def on_pan_move(self, event):
        if self.is_panning and event.x is not None and event.y is not None:
            dx_pixels = event.x - self.start_pixel_x
            dy_pixels = event.y - self.start_pixel_y
            fig_w = self.figure.bbox.width
            fig_h = self.figure.bbox.height
            self.graph_pan_x = self.orig_pan_x + (dx_pixels / fig_w)
            self.graph_pan_y = self.orig_pan_y + (dy_pixels / fig_h)
            self.apply_layout_transform()

    def on_pan_end(self, event):
        self.is_panning = False

    def apply_layout_transform(self):
        leg_pos = self.var_legend_pos.get() if hasattr(self, 'var_legend_pos') else "outside right"
        if leg_pos == "outside right":
            base_left, base_right = 0.16, 0.75
        else:
            base_left, base_right = 0.16, 0.92
            
        base_bottom, base_top = 0.18, 0.88
        base_w = base_right - base_left
        base_h = base_top - base_bottom
        
        cx = ((base_left + base_right) / 2) + self.graph_pan_x
        cy = ((base_bottom + base_top) / 2) + self.graph_pan_y
        
        w = base_w * self.graph_zoom
        h = base_h * self.graph_zoom
        
        left = cx - w / 2
        right = cx + w / 2
        bottom = cy - h / 2
        top = cy + h / 2
        
        if left >= right: right = left + 0.01
        if bottom >= top: top = bottom + 0.01
        
        try:
            self.figure.subplots_adjust(left=left, right=right, bottom=bottom, top=top)
            self.mpl_canvas.draw_idle()
        except Exception:
            pass

    def toggle_input_mode(self):
        mode = self.data_mode.get()
        if mode == "manual":
            self.file_frame.pack_forget()
            self.manual_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        else:
            self.manual_frame.pack_forget()
            self.file_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.generate_plot()

    def upload_data_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Data File (CSV or Excel)",
            filetypes=[("Data Files", "*.csv *.xlsx *.xls"), ("All Files", "*.*")]
        )
        if not file_path:
            return
        
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                xl = pd.ExcelFile(file_path)
                sheet_name = 'Global Overview' if 'Global Overview' in xl.sheet_names else xl.sheet_names[0]
                df = xl.parse(sheet_name)
            
            if 'File Name' not in df.columns:
                messagebox.showerror("Format Error", "The selected file does not contain a 'File Name' column.\nPlease import a properly formatted file.")
                return
            
            self.data_df = df
            self.lbl_file_status.config(text=f"Loaded: {os.path.basename(file_path)}", fg="#2e7d32")
            self.btn_map.config(state="normal")
            
            numeric_cols = list(df.select_dtypes(include=[np.number]).columns)
            excluded_keys = ['threshold', 'range', 'limit', 'strip', 'parameter']
            filtered_cols = [col for col in numeric_cols if not any(k in col.lower() for k in excluded_keys)]
            
            if not filtered_cols:
                filtered_cols = numeric_cols if numeric_cols else ["No numeric metrics found"]
            
            self.cb_y_metric.config(values=filtered_cols)
            
            default_col = None
            for col in filtered_cols:
                if "fluorescent" in col.lower() or "area" in col.lower() or "intensity" in col.lower():
                    default_col = col
                    break
            if not default_col and filtered_cols:
                default_col = filtered_cols[0]
                
            self.var_y_metric.set(default_col if default_col else "")
            
            # Start with blank mappings
            self.file_mappings = {}
            unique_files = list(df['File Name'].dropna().unique())
            for fn in unique_files:
                self.file_mappings[fn] = {'category': '', 'subgroup': ''}
                
            self.open_mapping_dialog()
            
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not read the selected file:\n{str(e)}")

    def open_mapping_dialog(self):
        if self.data_df is None:
            return
        
        dialog = tk.Toplevel(self)
        dialog.title("Configure Experimental Grouping & Mappings")
        
        # Set a sensible default size and restrict minimum resizing so it stays readable
        dialog.geometry("1000x700")
        dialog.minsize(850, 450)
        dialog.transient(self)
        dialog.grab_set()
        
        # --- 1. Variables (Declared early for closure scoping) ---
        var_delim = tk.StringVar(value="_")
        var_cat_idx = tk.IntVar(value=1)
        var_sub_idx = tk.IntVar(value=2)
        var_batch_cat = tk.StringVar()
        var_batch_sub = tk.StringVar()
        row_widgets = []
        
        # --- 2. Logic Callback Functions ---
        def select_all():
            for r in row_widgets: 
                r['var_select'].set(True)
                
        def deselect_all():
            for r in row_widgets: 
                r['var_select'].set(False)

        def apply_auto_parse():
            delim = var_delim.get()
            cat_pos = var_cat_idx.get() - 1
            sub_pos = var_sub_idx.get() - 1
            for r in row_widgets:
                fn_raw = os.path.splitext(r['filename'])[0]
                parts = fn_raw.split(delim)
                
                if 0 <= cat_pos < len(parts):
                    r['ent_cat'].delete(0, tk.END)
                    r['ent_cat'].insert(0, parts[cat_pos].strip())
                if 0 <= sub_pos < len(parts):
                    r['ent_sub'].delete(0, tk.END)
                    r['ent_sub'].insert(0, parts[sub_pos].strip())
                    
        def apply_batch():
            b_cat = var_batch_cat.get().strip()
            b_sub = var_batch_sub.get().strip()
            for r in row_widgets:
                if r['var_select'].get():
                    if b_cat:
                        r['ent_cat'].delete(0, tk.END)
                        r['ent_cat'].insert(0, b_cat)
                    if b_sub:
                        r['ent_sub'].delete(0, tk.END)
                        r['ent_sub'].insert(0, b_sub)
                        
        def save_and_close():
            for r in row_widgets:
                fn = r['filename']
                cat = r['ent_cat'].get().strip()
                sub = r['ent_sub'].get().strip()
                
                self.file_mappings[fn] = {
                    'category': cat if cat else "Uncategorized", 
                    'subgroup': sub if sub else "Ungrouped"
                }
            self.generate_plot()
            dialog.destroy()

        # --- 3. UI Construction ---
        
        # 1. TOP FRAME (LabelFrame using stacked sub-rows)
        top_frame = ttk.LabelFrame(dialog, text="⚡ Smart Helper & Batch Assigner")
        top_frame.pack(fill=tk.X, side=tk.TOP, padx=10, pady=5)
        
        # Row 1 Frame: Auto-Parsing
        row1_frame = tk.Frame(top_frame)
        row1_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(row1_frame, text="Delimiter:").pack(side=tk.LEFT, padx=2)
        ent_delim = tk.Entry(row1_frame, textvariable=var_delim, width=4)
        ent_delim.pack(side=tk.LEFT, padx=5)
        
        tk.Label(row1_frame, text="Cat Seg Index:").pack(side=tk.LEFT, padx=2)
        spn_cat = tk.Spinbox(row1_frame, from_=1, to=10, width=3, textvariable=var_cat_idx)
        spn_cat.pack(side=tk.LEFT, padx=5)
        
        tk.Label(row1_frame, text="Sub Seg Index:").pack(side=tk.LEFT, padx=2)
        spn_sub = tk.Spinbox(row1_frame, from_=1, to=10, width=3, textvariable=var_sub_idx)
        spn_sub.pack(side=tk.LEFT, padx=5)
        
        btn_parse = tk.Button(row1_frame, text="Auto-Parse Filenames 🧪", bg="#bbdefb", command=apply_auto_parse)
        btn_parse.pack(side=tk.LEFT, padx=15)
        
        # Row 2 Frame: Selection & Batch Customizer
        row2_frame = tk.Frame(top_frame)
        row2_frame.pack(fill=tk.X, padx=10, pady=5)
        
        btn_sel_all = tk.Button(row2_frame, text="Select All", command=select_all)
        btn_sel_all.pack(side=tk.LEFT, padx=2)
        btn_desel_all = tk.Button(row2_frame, text="Deselect All", command=deselect_all)
        btn_desel_all.pack(side=tk.LEFT, padx=2)
        
        # Spacer Line to separate actions
        tk.Label(row2_frame, text=" |  Batch Assign -> ", fg="gray").pack(side=tk.LEFT, padx=8)
        
        tk.Label(row2_frame, text="Cat:").pack(side=tk.LEFT, padx=2)
        ent_b_cat = tk.Entry(row2_frame, textvariable=var_batch_cat, width=12)
        ent_b_cat.pack(side=tk.LEFT, padx=5)
        
        tk.Label(row2_frame, text="Sub:").pack(side=tk.LEFT, padx=2)
        ent_b_sub = tk.Entry(row2_frame, textvariable=var_batch_sub, width=12)
        ent_b_sub.pack(side=tk.LEFT, padx=5)
        
        btn_apply_batch = tk.Button(row2_frame, text="Apply to Checked ✔", bg="#c8e6c9", command=apply_batch)
        btn_apply_batch.pack(side=tk.LEFT, padx=10)

        # 2. BOTTOM Frame Packed Second (Locks the save button securely to the bottom)
        bot_frame = tk.Frame(dialog)
        bot_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
            
        btn_save = tk.Button(bot_frame, text="💾 Save Groupings & Generate Plot", bg="#2e7d32", fg="white",
                             font=("Arial", 10, "bold"), command=save_and_close)
        btn_save.pack(pady=5)

        # 3. MIDDLE Frame Packed Last (Dynamically consumes all remaining window space)
        list_container = tk.Frame(dialog)
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(list_container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        scroll_content = tk.Frame(canvas)
        
        scroll_content.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        unique_files = list(self.data_df['File Name'].dropna().unique())
        
        headers = ["Select", "Image Source Filename", "Category (X-Axis Variable)", "Subgroup (Legend Item)"]
        for col_idx, text in enumerate(headers):
            lbl = tk.Label(scroll_content, text=text, font=("Arial", 10, "bold"))
            lbl.grid(row=0, column=col_idx, padx=10, pady=8, sticky="w")
            
        for idx, filename in enumerate(unique_files):
            r_idx = idx + 1
            
            var_sel = tk.BooleanVar(value=False)
            chk = tk.Checkbutton(scroll_content, variable=var_sel)
            chk.grid(row=r_idx, column=0, padx=5, pady=2)
            
            lbl_file = tk.Label(scroll_content, text=filename, font=("Consolas", 9), anchor="w")
            lbl_file.grid(row=r_idx, column=1, padx=5, pady=2, sticky="w")
            
            curr_cat = self.file_mappings.get(filename, {}).get('category', '')
            ent_cat = tk.Entry(scroll_content, width=25)
            ent_cat.insert(0, curr_cat)
            ent_cat.grid(row=r_idx, column=2, padx=5, pady=2, sticky="w")
            
            curr_sub = self.file_mappings.get(filename, {}).get('subgroup', '')
            ent_sub = tk.Entry(scroll_content, width=25)
            ent_sub.insert(0, curr_sub)
            ent_sub.grid(row=r_idx, column=3, padx=5, pady=2, sticky="w")
            
            row_widgets.append({
                'filename': filename,
                'var_select': var_sel,
                'ent_cat': ent_cat,
                'ent_sub': ent_sub
            })

    def parse_data_matrix(self):
        mode = self.data_mode.get()
        if mode == "manual":
            raw_text = self.txt_data.get("1.0", tk.END).strip().split('\n')
            parsed_records = []
            for line in raw_text:
                line = line.strip()
                if not line or line.startswith('#'): continue
                parts = line.split('|')
                if len(parts) < 3: continue
                category, subgroup = parts[0].strip(), parts[1].strip()
                try:
                    replicates = [float(x.strip()) for x in parts[2].split(',') if x.strip()]
                    for value in replicates:
                        parsed_records.append({"Category": category, "Subgroup": subgroup, "Value": value})
                except ValueError:
                    continue
            return pd.DataFrame(parsed_records)
        else:
            if self.data_df is None or not self.file_mappings:
                return pd.DataFrame()
            
            metric_col = self.var_y_metric.get()
            if not metric_col:
                return pd.DataFrame()
                
            parsed_records = []
            for _, row in self.data_df.iterrows():
                fn = row.get('File Name')
                val = row.get(metric_col)
                if pd.isna(val) or fn is None:
                    continue
                
                mapping = self.file_mappings.get(fn)
                if mapping:
                    parsed_records.append({
                        "Category": mapping['category'],
                        "Subgroup": mapping['subgroup'],
                        "Value": float(val)
                    })
            return pd.DataFrame(parsed_records)

    def get_color_cycle(self, subgroups_count):
        mode = self.var_palette.get()
        if mode == "Prism Bright":
            palette = ["#1abc9c", "#2ecc71", "#3498db", "#9b59b6", "#f1c40f", "#e67e22"]
        elif mode == "Seaborn Muted":
            palette = ["#4878d0", "#ee854a", "#6acc64", "#d65f5f", "#956cb4", "#8c613c"]
        elif mode == "Classic Grayscale":
            palette = ["#333333", "#666666", "#999999", "#cccccc", "#eeeeee"]
        else:
            palette = [c.strip() for c in self.var_hex_colors.get().split(',') if c.strip()]
                
        if not palette: palette = ["#4878d0", "#ee854a", "#6acc64", "#d65f5f", "#956cb4"]
        return list(itertools.islice(itertools.cycle(palette), subgroups_count))

    def _build_control_widgets(self, parent):
        frame_data = ttk.LabelFrame(parent, text="Step 1: Data Input & Setup")
        frame_data.pack(fill=tk.X, padx=10, pady=5)
        
        self.data_mode = tk.StringVar(value="manual")
        toggle_frame = tk.Frame(frame_data, bg="#f0f0f0")
        toggle_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Radiobutton(toggle_frame, text="✏ Manual Text", variable=self.data_mode, 
                       value="manual", command=self.toggle_input_mode, bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(toggle_frame, text="📊 File Import (CSV/Excel)", variable=self.data_mode, 
                       value="file", command=self.toggle_input_mode, bg="#f0f0f0").pack(side=tk.LEFT, padx=5)

        self.input_container = tk.Frame(frame_data, bg="#f0f0f0")
        self.input_container.pack(fill=tk.BOTH, expand=True)

        self.manual_frame = tk.Frame(self.input_container, bg="#f0f0f0")
        tk.Label(self.manual_frame, text="Format: Category | Subgroup | Replicates (comma-sep)", 
                 font=("Arial", 8, "italic"), fg="#555", bg="#f0f0f0").pack(anchor="w", padx=5, pady=(2,5))
        
        self.txt_data = tk.Text(self.manual_frame, height=8, width=42, font=("Consolas", 9))
        self.txt_data.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # REMOVED default manual data injection here!

        self.file_frame = tk.Frame(self.input_container, bg="#f0f0f0")
        btn_upload = tk.Button(self.file_frame, text="📂 Upload Data File (Excel/CSV)", bg="#e65100", fg="white", 
                               font=("Arial", 9, "bold"), command=self.upload_data_file)
        btn_upload.pack(fill=tk.X, padx=5, pady=5)
        
        self.lbl_file_status = tk.Label(self.file_frame, text="No file loaded", font=("Arial", 8, "italic"), fg="#777", bg="#f0f0f0")
        self.lbl_file_status.pack(anchor="w", padx=5)
        
        tk.Label(self.file_frame, text="Select Metric (Y-axis Value):", bg="#f0f0f0", font=("Arial", 9, "bold")).pack(anchor="w", padx=5, pady=(5, 2))
        self.var_y_metric = tk.StringVar()
        self.cb_y_metric = ttk.Combobox(self.file_frame, textvariable=self.var_y_metric, state="readonly")
        self.cb_y_metric.pack(fill=tk.X, padx=5, pady=2)
        self.cb_y_metric.bind("<<ComboboxSelected>>", lambda e: self.generate_plot())
        
        self.btn_map = tk.Button(self.file_frame, text="🔗 Configure Group Mappings", state="disabled", bg="#0288d1", fg="white", 
                                 font=("Arial", 9, "bold"), command=self.open_mapping_dialog)
        self.btn_map.pack(fill=tk.X, padx=5, pady=10)

        self.toggle_input_mode()

        frame_labels = ttk.LabelFrame(parent, text="Step 2: Axis & Titles")
        frame_labels.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(frame_labels, text="Graph Title:").pack(anchor="w", padx=5)
        self.var_title = tk.StringVar(value="Experimental Results")
        tk.Entry(frame_labels, textvariable=self.var_title).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_labels, text="Y-Axis Title:").pack(anchor="w", padx=5)
        self.var_y_title = tk.StringVar(value="Measured Value")
        tk.Entry(frame_labels, textvariable=self.var_y_title).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_labels, text="X-Axis Title:").pack(anchor="w", padx=5)
        self.var_x_title = tk.StringVar(value="Categories")
        tk.Entry(frame_labels, textvariable=self.var_x_title).pack(fill=tk.X, padx=5, pady=2)

        frame_type = ttk.LabelFrame(parent, text="Step 3: Graph Type & Style")
        frame_type.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_type, text="Chart Mode:").pack(anchor="w", padx=5)
        self.var_chart_mode = tk.StringVar(value="Grouped Bar + Scatter (Prism)")
        ttk.Combobox(frame_type, textvariable=self.var_chart_mode, state="readonly",
                     values=["Grouped Bar + Scatter (Prism)", "Grouped Bar (Mean + Error Only)", 
                             "Box Plot (Grouped)", "Violin Plot (Grouped)"]).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_type, text="Error Bar Metric:").pack(anchor="w", padx=5)
        self.var_error_type = tk.StringVar(value="SEM (Standard Error)")
        ttk.Combobox(frame_type, textvariable=self.var_error_type, state="readonly",
                     values=["SD (Standard Deviation)", "SEM (Standard Error)", "None"]).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_type, text="Replicate Point Jitter:").pack(anchor="w", padx=5)
        self.var_jitter = tk.DoubleVar(value=0.02)
        tk.Scale(frame_type, from_=0.0, to=0.1, resolution=0.01, orient=tk.HORIZONTAL, variable=self.var_jitter).pack(fill=tk.X, padx=5, pady=2)

        frame_colors = ttk.LabelFrame(parent, text="Step 4: Bar Colors & Textures")
        frame_colors.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_colors, text="Color Palette:").pack(anchor="w", padx=5)
        self.var_palette = tk.StringVar(value="Seaborn Muted")
        ttk.Combobox(frame_colors, textvariable=self.var_palette, state="readonly",
                     values=["Prism Bright", "Seaborn Muted", "Classic Grayscale", "Custom HEX List"]).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_colors, text="Custom Colors (Comma sep):").pack(anchor="w", padx=5)
        self.var_hex_colors = tk.StringVar(value="")
        tk.Entry(frame_colors, textvariable=self.var_hex_colors).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_colors, text="Hatch Patterns (Comma sep):").pack(anchor="w", padx=5)
        self.var_hatches = tk.StringVar(value="")
        tk.Entry(frame_colors, textvariable=self.var_hatches).pack(fill=tk.X, padx=5, pady=2)

        frame_geom = ttk.LabelFrame(parent, text="Step 5: Dimensional Settings")
        frame_geom.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_geom, text="Bar Width:").pack(anchor="w", padx=5)
        self.var_bar_width = tk.DoubleVar(value=0.20)
        tk.Spinbox(frame_geom, from_=0.05, to=0.5, increment=0.05, textvariable=self.var_bar_width).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_geom, text="Group Spacing:").pack(anchor="w", padx=5)
        self.var_spacing = tk.DoubleVar(value=0.05)
        tk.Spinbox(frame_geom, from_=0.0, to=0.3, increment=0.01, textvariable=self.var_spacing).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_geom, text="Y-Axis Limits (Min, Max or blank):").pack(anchor="w", padx=5)
        self.var_y_limits = tk.StringVar(value="")
        tk.Entry(frame_geom, textvariable=self.var_y_limits).pack(fill=tk.X, padx=5, pady=2)

        frame_fonts = ttk.LabelFrame(parent, text="Step 6: Typography & Layout")
        frame_fonts.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_fonts, text="Font Family:").pack(anchor="w", padx=5)
        self.var_font = tk.StringVar(value="Arial")
        ttk.Combobox(frame_fonts, textvariable=self.var_font, values=["Arial", "Times New Roman", "Helvetica", "Courier New"]).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_fonts, text="Font Sizes (Title, Axis, Ticks):").pack(anchor="w", padx=5)
        self.var_font_sizes = tk.StringVar(value="14,12,10")
        tk.Entry(frame_fonts, textvariable=self.var_font_sizes).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_fonts, text="Legend Position:").pack(anchor="w", padx=5)
        self.var_legend_pos = tk.StringVar(value="outside right")
        ttk.Combobox(frame_fonts, textvariable=self.var_legend_pos, state="readonly",
                     values=["outside right", "upper right", "upper left", "lower right", "lower left", "best", "None"]).pack(fill=tk.X, padx=5, pady=2)

        frame_sig = ttk.LabelFrame(parent, text="Step 7: Automated Statistics")
        frame_sig.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_sig, text="Significance Testing (α = 0.05):").pack(anchor="w", padx=5)
        self.var_stats_mode = tk.StringVar(value="T-Test (Compare to 1st Subgroup)")
        ttk.Combobox(frame_sig, textvariable=self.var_stats_mode, state="readonly",
                     values=[
                         "None", 
                         "T-Test (Compare to 1st Subgroup)", 
                         "T-Test (Compare Adjacent Pairs)",
                         "T-Test (All Pairwise Combinations)",
                         "Mann-Whitney U (Compare to 1st Subgroup)",
                         "One-Way ANOVA (Overall)"
                     ]).pack(fill=tk.X, padx=5, pady=2)

        lbl_stats_info = tk.Label(frame_sig, text="* p≤0.05, ** p≤0.01, *** p≤0.001", 
                                  font=("Arial", 8, "italic"), fg="#555", justify=tk.LEFT)
        lbl_stats_info.pack(anchor="w", padx=5, pady=5)

        btn_generate = tk.Button(parent, text="📊 Generate / Refresh Preview", bg="#0066cc", fg="white", 
                                 font=("Arial", 10, "bold"), command=self.generate_plot)
        btn_generate.pack(fill=tk.X, padx=10, pady=10)

        btn_export = tk.Button(parent, text="💾 Export Graph High-Res", bg="#2e7d32", fg="white", 
                               font=("Arial", 10, "bold"), command=self.export_plot)
        btn_export.pack(fill=tk.X, padx=10, pady=2)

    def generate_plot(self):
        if not hasattr(self, 'ax'):
            return
            
        self.ax.clear()
        
        df = self.parse_data_matrix()
        if df.empty:
            self.ax.axis('off')
            self.ax.text(0.5, 0.5, "Data table is empty.\nEnter data or load a file to generate plot.", 
                         ha='center', va='center', color='gray', fontsize=12)
            self.mpl_canvas.draw()
            return

        self.ax.axis('on') 
        categories = list(df["Category"].unique())
        subgroups = list(df["Subgroup"].unique())
        n_cats, n_subs = len(categories), len(subgroups)

        try:
            sizes = [int(s.strip()) for s in self.var_font_sizes.get().split(',')]
            title_sz, label_sz, tick_sz = sizes[0], sizes[1], sizes[2]
        except Exception:
            title_sz, label_sz, tick_sz = 14, 12, 10

        colors = self.get_color_cycle(n_subs)
        
        hatches_raw = [h.strip() for h in self.var_hatches.get().split(',')]
        hatches = ["" if h.lower() == "none" or not h else h for h in hatches_raw]
        if not hatches: hatches = [""]
        hatches_cycle = list(itertools.islice(itertools.cycle(hatches), n_subs))

        markers = ['o', 's', '^', 'v', 'D', 'p', '*']
        marker_cycle = list(itertools.islice(itertools.cycle(markers), n_subs))

        bar_width = self.var_bar_width.get()
        spacing = self.var_spacing.get()
        jitter = self.var_jitter.get()
        chart_mode = self.var_chart_mode.get()
        error_mode = self.var_error_type.get()
        r = np.arange(n_cats)

        bar_positions = {}
        bar_max_heights = {}
        highest_data_point = 0

        for i, sub in enumerate(subgroups):
            sub_df = df[df["Subgroup"] == sub]
            means, errors, raw_pts = [], [], []
            
            for c_idx, cat in enumerate(categories):
                cat_sub_vals = sub_df[sub_df["Category"] == cat]["Value"].values
                if len(cat_sub_vals) > 0:
                    mean_val = np.mean(cat_sub_vals)
                    means.append(mean_val)
                    raw_pts.append(cat_sub_vals)
                    
                    if error_mode == "SD (Standard Deviation)":
                        err_val = np.std(cat_sub_vals) if len(cat_sub_vals) > 1 else 0.0
                    elif error_mode == "SEM (Standard Error)":
                        err_val = np.std(cat_sub_vals) / np.sqrt(len(cat_sub_vals)) if len(cat_sub_vals) > 1 else 0.0
                    else:
                        err_val = 0.0
                    
                    errors.append(err_val)
                    highest_data_point = max(highest_data_point, mean_val + err_val, max(cat_sub_vals))
                else:
                    means.append(0.0)
                    errors.append(0.0)
                    raw_pts.append(np.array([]))

            offset = (i - (n_subs - 1) / 2) * (bar_width + spacing)
            x_pos = r + offset

            for c_idx, cat in enumerate(categories):
                bar_positions[(cat, sub)] = x_pos[c_idx]
                bar_max_heights[(cat, sub)] = means[c_idx] + errors[c_idx]

            if "Grouped Bar" in chart_mode:
                self.ax.bar(x_pos, means, color=colors[i], width=bar_width, 
                            edgecolor="black", hatch=hatches_cycle[i], linewidth=1.2, label=sub)
                
                if error_mode != "None":
                    self.ax.errorbar(x_pos, means, yerr=errors, fmt='none', ecolor='black', 
                                     capsize=4, elinewidth=1.2, capthick=1.2, zorder=4)
                
                if "Scatter" in chart_mode:
                    for j, x in enumerate(x_pos):
                        pts = raw_pts[j]
                        if len(pts) > 0:
                            jittered_x = np.random.normal(x, jitter, len(pts))
                            self.ax.scatter(jittered_x, pts, marker=marker_cycle[i], s=35, 
                                            facecolors='#333333' if colors[i].lower() != "#ffffff" else 'none', 
                                            edgecolors='black', zorder=5, alpha=0.9)
            
            # --- FIX APPLIED HERE ---
            elif "Box Plot" in chart_mode:
                # Zip raw points and their specific X positions together, then filter out empty ones
                valid_pairs = [(list(pts), x) for pts, x in zip(raw_pts, x_pos) if len(pts) > 0]
                if valid_pairs:
                    box_data = [p[0] for p in valid_pairs]
                    valid_x_pos = [p[1] for p in valid_pairs]
                    bp = self.ax.boxplot(box_data, positions=valid_x_pos, widths=bar_width, patch_artist=True, manage_ticks=False)
                    for patch in bp['boxes']: patch.set_facecolor(colors[i])
            
            # --- FIX APPLIED HERE ---
            elif "Violin Plot" in chart_mode:
                # Zip raw points and their specific X positions together, then filter out empty ones
                valid_pairs = [(pts, x) for pts, x in zip(raw_pts, x_pos) if len(pts) > 0]
                if valid_pairs:
                    violin_data = [p[0] for p in valid_pairs]
                    valid_x_pos = [p[1] for p in valid_pairs]
                    vp = self.ax.violinplot(violin_data, positions=valid_x_pos, widths=bar_width * 1.5, showmeans=True)
                    for pc in vp['bodies']: pc.set_facecolor(colors[i])

        stats_mode = self.var_stats_mode.get()
        dynamic_bracket_y = highest_data_point * 1.05
        bracket_increment = highest_data_point * 0.10
        
        if stats_mode != "None" and len(subgroups) > 1:
            for cat in categories:
                cat_df = df[df["Category"] == cat]
                bracket_level = 0
                
                if stats_mode == "One-Way ANOVA (Overall)":
                    group_data = [cat_df[cat_df["Subgroup"] == s]["Value"].values for s in subgroups]
                    valid_groups = [g for g in group_data if len(g) > 1]
                    if len(valid_groups) > 1:
                        try:
                            f_stat, p_val = stats.f_oneway(*valid_groups)
                            sig_text = f"ANOVA p={p_val:.4f}" + (" *" if p_val <= 0.05 else "")
                            x_center = np.mean([bar_positions[(cat, s)] for s in subgroups])
                            y_height = max([bar_max_heights[(cat, s)] for s in subgroups]) + (highest_data_point * 0.05)
                            self.ax.text(x_center, y_height, sig_text, ha='center', va='bottom', color='black', fontsize=label_sz)
                            dynamic_bracket_y = max(dynamic_bracket_y, y_height + bracket_increment)
                        except Exception: pass

                else:
                    pairs_to_test = []
                    if "Compare to 1st" in stats_mode:
                        control_sub = subgroups[0]
                        pairs_to_test = [(control_sub, s) for s in subgroups[1:]]
                    elif "Adjacent Pairs" in stats_mode:
                        pairs_to_test = [(subgroups[i], subgroups[i+1]) for i in range(len(subgroups)-1)]
                    elif "All Pairwise Combinations" in stats_mode:
                        pairs_to_test = list(itertools.combinations(subgroups, 2))

                    for (sub1, sub2) in pairs_to_test:
                        data1 = cat_df[cat_df["Subgroup"] == sub1]["Value"].values
                        data2 = cat_df[cat_df["Subgroup"] == sub2]["Value"].values
                        if len(data1) < 2 or len(data2) < 2: continue
                        
                        try:
                            if "Mann-Whitney" in stats_mode:
                                stat_val, p_val = stats.mannwhitneyu(data1, data2, alternative='two-sided')
                            else:
                                stat_val, p_val = stats.ttest_ind(data1, data2, equal_var=False)

                            sig_text = "***" if p_val <= 0.001 else "**" if p_val <= 0.01 else "*" if p_val <= 0.05 else ""
                            
                            if sig_text:
                                x1, x2 = bar_positions[(cat, sub1)], bar_positions[(cat, sub2)]
                                local_max = max(bar_max_heights[(cat, sub1)], bar_max_heights[(cat, sub2)])
                                
                                y_height = local_max + (highest_data_point * 0.05) + (bracket_level * bracket_increment)
                                bracket_h = highest_data_point * 0.02
                                
                                self.ax.plot([x1, x1, x2, x2], [y_height-bracket_h, y_height, y_height, y_height-bracket_h], lw=1.2, c='black')
                                self.ax.text((x1+x2)*.5, y_height, sig_text, ha='center', va='bottom', color='black', fontsize=label_sz, weight='bold')
                                
                                dynamic_bracket_y = max(dynamic_bracket_y, y_height + bracket_increment)
                                bracket_level += 1
                        except Exception:
                            pass

        self.ax.set_xticks(r)
        self.ax.set_xticklabels(categories, fontsize=tick_sz)
        self.ax.tick_params(axis='y', labelsize=tick_sz)
        
        if len(r) > 0:
            self.ax.set_xlim(-0.5, len(r) - 0.5)

        if self.var_title.get(): self.ax.set_title(self.var_title.get(), fontsize=title_sz, weight='bold', pad=12)
        if self.var_y_title.get(): self.ax.set_ylabel(self.var_y_title.get(), fontsize=label_sz, weight='bold')
        if self.var_x_title.get(): self.ax.set_xlabel(self.var_x_title.get(), fontsize=label_sz, weight='bold')

        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)
        self.ax.spines['left'].set_linewidth(1.5)
        self.ax.spines['bottom'].set_linewidth(1.5)

        if self.var_y_limits.get():
            try:
                lim_raw = [lim.strip() for lim in self.var_y_limits.get().split(',')]
                if len(lim_raw) == 2:
                    self.ax.set_ylim(float(lim_raw[0]), float(lim_raw[1]))
            except Exception: pass
        else:
            ymax = max(highest_data_point * 1.20, dynamic_bracket_y + (highest_data_point * 0.08))
            if ymax <= 0 or np.isnan(ymax): ymax = 1 
            self.ax.set_ylim(bottom=0, top=ymax)

        leg_pos = self.var_legend_pos.get()
        if leg_pos != "None":
            handles, labels = self.ax.get_legend_handles_labels()
            if handles:
                if leg_pos == "outside right":
                    self.ax.legend(handles[:n_subs], labels[:n_subs], loc="upper left", 
                                   bbox_to_anchor=(1.03, 1.0), frameon=True, fontsize=tick_sz)
                else:
                    self.ax.legend(handles[:n_subs], labels[:n_subs], loc=leg_pos, frameon=True, fontsize=tick_sz)

        self.apply_layout_transform()

    def export_plot(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("TIFF Image", "*.tiff")],
            title="Export High-Res Graph..."
        )
        if file_path:
            try:
                self.figure.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Graph successfully saved to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to save image:\n{str(e)}")