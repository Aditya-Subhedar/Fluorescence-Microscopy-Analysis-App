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
        
        # Rigid layout state variables for unified canvas manipulation
        self.graph_zoom = 1.0
        self.graph_pan_x = 0.0
        self.graph_pan_y = 0.0
        self.is_panning = False
        
        self.setup_ui()

    def setup_ui(self):
        self.pack(fill=tk.BOTH, expand=True)

        self.paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ==========================================
        # LEFT PANEL: SCROLLABLE CONTROLS
        # ==========================================
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
        
        # Structurally isolate scroll interactions only to the left UI panel tree
        self.bind_left_scroll_recursive(self.left_container)

        # ==========================================
        # RIGHT PANEL: MATPLOTLIB VIEWPORT
        # ==========================================
        self.right_panel = tk.Frame(self.paned_window, bg="white")
        self.paned_window.add(self.right_panel, weight=1)

        self.figure = Figure(figsize=(5, 4), dpi=100)
        self.figure.patch.set_facecolor('#ffffff')
        self.ax = self.figure.add_subplot(111)
        
        self.mpl_canvas = FigureCanvasTkAgg(self.figure, master=self.right_panel)
        self.canvas = self.mpl_canvas.get_tk_widget()
        
        # Connect listeners for rigid canvas transformation interactions
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
        """ Attaches mousewheel tracking only to left-panel elements """
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

    def on_zoom_layout(self, event):
        """ Scales the overall dimensions of the entire graph image uniformly """
        if event.step > 0:
            self.graph_zoom *= 1.15  # Zoom In
        else:
            self.graph_zoom /= 1.15  # Zoom Out

        self.graph_zoom = max(0.15, min(self.graph_zoom, 6.0))
        self.apply_layout_transform()

    def on_pan_start(self, event):
        """ Starts tracking pixel coordinates on left-click / trackpad press """
        if event.button == 1:
            self.is_panning = True
            self.start_pixel_x = event.x
            self.start_pixel_y = event.y
            self.orig_pan_x = self.graph_pan_x
            self.orig_pan_y = self.graph_pan_y

    def on_pan_move(self, event):
        """ Shifts the entire rigid block relative to cursor movement """
        if self.is_panning and event.x is not None and event.y is not None:
            dx_pixels = event.x - self.start_pixel_x
            dy_pixels = event.y - self.start_pixel_y
            
            fig_w = self.figure.bbox.width
            fig_h = self.figure.bbox.height
            
            # Translate raw pixel delta into localized viewport percentages
            self.graph_pan_x = self.orig_pan_x + (dx_pixels / fig_w)
            self.graph_pan_y = self.orig_pan_y + (dy_pixels / fig_h)
            
            self.apply_layout_transform()

    def on_pan_end(self, event):
        self.is_panning = False

    def apply_layout_transform(self):
        """ Shifts and resizes the entire subplots bounding area as a single unit """
        base_left, base_right = 0.20, 0.92
        base_bottom, base_top = 0.20, 0.90
        
        base_w = base_right - base_left
        base_h = base_top - base_bottom
        
        # Calculate true geometric center point with tracking offsets
        cx = 0.56 + self.graph_pan_x
        cy = 0.55 + self.graph_pan_y
        
        w = base_w * self.graph_zoom
        h = base_h * self.graph_zoom
        
        left = cx - w / 2
        right = cx + w / 2
        bottom = cy - h / 2
        top = cy + h / 2
        
        # Structural guardrails to prevent internal Matplotlib rendering errors
        if left >= right: right = left + 0.01
        if bottom >= top: top = bottom + 0.01
        
        try:
            self.figure.subplots_adjust(left=left, right=right, bottom=bottom, top=top)
            self.mpl_canvas.draw_idle()
        except Exception:
            pass

    def _build_control_widgets(self, parent):
        # --- SECTION 1: DATA TABLE ---
        frame_data = ttk.LabelFrame(parent, text="Step 1: Raw Replicate Data Table")
        frame_data.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(frame_data, text="Format: Category | Subgroup | Replicates (comma-sep)", 
                 font=("Arial", 8, "italic"), fg="#555").pack(anchor="w", padx=5, pady=(2,5))
        
        self.txt_data = tk.Text(frame_data, height=8, width=42, font=("Consolas", 9))
        self.txt_data.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # Data box left completely empty intentionally

        # --- SECTION 2: GRAPH LABELS ---
        frame_labels = ttk.LabelFrame(parent, text="Step 2: Axis & Titles")
        frame_labels.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(frame_labels, text="Graph Title:").pack(anchor="w", padx=5)
        self.var_title = tk.StringVar(value="")
        tk.Entry(frame_labels, textvariable=self.var_title).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_labels, text="Y-Axis Title:").pack(anchor="w", padx=5)
        self.var_y_title = tk.StringVar(value="")
        tk.Entry(frame_labels, textvariable=self.var_y_title).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_labels, text="X-Axis Title:").pack(anchor="w", padx=5)
        self.var_x_title = tk.StringVar(value="")
        tk.Entry(frame_labels, textvariable=self.var_x_title).pack(fill=tk.X, padx=5, pady=2)

        # --- SECTION 3: GRAPH TYPE & STYLE ---
        frame_type = ttk.LabelFrame(parent, text="Step 3: Graph Type & Style")
        frame_type.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_type, text="Chart Mode:").pack(anchor="w", padx=5)
        self.var_chart_mode = tk.StringVar(value="Grouped Bar (Mean + Error Only)")
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

        # --- SECTION 4: COLOR CONFIGURATOR ---
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

        # --- SECTION 5: GEOMETRY & SPACING ---
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

        # --- SECTION 6: FONTS & LEGEND ---
        frame_fonts = ttk.LabelFrame(parent, text="Step 6: Typography & Layout")
        frame_fonts.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_fonts, text="Font Family:").pack(anchor="w", padx=5)
        self.var_font = tk.StringVar(value="Arial")
        ttk.Combobox(frame_fonts, textvariable=self.var_font, values=["Arial", "Times New Roman", "Helvetica", "Courier New"]).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_fonts, text="Font Sizes (Title, Axis, Ticks):").pack(anchor="w", padx=5)
        self.var_font_sizes = tk.StringVar(value="")
        tk.Entry(frame_fonts, textvariable=self.var_font_sizes).pack(fill=tk.X, padx=5, pady=2)

        tk.Label(frame_fonts, text="Legend Position:").pack(anchor="w", padx=5)
        self.var_legend_pos = tk.StringVar(value="upper right")
        ttk.Combobox(frame_fonts, textvariable=self.var_legend_pos, state="readonly",
                     values=["upper right", "upper left", "lower right", "lower left", "best", "None"]).pack(fill=tk.X, padx=5, pady=2)

        # --- SECTION 7: AUTOMATED STATISTICS ---
        frame_sig = ttk.LabelFrame(parent, text="Step 7: Automated Statistics")
        frame_sig.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_sig, text="Significance Testing (α = 0.05):").pack(anchor="w", padx=5)
        self.var_stats_mode = tk.StringVar(value="None")
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

        # --- ACTIONS ---
        btn_generate = tk.Button(parent, text="📊 Generate / Refresh Preview", bg="#0066cc", fg="white", 
                                 font=("Arial", 10, "bold"), command=self.generate_plot)
        btn_generate.pack(fill=tk.X, padx=10, pady=10)

        btn_export = tk.Button(parent, text="💾 Export Graph High-Res", bg="#2e7d32", fg="white", 
                               font=("Arial", 10, "bold"), command=self.export_plot)
        btn_export.pack(fill=tk.X, padx=10, pady=2)

    def _on_canvas_configure(self, event):
        self.canvas_left.itemconfig(self.canvas_window, width=event.width)

    def parse_data_matrix(self):
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
                
        if not palette: palette = ["#4A90E2", "#F39C12", "#2ecc71", "#e74c3c", "#9b59b6"]
        return list(itertools.islice(itertools.cycle(palette), subgroups_count))

    def generate_plot(self):
        self.ax.clear()
        
        df = self.parse_data_matrix()
        if df.empty:
            self.ax.axis('off')
            # Handles the empty state gracefully, showing instructions in the viewport
            self.ax.text(0.5, 0.5, "Data table is empty.\nEnter data to generate plot.", 
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
            elif "Box Plot" in chart_mode:
                box_data = [list(pts) for pts in raw_pts if len(pts) > 0]
                if box_data:
                    bp = self.ax.boxplot(box_data, positions=x_pos, widths=bar_width, patch_artist=True, manage_ticks=False)
                    for patch in bp['boxes']: patch.set_facecolor(colors[i])
            elif "Violin Plot" in chart_mode:
                violin_data = [pts for pts in raw_pts if len(pts) > 0]
                if violin_data:
                    vp = self.ax.violinplot(violin_data, positions=x_pos, widths=bar_width * 1.5, showmeans=True)
                    for pc in vp['bodies']: pc.set_facecolor(colors[i])

        # --- AUTOMATED STATISTICAL SIGNIFICANCE ---
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

        # --- FORMATTING ---
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
            ymax = max(highest_data_point * 1.15, dynamic_bracket_y * 1.05)
            if ymax <= 0 or np.isnan(ymax): ymax = 1 
            self.ax.set_ylim(bottom=0, top=ymax)

        leg_pos = self.var_legend_pos.get()
        if leg_pos != "None":
            handles, labels = self.ax.get_legend_handles_labels()
            if handles: self.ax.legend(handles[:n_subs], labels[:n_subs], loc=leg_pos, frameon=True, fontsize=tick_sz)

        # Apply structural changes and refresh rendering layout seamlessly
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