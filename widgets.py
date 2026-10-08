import tkinter as tk
import numpy as np
from PIL import Image, ImageTk
import cv2

class ColorRangeSlider(tk.Canvas):
    """Custom Tkinter Widget for a Double-Pointer Range Slider (Hue, Intensity, Area)"""
    def __init__(self, parent, width=500, height=42, slider_type="hue", abs_min=0, abs_max=179, command=None, **kwargs):
        super().__init__(parent, width=width, height=height, bg=parent.cget("bg"), highlightthickness=0, **kwargs)
        self.command = command
        self.width = width
        self.height = height
        
        self.slider_type = slider_type
        self.abs_min = abs_min
        self.abs_max = abs_max
        
        self.cur_min = abs_min
        self.cur_max = abs_max
        self.active_handle = None

        # --- SHIFTED DOWNWARDS TO PREVENT CLIPPING ---
        self.pad_x = 25              
        self.slider_y_top = 24       # Shifted down from 20 to clear space on top
        self.slider_y_bottom = 42    # Lowered proportionally

        self._create_background()
        
        self.left_mask = self.create_rectangle(self.pad_x, self.slider_y_top, self.pad_x, self.slider_y_bottom, fill="black", stipple="gray50", outline="")
        self.right_mask = self.create_rectangle(self.pad_x, self.slider_y_top, width - self.pad_x, self.slider_y_bottom, fill="black", stipple="gray50", outline="")
        
        self.min_handle = self.create_line(0, self.slider_y_top, 0, self.slider_y_bottom, fill="red", width=2)
        self.max_handle = self.create_line(0, self.slider_y_top, 0, self.slider_y_bottom, fill="blue", width=2)
        
        # Positioned with comfortable downward padding relative to the new slider top coordinate
        self.min_text = self.create_text(0, self.slider_y_top - 10, text=str(self.cur_min), fill="red", font=("Arial", 8, "bold"), anchor=tk.S)
        self.max_text = self.create_text(width, self.slider_y_top - 10, text=str(self.cur_max), fill="blue", font=("Arial", 8, "bold"), anchor=tk.S)
        
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Configure>", self.on_resize)
        
        self.update_ui()

    def on_resize(self, event):
        if event.width != self.width or event.height != self.height:
            self.width = max(1, event.width)
            self.height = max(1, event.height)
            self._create_background()
            self.update_ui()

    def _create_background(self):
        self.delete("bg_img")
        
        if self.slider_type == "hue":
            grad_img = np.zeros((1, 180, 3), dtype=np.uint8)
            for i in range(180):
                grad_img[0, i] = [i, 255, 255]
            grad_rgb = cv2.cvtColor(grad_img, cv2.COLOR_HSV2RGB)
        elif self.slider_type == "intensity":
            grad_img = np.zeros((1, 256, 3), dtype=np.uint8)
            for i in range(256):
                grad_img[0, i] = [i, i, i]
            grad_rgb = grad_img
        else:
            grad_rgb = np.full((1, 100, 3), [100, 150, 200], dtype=np.uint8)

        track_width = max(1, self.width - (2 * self.pad_x))
        track_height = max(1, self.slider_y_bottom - self.slider_y_top)
        
        grad_rgb = cv2.resize(grad_rgb, (track_width, track_height))
        self.bg_image = ImageTk.PhotoImage(Image.fromarray(grad_rgb))
        
        self.create_image(self.pad_x, self.slider_y_top, anchor='nw', image=self.bg_image, tags="bg_img")
        self.tag_lower("bg_img")

    def _get_slider_dimensions(self):
        slider_w = self.width - (2 * self.pad_x)
        span = self.abs_max - self.abs_min
        return slider_w, span if span > 0 else 1

    def update_ui(self):
        slider_w, span = self._get_slider_dimensions()
        
        x_min = self.pad_x + (((self.cur_min - self.abs_min) / span) * slider_w)
        x_max = self.pad_x + (((self.cur_max - self.abs_min) / span) * slider_w)
        
        self.coords(self.left_mask, self.pad_x, self.slider_y_top, x_min, self.slider_y_bottom)
        self.coords(self.right_mask, x_max, self.slider_y_top, self.width - self.pad_x, self.slider_y_bottom)
        
        self.coords(self.min_handle, x_min, self.slider_y_top - 2, x_min, self.slider_y_bottom + 2)
        self.coords(self.max_handle, x_max, self.slider_y_top - 2, x_max, self.slider_y_bottom + 2)
        
        offset_min_x = x_min
        offset_max_x = x_max
        if abs(x_max - x_min) < 30:
            offset_min_x = x_min - 12
            offset_max_x = x_max + 12
            
        self.coords(self.min_text, offset_min_x, self.slider_y_top - 4)
        self.itemconfig(self.min_text, text=str(self.cur_min))
        
        self.coords(self.max_text, offset_max_x, self.slider_y_top - 4)
        self.itemconfig(self.max_text, text=str(self.cur_max))

    def on_click(self, event):
        slider_w, span = self._get_slider_dimensions()
        x_min = self.pad_x + (((self.cur_min - self.abs_min) / span) * slider_w)
        x_max = self.pad_x + (((self.cur_max - self.abs_min) / span) * slider_w)
        
        if abs(event.x - x_min) < abs(event.x - x_max):
            self.active_handle = 'min'
        else:
            self.active_handle = 'max'
        self.on_drag(event)

    def on_drag(self, event):
        slider_w, span = self._get_slider_dimensions()
        clamped_x = max(self.pad_x, min(self.width - self.pad_x, event.x))
        new_val = int(self.abs_min + ((clamped_x - self.pad_x) / slider_w) * span)
        
        if self.active_handle == 'min':
            self.cur_min = min(new_val, self.cur_max - 1)
        elif self.active_handle == 'max':
            self.cur_max = max(new_val, self.cur_min + 1)
            
        self.update_ui()
        if self.command:
            self.command() 

    def on_release(self, event):
        self.active_handle = None
        if self.command:
            self.command()

    def set_values(self, min_val, max_val):
        self.cur_min = max(self.abs_min, min(self.abs_max, min_val))
        self.cur_max = max(self.abs_min, min(self.abs_max, max_val))
        self.update_ui()

    def get_values(self):
        return self.cur_min, self.cur_max


# ---> NEW WIDGET: Single Slider for Circularity/Splitting <---
class SingleSlider(tk.Canvas):
    """Custom Tkinter Widget for a Single-Pointer Slider"""
    def __init__(self, parent, width=500, height=42, abs_min=0, abs_max=100, default_value=None, command=None, **kwargs):
        super().__init__(parent, width=width, height=height, bg=parent.cget("bg"), highlightthickness=0, **kwargs)
        self.command = command
        self.width = width
        self.height = height
        
        self.abs_min = abs_min
        self.abs_max = abs_max
        
        # --- SHIFTED DOWNWARDS TO PREVENT CLIPPING ---
        self.pad_x = 25
        self.slider_y_top = 24       # Shifted down from 20 to make upper room
        self.slider_y_bottom = 42    # Lowered proportionally
        
        if default_value is not None:
            self.cur_val = max(self.abs_min, min(self.abs_max, default_value))
        else:
            self.cur_val = abs_min
        
        self._create_background()
        
        self.right_mask = self.create_rectangle(self.pad_x, self.slider_y_top, width - self.pad_x, self.slider_y_bottom, fill="black", stipple="gray50", outline="")
        
        self.center_line = None
        if self.abs_min < 0 < self.abs_max:
            self.center_line = self.create_line(0, self.slider_y_top, 0, self.slider_y_bottom, fill="white", dash=(2, 2))

        self.handle = self.create_line(0, self.slider_y_top, 0, self.slider_y_bottom, fill="black", width=2)
        self.val_text = self.create_text(0, self.slider_y_top - 10, text=str(self.cur_val), fill="black", font=("Arial", 8, "bold"), anchor=tk.S)
        
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Configure>", self.on_resize)
        
        self.update_ui()

    def on_resize(self, event):
        if event.width != self.width or event.height != self.height:
            self.width = max(1, event.width)
            self.height = max(1, event.height)
            self._create_background()
            self.update_ui()

    def _create_background(self):
        self.delete("bg_img")
        grad_rgb = np.full((1, 100, 3), [160, 120, 200], dtype=np.uint8)
        
        track_width = max(1, self.width - (2 * self.pad_x))
        track_height = max(1, self.slider_y_bottom - self.slider_y_top)
        
        grad_rgb = cv2.resize(grad_rgb, (track_width, track_height))
        self.bg_image = ImageTk.PhotoImage(Image.fromarray(grad_rgb))
        
        self.create_image(self.pad_x, self.slider_y_top, anchor='nw', image=self.bg_image, tags="bg_img")
        self.tag_lower("bg_img")

    def _get_slider_dimensions(self):
        slider_w = self.width - (2 * self.pad_x)
        span = self.abs_max - self.abs_min
        return slider_w, span if span > 0 else 1

    def _val_to_x(self, val):
        slider_w, span = self._get_slider_dimensions()
        return self.pad_x + (((val - self.abs_min) / span) * slider_w)

    def _x_to_val(self, x):
        slider_w, span = self._get_slider_dimensions()
        clamped_x = max(self.pad_x, min(self.width - self.pad_x, x))
        return int(self.abs_min + ((clamped_x - self.pad_x) / slider_w) * span)

    def update_ui(self):
        x = self._val_to_x(self.cur_val)
        self.coords(self.right_mask, x, self.slider_y_top, self.width - self.pad_x, self.slider_y_bottom)
        
        if self.center_line:
            zero_x = self._val_to_x(0)
            self.coords(self.center_line, zero_x, self.slider_y_top, zero_x, self.slider_y_bottom)
        
        self.coords(self.handle, x, self.slider_y_top - 2, x, self.slider_y_bottom + 2)
        self.coords(self.val_text, x, self.slider_y_top - 4)
        self.itemconfig(self.val_text, text=str(self.cur_val))

    def on_click(self, event):
        self.on_drag(event)

    def on_drag(self, event):
        self.cur_val = self._x_to_val(event.x)
        self.update_ui()
        if self.command:
            self.command() 

    def on_release(self, event):
        if self.command:
            self.command()

    def set_values(self, val):
        self.cur_val = max(self.abs_min, min(self.abs_max, val))
        self.update_ui()

    def get_values(self):
        return self.cur_val
