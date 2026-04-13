import tkinter as tk
import numpy as np
from PIL import Image, ImageTk
import cv2

class ColorRangeSlider(tk.Canvas):
    """Custom Tkinter Widget for a Double-Pointer Range Slider (Hue, Intensity, Area)"""
    def __init__(self, parent, width=500, height=40, slider_type="hue", abs_min=0, abs_max=179, command=None, **kwargs):
        super().__init__(parent, width=width, height=height, bg='gray', highlightthickness=0, **kwargs)
        self.command = command
        self.width = width
        self.height = height
        
        self.slider_type = slider_type
        self.abs_min = abs_min
        self.abs_max = abs_max
        
        self.cur_min = abs_min
        self.cur_max = abs_max
        self.active_handle = None

        self._create_background()
        
        self.left_mask = self.create_rectangle(0, 0, 0, height, fill="black", stipple="gray50", outline="")
        self.right_mask = self.create_rectangle(0, 0, width, height, fill="black", stipple="gray50", outline="")
        
        self.min_handle = self.create_polygon(0,0, 0,0, 0,0, fill="white", outline="black", width=1)
        self.max_handle = self.create_polygon(0,0, 0,0, 0,0, fill="white", outline="black", width=1)
        
        # Add text labels so we can see the values!
        self.min_text = self.create_text(0, height/2, text=str(self.cur_min), fill="white", font=("Arial", 10, "bold"))
        self.max_text = self.create_text(width, height/2, text=str(self.cur_max), fill="white", font=("Arial", 10, "bold"))
        
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Configure>", self.on_resize) # ---> FIX: Listens for window resize
        
        self.update_ui()

    def on_resize(self, event):
        # ---> FIX: Recalculates slider width and redraws to eliminate dead space
        if event.width != self.width or event.height != self.height:
            self.width = max(1, event.width)
            self.height = max(1, event.height)
            self._create_background()
            self.update_ui()

    def _create_background(self):
        self.delete("bg_img") # Clear old background when resizing
        
        if self.slider_type == "hue":
            # Rainbow Gradient
            grad_img = np.zeros((1, 180, 3), dtype=np.uint8)
            for i in range(180):
                grad_img[0, i] = [i, 255, 255]
            grad_rgb = cv2.cvtColor(grad_img, cv2.COLOR_HSV2RGB)
            
        elif self.slider_type == "intensity":
            # Grayscale Gradient
            grad_img = np.zeros((1, 256, 3), dtype=np.uint8)
            for i in range(256):
                grad_img[0, i] = [i, i, i]
            grad_rgb = grad_img
            
        else:
            # Area (Solid soft blue)
            grad_rgb = np.full((1, 100, 3), [100, 150, 200], dtype=np.uint8)

        # Ensure width/height are valid for cv2.resize
        safe_width = max(1, self.width)
        safe_height = max(1, self.height)
        
        grad_rgb = cv2.resize(grad_rgb, (safe_width, safe_height))
        self.bg_image = ImageTk.PhotoImage(Image.fromarray(grad_rgb))
        # Add tags="bg_img" so we can delete and recreate it safely on resize
        self.create_image(0, 0, anchor='nw', image=self.bg_image, tags="bg_img")
        self.tag_lower("bg_img") # Keep it behind the masks/pointers

    def _val_to_x(self, val):
        # Convert a value to a pixel X coordinate
        span = self.abs_max - self.abs_min
        if span == 0: return 0
        return int(((val - self.abs_min) / span) * self.width)

    def _x_to_val(self, x):
        # Convert a pixel X coordinate back to a value
        x = max(0, min(self.width, x))
        span = self.abs_max - self.abs_min
        return int(self.abs_min + (x / self.width) * span)

    def update_ui(self):
        x_min = self._val_to_x(self.cur_min)
        x_max = self._val_to_x(self.cur_max)
        
        self.coords(self.left_mask, 0, 0, x_min, self.height)
        self.coords(self.right_mask, x_max, 0, self.width, self.height)
        
        h_size = 8
        self.coords(self.min_handle, x_min-h_size, 0, x_min+h_size, 0, x_min, h_size+4)
        self.coords(self.max_handle, x_max-h_size, self.height, x_max+h_size, self.height, x_max, self.height-h_size-4)
        
        # Move text and update numbers
        self.coords(self.min_text, max(15, x_min - 15), self.height/2)
        self.itemconfig(self.min_text, text=str(self.cur_min))
        
        self.coords(self.max_text, min(self.width - 15, x_max + 15), self.height/2)
        self.itemconfig(self.max_text, text=str(self.cur_max))

    def on_click(self, event):
        x_min = self._val_to_x(self.cur_min)
        x_max = self._val_to_x(self.cur_max)
        if abs(event.x - x_min) < abs(event.x - x_max):
            self.active_handle = 'min'
        else:
            self.active_handle = 'max'
        self.on_drag(event)

    def on_drag(self, event):
        new_val = self._x_to_val(event.x)
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
    def __init__(self, parent, width=500, height=40, abs_min=0, abs_max=100, command=None, **kwargs):
        super().__init__(parent, width=width, height=height, bg='gray', highlightthickness=0, **kwargs)
        self.command = command
        self.width = width
        self.height = height
        
        self.abs_min = abs_min
        self.abs_max = abs_max
        self.cur_val = abs_min
        
        self._create_background()
        
        self.right_mask = self.create_rectangle(0, 0, width, height, fill="black", stipple="gray50", outline="")
        self.handle = self.create_polygon(0,0, 0,0, 0,0, fill="white", outline="black", width=1)
        self.val_text = self.create_text(0, height/2, text=str(self.cur_val), fill="white", font=("Arial", 10, "bold"))
        
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
        
        # Solid soft purple
        grad_rgb = np.full((1, 100, 3), [160, 120, 200], dtype=np.uint8)
        
        safe_width = max(1, self.width)
        safe_height = max(1, self.height)
        
        grad_rgb = cv2.resize(grad_rgb, (safe_width, safe_height))
        self.bg_image = ImageTk.PhotoImage(Image.fromarray(grad_rgb))
        self.create_image(0, 0, anchor='nw', image=self.bg_image, tags="bg_img")
        self.tag_lower("bg_img")

    def _val_to_x(self, val):
        span = self.abs_max - self.abs_min
        if span == 0: return 0
        return int(((val - self.abs_min) / span) * self.width)

    def _x_to_val(self, x):
        x = max(0, min(self.width, x))
        span = self.abs_max - self.abs_min
        return int(self.abs_min + (x / self.width) * span)

    def update_ui(self):
        x = self._val_to_x(self.cur_val)
        
        # Mask everything to the right of the pointer
        self.coords(self.right_mask, x, 0, self.width, self.height)
        
        h_size = 8
        self.coords(self.handle, x-h_size, 0, x+h_size, 0, x, h_size+4)
        
        self.coords(self.val_text, max(15, x - 15), self.height/2)
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