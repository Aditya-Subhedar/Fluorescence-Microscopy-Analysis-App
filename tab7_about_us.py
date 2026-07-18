import tkinter as tk
import webbrowser

class AboutUsTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#f8f9fa")
        
        # Center container card for a clean UI
        container = tk.Frame(self, bg="#ffffff", bd=1, relief=tk.SOLID, padx=40, pady=40)
        container.place(relx=0.5, rely=0.5, anchor="center")
        
        # Header
        lbl_header = tk.Label(container, text="About CytoQuant", font=("Arial", 22, "bold"), bg="#ffffff", fg="#2c3e50")
        lbl_header.pack(pady=(0, 20))
        
        # Main Info
        info_text = (
            "This software is the creation of the Neuroepigenitics lab (Dr. Amul J Sakharkar) Department of Biotechnology at\n"
            "Savitribai Phule Pune University (SPPU).\n\n"
            "Under the Guidance of:\n"
            "• Dr. Amul J. Sakharkar (SPPU)\n"
            "• Prof. Aurnab Ghose (IISER Pune)\n"
            "• Prof. Nishikant Subhedar (IISER Pune)\n\n"
            "Developed By:\n"
            "Aditya Subhedar (Student of MIT-WPU, Intern at SPPU)\n\n"
            "In Collaboration with PhD and MSc Students:\n"
            "Namrata Pawar (SPPU), Vaishnavi Borade (SPPU), Maithili Borkute (SPPU), Anushka Walupante (SPPU), Unmani Harsulkar (SPPU)"
        )
        
        lbl_info = tk.Label(container, text=info_text, font=("Arial", 11), bg="#ffffff", fg="#495057", justify=tk.CENTER)
        lbl_info.pack(pady=10)
        
        # GitHub Link
        link_frame = tk.Frame(container, bg="#ffffff")
        link_frame.pack(pady=(20, 0))
        
        git_icon = tk.Label(link_frame, text="💻", font=("Arial", 14), bg="#ffffff")
        git_icon.pack(side=tk.LEFT, padx=(0, 5))
        
        lbl_git = tk.Label(link_frame, text="Contribute on GitHub", font=("Arial", 12, "underline"), bg="#ffffff", fg="#0066cc", cursor="hand2")
        lbl_git.pack(side=tk.LEFT)
        
        # Open URL function
        def open_github(event):
            webbrowser.open_new("https://github.com/Aditya-Subhedar/Fluorescence-Microscopy-Analysis-App")
            
        lbl_git.bind("<Button-1>", open_github)