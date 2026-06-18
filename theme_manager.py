import tkinter as tk
from tkinter import ttk
from config import Config

class ThemeManager:
    """Manages switching between Light and Dark mode dynamically."""
    
    # Save the original light colors
    LIGHT_COLORS = Config.COLORS.copy()
    
    # Define dark mode palette
    DARK_COLORS = {
        "primary": "#E2B785",          # Light Gold/Brown
        "secondary": "#8C6A50",        # Medium Brown
        "accent": "#A38A75",           # Lighter Brown
        "success": "#7C9C49",          # Brighter Olive Green
        "warning": "#F58A33",          # Brighter Orange
        "danger": "#D64040",           # Brighter Red
        "info": "#609CCC",             # Brighter Blue
        "light": "#121212",            # Main Background (Very Dark)
        "dark": "#EBE7E0",             # Text (Light Beige)
        "white": "#1E1E1E",            # Card Background (Dark Grey)
        "black": "#FFFFFF",            # Pure White
        "gray": "#A0A0A0",             # Light Gray
        "border": "#333333",           # Dark Border
        "hover": "#C9A07E",            # Hover state
        "gradient_start": "#1E1E1E",
        "gradient_end": "#121212",
        "alternate_row": "#252525",    # Dark Grey for alternate rows
        "weekend": "#2A2A2A",          
        "sunday": "#333333",           
        "email_highlight": "#1E1E1E",  
        "scheduled": "#252525",        
        "card_bg": "#1E1E1E",          
        "selected_row": "#444444",     
        "selected_border": "#E2B785",  
    }
    
    _is_dark_mode = False
    
    @classmethod
    def toggle_theme(cls, root_window):
        cls._is_dark_mode = not cls._is_dark_mode
        if cls._is_dark_mode:
            Config.COLORS.update(cls.DARK_COLORS)
        else:
            Config.COLORS.update(cls.LIGHT_COLORS)
            
        # Reconfigure ttk styles
        cls._configure_ttk_styles()
        
        # Recursively update all existing tk widgets
        cls._update_widgets(root_window)
        
    @classmethod
    def is_dark_mode(cls):
        return cls._is_dark_mode
        
    @classmethod
    def _configure_ttk_styles(cls):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Notebook
        style.configure("TNotebook", background=Config.COLORS["light"])
        style.configure("TNotebook.Tab", background=Config.COLORS["white"], foreground=Config.COLORS["dark"])
        
        # Treeview
        style.configure("Treeview", 
                        background=Config.COLORS["white"], 
                        foreground=Config.COLORS["dark"], 
                        fieldbackground=Config.COLORS["white"])
        style.map("Treeview", background=[('selected', Config.COLORS["selected_row"])], 
                  foreground=[('selected', Config.COLORS["black"])])
                  
        # Scrollbars
        style.configure("Vertical.TScrollbar", background=Config.COLORS["border"], troughcolor=Config.COLORS["light"])
        style.configure("Horizontal.TScrollbar", background=Config.COLORS["border"], troughcolor=Config.COLORS["light"])

    @classmethod
    def _update_widgets(cls, widget):
        """Recursively update colors of standard tk widgets."""
        try:
            wtype = widget.winfo_class()
            
            # Map widget classes to color keys they typically use
            # Note: This is a best-effort update. Custom drawn canvas buttons might need re-drawing.
            if wtype in ('Frame', 'Toplevel', 'Tk'):
                # Heuristic: If it was light, make it light (which is now dark)
                # This requires knowing what it was. We just set to Config.COLORS["light"] or "white"
                # based on its current color.
                current_bg = widget.cget("bg")
                if current_bg in (cls.LIGHT_COLORS["light"], cls.DARK_COLORS["light"]):
                    widget.configure(bg=Config.COLORS["light"])
                elif current_bg in (cls.LIGHT_COLORS["white"], cls.DARK_COLORS["white"]):
                    widget.configure(bg=Config.COLORS["white"])
                elif current_bg in (cls.LIGHT_COLORS["primary"], cls.DARK_COLORS["primary"]):
                    widget.configure(bg=Config.COLORS["primary"])
                elif current_bg in (cls.LIGHT_COLORS["border"], cls.DARK_COLORS["border"]):
                    widget.configure(bg=Config.COLORS["border"])
                    
            elif wtype == 'Label':
                current_bg = widget.cget("bg")
                current_fg = widget.cget("fg")
                
                # Backgrounds
                if current_bg in (cls.LIGHT_COLORS["light"], cls.DARK_COLORS["light"]):
                    widget.configure(bg=Config.COLORS["light"])
                elif current_bg in (cls.LIGHT_COLORS["white"], cls.DARK_COLORS["white"]):
                    widget.configure(bg=Config.COLORS["white"])
                elif current_bg in (cls.LIGHT_COLORS["primary"], cls.DARK_COLORS["primary"]):
                    widget.configure(bg=Config.COLORS["primary"])
                    
                # Foregrounds
                if current_fg in (cls.LIGHT_COLORS["dark"], cls.DARK_COLORS["dark"], "#333333"):
                    widget.configure(fg=Config.COLORS["dark"])
                elif current_fg in (cls.LIGHT_COLORS["primary"], cls.DARK_COLORS["primary"]):
                    widget.configure(fg=Config.COLORS["primary"])
                elif current_fg in (cls.LIGHT_COLORS["gray"], cls.DARK_COLORS["gray"]):
                    widget.configure(fg=Config.COLORS["gray"])
                    
            elif wtype == 'Button':
                current_bg = widget.cget("bg")
                if current_bg in (cls.LIGHT_COLORS["primary"], cls.DARK_COLORS["primary"]):
                    widget.configure(bg=Config.COLORS["primary"])
                elif current_bg in (cls.LIGHT_COLORS["success"], cls.DARK_COLORS["success"]):
                    widget.configure(bg=Config.COLORS["success"])
                elif current_bg in (cls.LIGHT_COLORS["warning"], cls.DARK_COLORS["warning"]):
                    widget.configure(bg=Config.COLORS["warning"])
                elif current_bg in (cls.LIGHT_COLORS["white"], cls.DARK_COLORS["white"]):
                    widget.configure(bg=Config.COLORS["white"], fg=Config.COLORS["primary"])
                    
            elif wtype == 'Canvas':
                current_bg = widget.cget("bg")
                if current_bg in (cls.LIGHT_COLORS["white"], cls.DARK_COLORS["white"]):
                    widget.configure(bg=Config.COLORS["white"])
                elif current_bg in (cls.LIGHT_COLORS["light"], cls.DARK_COLORS["light"]):
                    widget.configure(bg=Config.COLORS["light"])
                    
        except Exception:
            pass
            
        # Recurse children
        for child in widget.winfo_children():
            cls._update_widgets(child)
