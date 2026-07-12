import customtkinter as ctk
import json
import os
import threading
from datetime import datetime
from pynput import mouse
import keyboard
import queue
import re
import sys
from PIL import Image

CONFIG_FILE = "config.json"

def get_base_path():
    return os.path.dirname(os.path.abspath(__file__))

ICON_DIR = os.path.join(get_base_path(), "icones")
ARMAS_DIR = os.path.join(get_base_path(), "armas")
LOGO_DIR = os.path.join(get_base_path(), "logos")

def crop_transparent(img):
    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
        alpha = img.convert('RGBA').split()[-1]
        bbox = alpha.getbbox()
        if bbox:
            return img.crop(bbox)
    return img

def get_proportional_size(img, max_dim=60):
    w, h = img.size
    if w > h:
        new_w = max_dim
        new_h = int(h * (max_dim / w))
    else:
        new_h = max_dim
        new_w = int(w * (max_dim / h))
    if new_w == 0: new_w = 1
    if new_h == 0: new_h = 1
    return (new_w, new_h)

def load_grip_image(grip_name, max_dim=60):
    mapping = {
        "Meio Punho": "meio_punho.png",
        "Punho Horizontal": "punho_angular.png", 
        "Punho de Polegar": "punho_polegar.png",
        "Punho Leve": "punho_leve.png",
        "Punho Vertical": "punho_vertical.png"
    }
    filename = mapping.get(grip_name)
    if filename:
        path = os.path.join(ICON_DIR, filename)
        if os.path.exists(path):
            try:
                img = Image.open(path)
                img = crop_transparent(img)
                size = get_proportional_size(img, max_dim)
                return ctk.CTkImage(light_image=img, dark_image=img, size=size)
            except Exception:
                pass
    return None

def load_weapon_image(weapon_name, max_dim=100):
    name_clean = weapon_name.lower().strip()
    aliases = {
        "m4": "m416",
        "scarl": "scar-l",
        "berryl": "beryl",
        "scarl-l": "scar-l"
    }
    
    names_to_try = [name_clean]
    if name_clean in aliases:
        names_to_try.append(aliases[name_clean])
        
    for name in names_to_try:
        for ext in ['.webp', '.png', '.jpg', '.jpeg']:
            path = os.path.join(ARMAS_DIR, name + ext)
            if os.path.exists(path):
                try:
                    img = Image.open(path)
                    img = crop_transparent(img)
                    size = get_proportional_size(img, max_dim)
                    return ctk.CTkImage(light_image=img, dark_image=img, size=size)
                except Exception:
                    pass
    return None

def load_logo_image(max_dim=120):
    try:
        path = os.path.join(LOGO_DIR, "PUBG-Logo-PNG-Transparent-Image.png")
        if not os.path.exists(path):
            path = os.path.join(LOGO_DIR, "PlayerUnknowns-Battlegrounds-PUBG-Free-PNG-Image.png")
            
        if os.path.exists(path):
            img = Image.open(path)
            img = crop_transparent(img)
            size = get_proportional_size(img, max_dim)
            return ctk.CTkImage(light_image=img, dark_image=img, size=size)
    except Exception:
        pass
    return None


def tts_worker(q, get_vol_func, get_enabled_func, get_device_func):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.Rate = 5
        outputs = speaker.GetAudioOutputs()
        audio_devices = {outputs.Item(i).GetDescription(): outputs.Item(i) for i in range(outputs.Count)}
    except Exception as e:
        print("Erro ao iniciar voz:", e)
        return

    current_device_name = None

    while True:
        text = q.get()
        if text is None:
            break
        
        if get_enabled_func():
            vol = get_vol_func()
            target_device = get_device_func()
            
            if target_device and target_device != current_device_name and target_device in audio_devices:
                try:
                    speaker.AudioOutput = audio_devices[target_device]
                    current_device_name = target_device
                except Exception as e:
                    print("Erro ao trocar dispositivo:", e)

            if vol > 0:
                try:
                    speaker.Volume = int(vol * 100)
                    speaker.Speak(text, 3)
                except Exception as e:
                    print("TTS Erro:", e)


WEAPONS_DATA = [
    {"category": "MAIS USADAS", "color": "#F2A900", "items": [
        {"name": "BERRYL", "normal": "1.10", "boca1": "0.95", "boca2": "1.30", "comp1": "0.90", "comp2": "1.20", "grip": "Meio Punho"},
        {"name": "AUG",    "normal": "1.25", "boca1": "0.90", "boca2": "1.15", "comp1": "0.80", "comp2": "0.95", "grip": "Punho Horizontal"},
        {"name": "AKM",    "normal": "0.90", "boca1": "0.80", "boca2": "1.05", "comp1": "0.75", "comp2": "0.95", "grip": "Não usa"},
        {"name": "ACE",    "normal": "1.15", "boca1": "0.80", "boca2": "0.95", "comp1": "0.70", "comp2": "0.85", "grip": "Meio Punho"},
        {"name": "K2",     "normal": "0.85", "boca1": "0.70", "boca2": "0.95", "comp1": "0.65", "comp2": "0.90", "grip": "Punho Horizontal"},
        {"name": "M4",     "normal": "0.85", "boca1": "0.55", "boca2": "0.70", "comp1": "0.55", "comp2": "0.70", "grip": "Meio Punho"},
    ]},
    {"category": "OUTRAS", "color": "#42A5F5", "items": [
        {"name": "SCARL",  "normal": "0.90", "boca1": "0.65", "boca2": "0.85", "comp1": "0.65", "comp2": "0.85", "grip": "Punho Horizontal"},
        {"name": "QBZ",    "normal": "1.00", "boca1": "0.65", "boca2": "0.90", "comp1": "0.60", "comp2": "0.85", "grip": "Punho Horizontal"},
        {"name": "G36C",   "normal": "1.00", "boca1": "0.70", "boca2": "0.90", "comp1": "0.65", "comp2": "0.90", "grip": "Punho Horizontal"},
        {"name": "VECTOR", "normal": "1.25", "boca1": "0.95", "boca2": "1.35", "comp1": "0.75", "comp2": "1.10", "grip": "Punho de Polegar"},
        {"name": "UMP",    "normal": "0.75", "boca1": "0.55", "boca2": "0.80", "comp1": "0.50", "comp2": "0.65", "grip": "Punho de Polegar"},
        {"name": "JS9",    "normal": "0.90", "boca1": "0.70", "boca2": "1.00", "comp1": "0.50", "comp2": "0.75", "grip": "Punho de Polegar"},
        {"name": "MP9",    "normal": "0.50", "boca1": "DESLIGADO", "boca2": "", "comp1": "", "comp2": "", "grip": "Não usa"},
        {"name": "TOMMY",  "normal": "0.95", "boca1": "0.70 VERTICAL", "boca2": "", "comp1": "", "comp2": "", "grip": "Não usa"},
        {"name": "DMRs",   "normal": "1.25", "boca1": "1.55 MIRA AMP.", "boca2": "", "comp1": "", "comp2": "", "grip": "Punho Leve"},
    ]},
    {"category": "ARMAS DROP", "color": "#FF5252", "items": [
        {"name": "FAMAS",  "normal": "1.00", "boca1": "0.80", "boca2": "1.15", "comp1": "0.75", "comp2": "0.95", "grip": "Punho Horizontal"},
        {"name": "GROZA",  "normal": "1X 0.85", "boca1": "2x 1.20", "boca2": "", "comp1": "", "comp2": "", "grip": "Não usa"},
        {"name": "MG3",    "normal": "PÉ 0.50", "boca1": "DEITA DESL.", "boca2": "", "comp1": "", "comp2": "", "grip": "Não usa"},
        {"name": "P90",    "normal": "0.50", "boca1": "DESLIGADO", "boca2": "", "comp1": "", "comp2": "", "grip": "Não usa"},
    ]}
]

def check_match(val_str, current_val):
    if not val_str: return False
    val_str = str(val_str).upper()
    if "DESLIGADO" in val_str: return False
    
    matches = re.findall(r"(\d+\.\d+)", val_str)
    if not matches: return False
    
    floats = [float(m) for m in matches]
    for f in floats:
        if abs(f - current_val) < 0.001:
            return True
            
    return False


class ControlSensiApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Color Palette - Modern PUBG Gaming Theme
        self.c_bg = "#0A0A0C"            # Very deep background
        self.c_panel = "#141416"         # Panel background
        self.c_border = "#2A2A2E"        # Subtle borders
        self.c_accent = "#F2A900"        # PUBG Yellow
        self.c_green = "#00FF66"         # Neon green for active values
        self.c_text = "#EBEBEB"          # Main text
        self.c_muted = "#737373"         # Inactive text

        self.title("ControlSensi")
        self.geometry("780x600")
        self.minsize(750, 500)
        
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color=self.c_bg)

        self.current_value = 1.00
        self.always_on_top = False
        
        self.voice_enabled = True
        self.volume = 0.5
        self.tts_device = "Padrão"
        
        self.hotkey_up = "page up"
        self.hotkey_down = "page down"

        self.audio_devices = self.get_audio_devices()

        self.load_config()
        self.attributes("-topmost", self.always_on_top)
        
        self.tts_queue = queue.Queue()
        self.tts_thread = threading.Thread(target=tts_worker, args=(self.tts_queue, lambda: self.volume, lambda: self.voice_enabled, lambda: self.tts_device), daemon=True)
        self.tts_thread.start()

        # ================== LAYOUT ==================
        self.topmost_checkbox = ctk.CTkCheckBox(
            self, text="Sempre no topo", 
            command=self.toggle_topmost,
            fg_color=self.c_accent, hover_color="#C98B00", text_color=self.c_muted
        )
        if self.always_on_top:
            self.topmost_checkbox.select()
        self.topmost_checkbox.place(relx=0.98, rely=0.02, anchor="ne")

        self.tabs = ctk.CTkTabview(
            self, 
            fg_color=self.c_panel, 
            segmented_button_fg_color=self.c_bg,
            segmented_button_selected_color=self.c_accent,
            segmented_button_selected_hover_color="#C98B00",
            segmented_button_unselected_color=self.c_panel,
            text_color="#000",
            border_width=1,
            border_color=self.c_border
        )
        self.tabs.pack(fill="both", expand=True, padx=20, pady=(35, 20))
        
        self.tab_sensi = self.tabs.add("Sensibilidade")
        self.tab_punhos = self.tabs.add("Punhos")
        self.tab_som = self.tabs.add("Configurações")
        
        self.tab_sensi.configure(fg_color="transparent")
        self.tab_punhos.configure(fg_color="transparent")
        self.tab_som.configure(fg_color="transparent")

        self.weapon_ui_refs = []
        self.grip_ui_refs = []

        self.build_sensi_tab()
        self.build_punhos_tab()
        self.build_som_tab()

        self.populate_sensi_table()
        self.populate_punhos_table()

        self.update_legend_colors()
        self.setup_hooks()

    def get_audio_devices(self):
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            outputs = speaker.GetAudioOutputs()
            return [outputs.Item(i).GetDescription() for i in range(outputs.Count)]
        except Exception:
            return ["Padrão"]

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                    self.current_value = data.get("recoil_value", 1.00)
                    self.always_on_top = data.get("always_on_top", False)
                    self.hotkey_up = data.get("hotkey_up", "page up")
                    self.hotkey_down = data.get("hotkey_down", "page down")
                    self.voice_enabled = data.get("voice_enabled", True)
                    self.volume = data.get("volume", 0.5)
                    self.tts_device = data.get("tts_device", "Padrão")
            except Exception:
                pass

    def save_config(self):
        data = {
            "recoil_value": self.current_value,
            "always_on_top": self.always_on_top,
            "hotkey_up": self.hotkey_up,
            "hotkey_down": self.hotkey_down,
            "voice_enabled": self.voice_enabled,
            "volume": self.volume,
            "tts_device": self.tts_device
        }
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(data, f, indent=4)
        except Exception:
            pass

    # ================== UI BUILDERS ==================

    def build_sensi_tab(self):
        self.tab_sensi.grid_columnconfigure(0, weight=0)
        self.tab_sensi.grid_columnconfigure(1, weight=1)
        self.tab_sensi.grid_rowconfigure(0, weight=1)

        # Left panel: Control
        self.left_frame = ctk.CTkFrame(self.tab_sensi, width=280, fg_color="transparent")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=0)

        # Logo
        logo_img = load_logo_image(max_dim=80)
        if logo_img:
            self.logo_label = ctk.CTkLabel(self.left_frame, text="", image=logo_img)
            self.logo_label.pack(pady=(5, 2))
        else:
            self.title_label = ctk.CTkLabel(self.left_frame, text="CONTROLSENSI", font=("Inter", 20, "bold"), text_color=self.c_accent)
            self.title_label.pack(pady=(10, 2))

        self.value_label = ctk.CTkLabel(self.left_frame, text=f"{self.current_value:.2f}", font=("Inter", 65, "bold"), text_color="#FFFFFF")
        self.value_label.pack(pady=0)

        self.info_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.info_frame.pack(pady=0)
        ctk.CTkLabel(self.info_frame, text="AUMENTAR (+0.05)", font=("Inter", 9, "bold"), text_color=self.c_muted).pack(side="left", padx=5)
        ctk.CTkLabel(self.info_frame, text="DIMINUIR (-0.05)", font=("Inter", 9, "bold"), text_color=self.c_muted).pack(side="right", padx=5)

        self.anim_label = ctk.CTkLabel(self.left_frame, text="", font=("Inter", 16, "bold"))
        self.anim_label.pack(pady=(2, 2))
        self.anim_timer = None

        # Destaque Box Title
        destaque_title = ctk.CTkFrame(self.left_frame, fg_color=self.c_accent, corner_radius=4)
        destaque_title.pack(fill="x", padx=10, pady=(0, 0))
        ctk.CTkLabel(destaque_title, text="ARMAS PARA USAR", font=("Inter", 12, "bold"), text_color="#000000").pack(pady=2)

        self.destaque_scroll = ctk.CTkScrollableFrame(self.left_frame, width=280, height=380, fg_color="transparent")
        self.destaque_scroll.pack(fill="both", expand=True, padx=0, pady=2)

        # Right panel: Table
        self.right_frame = ctk.CTkScrollableFrame(self.tab_sensi, fg_color="transparent")
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 5), pady=0)

    def populate_sensi_table(self):
        for widget in self.right_frame.winfo_children():
            widget.destroy()
            
        self.weapon_ui_refs = []
        row_idx = 0
        
        for cat in WEAPONS_DATA:
            ctk.CTkLabel(self.right_frame, text=cat["category"], font=("Inter", 14, "bold"), text_color=cat["color"]).grid(row=row_idx, column=0, columnspan=4, sticky="w", pady=(10, 4), padx=2)
            row_idx += 1
            
            headers = ["ARMA", "SEM NADA", "1x BOCA - 2x BOCA", "1x COMP - 2x COMP"]
            for col, text in enumerate(headers):
                ctk.CTkLabel(self.right_frame, text=text, font=("Inter", 10, "bold"), text_color=self.c_muted).grid(row=row_idx, column=col, padx=(2, 10), pady=0, sticky="w")
            row_idx += 1

            for item in cat["items"]:
                # Col 0: ARMA
                lbl_name = ctk.CTkLabel(self.right_frame, text=item["name"], font=("Inter", 13, "bold"), anchor="w")
                lbl_name.grid(row=row_idx, column=0, padx=(2, 10), pady=1, sticky="w")
                
                # Col 1: NORMAL
                lbl_norm = ctk.CTkLabel(self.right_frame, text=item["normal"], font=("Inter", 13), anchor="w")
                lbl_norm.grid(row=row_idx, column=1, padx=(2, 10), pady=1, sticky="w")
                
                # Col 2: BOCA (1x - 2x)
                boca_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
                boca_frame.grid(row=row_idx, column=2, padx=(2, 10), pady=1, sticky="w")
                
                lbl_b1 = ctk.CTkLabel(boca_frame, text=item["boca1"], font=("Inter", 13))
                lbl_b1.pack(side="left")
                
                if item["boca2"]:
                    ctk.CTkLabel(boca_frame, text=" - ", font=("Inter", 13), text_color=self.c_muted).pack(side="left")
                    lbl_b2 = ctk.CTkLabel(boca_frame, text=item["boca2"], font=("Inter", 13))
                    lbl_b2.pack(side="left")
                else:
                    lbl_b2 = None

                # Col 3: COMP (1x - 2x)
                comp_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
                comp_frame.grid(row=row_idx, column=3, padx=(2, 10), pady=1, sticky="w")
                
                lbl_c1 = ctk.CTkLabel(comp_frame, text=item["comp1"], font=("Inter", 13))
                lbl_c1.pack(side="left")
                
                if item["comp2"]:
                    ctk.CTkLabel(comp_frame, text=" - ", font=("Inter", 13), text_color=self.c_muted).pack(side="left")
                    lbl_c2 = ctk.CTkLabel(comp_frame, text=item["comp2"], font=("Inter", 13))
                    lbl_c2.pack(side="left")
                else:
                    lbl_c2 = None
                
                self.weapon_ui_refs.append({
                    "item": item,
                    "lbl_name": lbl_name,
                    "lbl_norm": lbl_norm,
                    "lbl_b1": lbl_b1,
                    "lbl_b2": lbl_b2,
                    "lbl_c1": lbl_c1,
                    "lbl_c2": lbl_c2
                })
                row_idx += 1

    def build_punhos_tab(self):
        self.punhos_scroll = ctk.CTkScrollableFrame(self.tab_punhos, fg_color="transparent")
        self.punhos_scroll.pack(fill="both", expand=True, padx=20, pady=20)

    def populate_punhos_table(self):
        for cat in WEAPONS_DATA:
            ctk.CTkLabel(self.punhos_scroll, text=cat["category"], font=("Inter", 16, "bold"), text_color=cat["color"]).pack(anchor="w", pady=(15, 5))
            
            grid_frame = ctk.CTkFrame(self.punhos_scroll, fg_color=self.c_bg, border_width=1, border_color=self.c_border, corner_radius=6)
            grid_frame.pack(fill="x", padx=10, pady=5)
            
            ctk.CTkLabel(grid_frame, text="ARMA", font=("Inter", 13, "bold"), text_color=self.c_muted, width=150, anchor="w").grid(row=0, column=0, sticky="w", padx=15, pady=(10,0))
            ctk.CTkLabel(grid_frame, text="PUNHO IDEAL", font=("Inter", 13, "bold"), text_color=self.c_muted, width=250, anchor="w").grid(row=0, column=1, sticky="w", pady=(10,0))
            
            row_idx = 1
            for item in cat["items"]:
                lbl_name = ctk.CTkLabel(grid_frame, text=item["name"], font=("Inter", 15, "bold"), width=150, anchor="w")
                lbl_name.grid(row=row_idx, column=0, pady=10, padx=15, sticky="w")
                
                grip_img = load_grip_image(item["grip"], max_dim=60)
                if grip_img:
                    lbl_grip = ctk.CTkLabel(grid_frame, text="  " + item["grip"], image=grip_img, compound="left", font=("Inter", 15), width=250, anchor="w")
                else:
                    lbl_grip = ctk.CTkLabel(grid_frame, text=item["grip"], font=("Inter", 15), width=250, anchor="w")
                    
                lbl_grip.grid(row=row_idx, column=1, pady=10, sticky="w")
                
                self.grip_ui_refs.append({
                    "item": item,
                    "lbl_name": lbl_name,
                    "lbl_grip": lbl_grip
                })
                row_idx += 1

    def build_som_tab(self):
        container = ctk.CTkFrame(self.tab_som, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=40, pady=20)
        
        # Audio
        self.audio_frame = ctk.CTkFrame(container, fg_color=self.c_bg, border_width=1, border_color=self.c_border)
        self.audio_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(self.audio_frame, text="Configurações de Áudio", font=("Inter", 16, "bold"), text_color=self.c_accent).pack(pady=(15, 10))
        
        self.voice_checkbox = ctk.CTkCheckBox(self.audio_frame, text="Voz Ativada", command=self.toggle_voice, fg_color=self.c_accent, hover_color="#C98B00")
        if self.voice_enabled:
            self.voice_checkbox.select()
        self.voice_checkbox.pack(pady=(5, 5))
        
        self.device_combo = ctk.CTkOptionMenu(self.audio_frame, values=self.audio_devices, command=self.change_tts_device, fg_color=self.c_panel, button_color=self.c_accent, button_hover_color="#C98B00")
        if self.tts_device in self.audio_devices:
            self.device_combo.set(self.tts_device)
        elif self.audio_devices:
            self.device_combo.set(self.audio_devices[0])
            self.tts_device = self.audio_devices[0]
        self.device_combo.pack(pady=15, padx=30, fill="x")
        
        vol_inner = ctk.CTkFrame(self.audio_frame, fg_color="transparent")
        vol_inner.pack(fill="x", padx=30, pady=(0, 20))
        
        ctk.CTkLabel(vol_inner, text="Volume:", font=("Inter", 12)).pack(side="left", padx=5)
        self.vol_slider = ctk.CTkSlider(vol_inner, from_=0.0, to=1.0, command=self.change_volume, button_color=self.c_accent, button_hover_color="#C98B00", progress_color=self.c_accent)
        self.vol_slider.set(self.volume)
        self.vol_slider.pack(side="left", fill="x", expand=True, padx=10)

        self.vol_entry = ctk.CTkEntry(vol_inner, width=45, justify="center", fg_color=self.c_bg, border_color=self.c_border)
        self.vol_entry.insert(0, str(int(self.volume * 100)))
        self.vol_entry.pack(side="left", padx=5)
        self.vol_entry.bind("<Return>", self.update_volume_from_entry)
        ctk.CTkLabel(vol_inner, text="%").pack(side="left")

        # Hotkeys
        self.hotkey_frame = ctk.CTkFrame(container, fg_color=self.c_bg, border_width=1, border_color=self.c_border)
        self.hotkey_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(self.hotkey_frame, text="Atalhos de Teclado", font=("Inter", 16, "bold"), text_color=self.c_accent).pack(pady=(15, 10))
        
        hk_inner = ctk.CTkFrame(self.hotkey_frame, fg_color="transparent")
        hk_inner.pack(pady=10)
        
        ctk.CTkLabel(hk_inner, text="Aumentar:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_hk_up = ctk.CTkEntry(hk_inner, width=120, fg_color=self.c_panel, border_color=self.c_border)
        self.entry_hk_up.insert(0, self.hotkey_up)
        self.entry_hk_up.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(hk_inner, text="Diminuir:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_hk_down = ctk.CTkEntry(hk_inner, width=120, fg_color=self.c_panel, border_color=self.c_border)
        self.entry_hk_down.insert(0, self.hotkey_down)
        self.entry_hk_down.grid(row=1, column=1, padx=5, pady=5)
        
        self.save_hk_btn = ctk.CTkButton(hk_inner, text="Salvar Atalhos", command=self.update_hotkeys, fg_color=self.c_accent, hover_color="#C98B00", text_color="#000")
        self.save_hk_btn.grid(row=2, column=0, columnspan=2, pady=15)

        # Reset button
        self.reset_btn = ctk.CTkButton(container, text="RESETAR SENSI (1.00)", font=("Inter", 14, "bold"), fg_color="#E53935", hover_color="#B71C1C", command=lambda: self.schedule_action(self.reset))
        self.reset_btn.pack(pady=30)

    # ================== LOGIC ==================

    def update_legend_colors(self):
        for widget in self.destaque_scroll.winfo_children():
            widget.destroy()

        active_weapons = []

        for ref in self.weapon_ui_refs:
            item = ref["item"]
            
            m_norm = check_match(item["normal"], self.current_value)
            m_b1 = check_match(item["boca1"], self.current_value)
            m_b2 = check_match(item["boca2"], self.current_value)
            m_c1 = check_match(item["comp1"], self.current_value)
            m_c2 = check_match(item["comp2"], self.current_value)
            
            is_any_active = m_norm or m_b1 or m_b2 or m_c1 or m_c2
            if is_any_active:
                matches_str = []
                if m_norm: matches_str.append("NORMAL")
                if m_b1: matches_str.append("BOCA 1X")
                if m_b2: matches_str.append("BOCA 2X")
                if m_c1: matches_str.append("COMP 1X")
                if m_c2: matches_str.append("COMP 2X")
                active_weapons.append((item, " + ".join(matches_str)))
                
            color_active = self.c_green
            color_inactive_name = self.c_text
            color_inactive_val = self.c_muted
            
            ref["lbl_name"].configure(text_color=color_active if is_any_active else color_inactive_name)
            ref["lbl_norm"].configure(text_color=color_active if m_norm else color_inactive_val)
            ref["lbl_b1"].configure(text_color=color_active if m_b1 else color_inactive_val)
            if ref["lbl_b2"]: ref["lbl_b2"].configure(text_color=color_active if m_b2 else color_inactive_val)
            ref["lbl_c1"].configure(text_color=color_active if m_c1 else color_inactive_val)
            if ref["lbl_c2"]: ref["lbl_c2"].configure(text_color=color_active if m_c2 else color_inactive_val)
            
        for ref in self.grip_ui_refs:
            item = ref["item"]
            is_active = any(w[0]["name"] == item["name"] for w in active_weapons)
            
            ref["lbl_name"].configure(text_color=self.c_green if is_active else self.c_text)
            ref["lbl_grip"].configure(text_color=self.c_green if is_active else self.c_muted)

        if not active_weapons:
            ctk.CTkLabel(self.destaque_scroll, text="NENHUMA ARMA", text_color=self.c_muted, font=("Inter", 12, "bold")).pack(pady=20)
        else:
            for item, match_info in active_weapons:
                # Stylish highlight frame
                f = ctk.CTkFrame(self.destaque_scroll, fg_color=self.c_bg, border_width=1, border_color=self.c_border, corner_radius=6)
                f.pack(fill="x", pady=6, padx=2)
                
                # Top header for name and matching slot
                top_row = ctk.CTkFrame(f, fg_color="transparent")
                top_row.pack(fill="x", padx=12, pady=(10, 4))
                lbl_name_top = ctk.CTkLabel(top_row, text=item["name"], font=("Inter", 17, "bold"), text_color=self.c_green)
                lbl_name_top.pack(side="left")
                lbl_match = ctk.CTkLabel(top_row, text=f"{match_info}", font=("Inter", 11, "bold"), text_color=self.c_accent)
                lbl_match.pack(side="right")
                
                # Bottom content for Weapon Image and Grip
                bot_row = ctk.CTkFrame(f, fg_color="transparent")
                bot_row.pack(fill="x", padx=12, pady=(2, 10))
                
                lbl_img = None
                weapon_img = load_weapon_image(item["name"], max_dim=75)
                if weapon_img:
                    lbl_img = ctk.CTkLabel(bot_row, text="", image=weapon_img)
                    lbl_img.pack(side="left", padx=(0, 15))
                
                lbl_grip = None
                grip_text = item["grip"]
                if grip_text and grip_text != "Não usa":
                    grip_img = load_grip_image(grip_text, max_dim=50)
                    if grip_img:
                        lbl_grip = ctk.CTkLabel(bot_row, text=" " + grip_text, image=grip_img, compound="left", font=("Inter", 13, "bold"), text_color=self.c_text)
                    else:
                        lbl_grip = ctk.CTkLabel(bot_row, text=f"• {grip_text}", font=("Inter", 13), text_color=self.c_text)
                    lbl_grip.pack(side="left")

                # Interactive Hover Effect
                def on_enter(e, widget=f):
                    widget.configure(fg_color=self.c_panel, border_color=self.c_accent)
                def on_leave(e, widget=f):
                    widget.configure(fg_color=self.c_bg, border_color=self.c_border)
                    
                elements = [f, top_row, bot_row, lbl_name_top, lbl_match]
                if lbl_img: elements.append(lbl_img)
                if lbl_grip: elements.append(lbl_grip)
                
                for el in elements:
                    el.bind("<Enter>", on_enter)
                    el.bind("<Leave>", on_leave)

    def change_tts_device(self, val):
        self.tts_device = val
        self.save_config()
        self.queue_voice("Dispositivo alterado")

    def toggle_topmost(self):
        self.always_on_top = self.topmost_checkbox.get() == 1
        self.attributes("-topmost", self.always_on_top)
        self.save_config()

    def toggle_voice(self):
        self.voice_enabled = self.voice_checkbox.get() == 1
        self.save_config()

    def change_volume(self, val):
        self.volume = float(val)
        self.vol_entry.delete(0, "end")
        self.vol_entry.insert(0, str(int(self.volume * 100)))
        self.save_config()

    def update_volume_from_entry(self, event=None):
        try:
            val = int(self.vol_entry.get())
            val = max(0, min(100, val))
            self.volume = val / 100.0
            self.vol_slider.set(self.volume)
            self.vol_entry.delete(0, "end")
            self.vol_entry.insert(0, str(val))
            self.save_config()
            self.queue_voice("Volume atualizado")
        except ValueError:
            self.vol_entry.delete(0, "end")
            self.vol_entry.insert(0, str(int(self.volume * 100)))

    def update_hotkeys(self):
        keyboard.unhook_all()
        self.hotkey_up = self.entry_hk_up.get().strip().lower()
        self.hotkey_down = self.entry_hk_down.get().strip().lower()
        self.save_config()
        self.setup_keyboard_hooks()
        
        original_text = self.save_hk_btn.cget("text")
        self.save_hk_btn.configure(text="Salvo!")
        self.after(1500, lambda: self.save_hk_btn.configure(text=original_text))

    def setup_hooks(self):
        self.mouse_listener = mouse.Listener(on_click=self.on_mouse_click)
        self.mouse_listener.start()
        self.setup_keyboard_hooks()
        
    def setup_keyboard_hooks(self):
        keyboard.unhook_all()
        try:
            keyboard.on_press_key(self.hotkey_up, lambda e: self.schedule_action(self.increase))
            keyboard.on_press_key(self.hotkey_down, lambda e: self.schedule_action(self.decrease))
        except Exception:
            pass

    def on_mouse_click(self, x, y, button, pressed):
        if pressed:
            if button == mouse.Button.x2:
                self.schedule_action(self.increase)
            elif button == mouse.Button.x1:
                self.schedule_action(self.decrease)

    def schedule_action(self, action_func):
        self.after(0, action_func)

    def increase(self):
        if self.current_value < 2.00:
            self.current_value = round(self.current_value + 0.05, 2)
            if self.current_value > 2.00:
                self.current_value = 2.00
            
            self.update_main_label()
            self.show_animation("▲ +0.05", self.c_green)
            self.queue_voice(f"{self.current_value:.2f}")
            self.update_legend_colors()
            self.save_config()

    def decrease(self):
        if self.current_value > 0.50:
            self.current_value = round(self.current_value - 0.05, 2)
            if self.current_value < 0.50:
                self.current_value = 0.50
                
            self.update_main_label()
            self.show_animation("▼ -0.05", "#F44336")
            self.queue_voice(f"{self.current_value:.2f}")
            self.update_legend_colors()
            self.save_config()

    def reset(self):
        if self.current_value != 1.00:
            self.current_value = 1.00
            self.update_main_label()
            self.show_animation("↺ RESETADO", self.c_accent)
            self.queue_voice("1 ponto 0")
            self.update_legend_colors()
            self.save_config()

    def update_main_label(self):
        self.value_label.configure(text=f"{self.current_value:.2f}")

    def show_animation(self, text, color):
        if self.anim_timer:
            self.after_cancel(self.anim_timer)
            
        self.anim_label.configure(text=text, text_color=color)
        self.anim_timer = self.after(1000, lambda: self.anim_label.configure(text=""))

    def queue_voice(self, text_to_say):
        while not self.tts_queue.empty():
            try:
                self.tts_queue.get_nowait()
            except queue.Empty:
                break
        
        text_spoken = text_to_say.replace(".", " ponto ")
        self.tts_queue.put(text_spoken)

if __name__ == "__main__":
    app = ControlSensiApp()
    app.mainloop()
