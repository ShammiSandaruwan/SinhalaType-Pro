import customtkinter as ctk
import win32com.client
import pyperclip
import re
import threading
import unicodedata
import sys
import os

# --- CONFIGURATION ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# --- CONVERSION LOGIC ---
class SinhalaConverter:
    def __init__(self):
        # Full Unicode to FM Abhaya Mapping
        self.MAPPING = {
            # Independent Vowels
            'අ': 'w', 'ආ': 'wd', 'ඇ': 'we', 'ඈ': 'wE', 'ඉ': 'b', 'ඊ': 'B', 
            'උ': 'W', 'ඌ': 'W', 'එ': 't', 'ඒ': 'ta', 'ඔ': 'T', 'ඕ': 'T', 'ඖ': 'T',
            
            # Consonants
            'ක': 'l', 'ඛ': 'L', 'ග': '.', 'ඝ': '>', 'ඞ': '?', 'ඟ': 'ň',
            'ච': 'p', 'ඡ': 'P', 'ජ': 'c', 'ඣ': 'C', 'ඤ': 'Z', 'ඥ': '[', 'ඦ': 'x',
            'ට': 'gd', 'ඨ': 'GD', 'ඩ': 'v', 'ඪ': 'V', 'ණ': 'K', 'ඬ': 'n',
            'ත': ';', 'ථ': ':', 'ද': 'o', 'ධ': 'O', 'න': 'k', 'ඳ': 'o',
            'ප': 'm', 'ඵ': 'M', 'බ': 'n', 'භ': 'N', 'ම': 'u', 'ඹ': 'B',
            'ය': 'h', 'ර': 'r', 'ල': ',', 'ව': 'j', 'ශ': 'Y', 'ෂ': 'I',
            'ස': 'i', 'හ': 'y', 'ළ': 'e', 'ෆ': 'q',
            
            # Modifiers (Pillam)
            'ා': 'd', 'ැ': 'e', 'ෑ': 'E', 'ි': 's', 'ී': 'S', 
            'ු': 'q', 'ූ': 'Q', # Basic Papilla (See logic for variations)
            'ෘ': 'D', 
            'ෙ': 'f', 'ේ': 'fa', 'ෛ': 'ff', # Kombuwa bases
            'ඃ': 'H', 'ං': 'x', '්': 'A'
        }
        
        # Special Mappings for Rephaya/Rakaaransaya (FM Standards)
        self.REPHAYA_CHAR = '¾'       # Alt+0190 in FM
        self.RAKAARANSAYA_CHAR = 'ƒ'  # Alt+0131 in FM (Alternative to S which is Diga Ispilla)

    def process(self, text):
        if not text:
            return ""

        # 1. Normalize Unicode (NFC)
        text = unicodedata.normalize('NFC', text)
        
        # 2. Remove ZWJ/ZWNJ which confuse logic
        text = text.replace('\u200d', '').replace('\u200c', '')

        # 3. Handle Rephaya (Ra + Hal + Consonant -> Consonant + Rephaya)
        # Regex: ර් ([ක-ෆ]) -> \1 + REPHAYA
        text = re.sub(r'ර්([ක-ෆ])', r'\1' + self.REPHAYA_CHAR, text)

        # 4. Handle Rakaaransaya (Consonant + Hal + Ra -> Consonant + Rakaaransaya)
        # Regex: ([ක-ෆ])්ර -> \1 + RAKAARANSAYA
        text = re.sub(r'([ක-ෆ])්ර', r'\1' + self.RAKAARANSAYA_CHAR, text)

        # 5. Tokenize and Reorder Kombuwa (The "Split-Swap-Append" Rule)
        chars = list(text)
        result = []
        i = 0
        length = len(chars)

        while i < length:
            char = chars[i]
            
            # Lookahead for Vowel Signs
            next_char = chars[i+1] if i + 1 < length else None
            
            # Define Kombuwa variants
            is_kombuwa = next_char == 'ෙ'
            is_kombuwa_aa = next_char == 'ේ'
            is_kombu_dekka = next_char == 'ෛ'
            is_kombuwa_aela = next_char == 'ො' # o
            is_kombuwa_diga_aela = next_char == 'ෝ' # O
            is_kombuwa_gayanukitta = next_char == 'ෞ' # au

            if char in self.MAPPING and (is_kombuwa or is_kombuwa_aa or is_kombu_dekka or 
                                         is_kombuwa_aela or is_kombuwa_diga_aela or is_kombuwa_gayanukitta):
                
                # A. Append the Prefix (Kombuwa)
                if is_kombu_dekka:
                    result.append('ff')
                else:
                    result.append('f') # Standard Kombuwa
                
                # B. Append the Consonant (Mapped)
                result.append(self.MAPPING.get(char, char))
                
                # C. Append the Suffix (if complex vowel)
                if is_kombuwa_aa:
                    result.append('a') # Aela Pilla (roughly)
                elif is_kombuwa_aela: # ො (o)
                    result.append('d') # Aela Pilla
                elif is_kombuwa_diga_aela: # ෝ (O)
                    result.append('D') # Diga Aela Pilla (varies by font, usually D)
                elif is_kombuwa_gayanukitta: # ෞ (au)
                    result.append('W') # Gayanukitta mapping
                
                i += 2 # Skip both consonant and vowel
            else:
                # Normal Mapping
                result.append(self.MAPPING.get(char, char))
                i += 1
                
        return "".join(result)


# --- GUI APPLICATION ---
class SinhalaTypeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window Setup
        self.title("SinhalaType Pro")
        self.geometry("400x520")
        self.resizable(False, False)
        self.attributes('-topmost', True) # Float on top
        
        # Data
        self.converter = SinhalaConverter()
        
        # Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. Header
        self.lbl_header = ctk.CTkLabel(self, text="SinhalaType Pro", font=("Roboto Medium", 24))
        self.lbl_header.grid(row=0, column=0, pady=(20, 10))

        # 2. Input
        self.lbl_input = ctk.CTkLabel(self, text="Input (Unicode/Singlish):", anchor="w")
        self.lbl_input.grid(row=1, column=0, padx=20, sticky="w")

        self.txt_input = ctk.CTkTextbox(self, height=150)
        self.txt_input.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")

        # 3. Settings
        self.frm_settings = ctk.CTkFrame(self, fg_color="transparent")
        self.frm_settings.grid(row=3, column=0, padx=20, pady=5, sticky="ew")

        self.lbl_font = ctk.CTkLabel(self.frm_settings, text="Target Font:")
        self.lbl_font.pack(side="left")
        
        self.opt_font = ctk.CTkOptionMenu(self.frm_settings, values=["FM Abhaya", "ISI Fonts", "Unicode"])
        self.opt_font.pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.chk_autocopy = ctk.CTkCheckBox(self, text="Auto-Copy to Clipboard")
        self.chk_autocopy.grid(row=4, column=0, padx=20, pady=10, sticky="w")
        self.chk_autocopy.select()

        # 4. Buttons
        self.frm_buttons = ctk.CTkFrame(self, fg_color="transparent")
        self.frm_buttons.grid(row=5, column=0, padx=20, pady=10)

        self.btn_ps = ctk.CTkButton(self.frm_buttons, text="Send to Photoshop", 
                                    command=self.start_ps_thread, 
                                    fg_color="#E31C5F", hover_color="#C41550", width=170)
        self.btn_ps.pack(side="left", padx=5)

        self.btn_copy = ctk.CTkButton(self.frm_buttons, text="Copy Text", 
                                      command=self.copy_text, width=120)
        self.btn_copy.pack(side="left", padx=5)

        # 5. Status
        self.lbl_status = ctk.CTkLabel(self, text="Ready", text_color="gray")
        self.lbl_status.grid(row=6, column=0, pady=(0, 20))

    def copy_text(self):
        raw_text = self.txt_input.get("0.0", "end").strip()
        if not raw_text:
            self.update_status("Error: Input is empty", "red")
            return
        
        converted = self.converter.process(raw_text)
        pyperclip.copy(converted)
        self.update_status("Copied to clipboard!", "green")

    def start_ps_thread(self):
        """Runs Photoshop automation in a background thread to prevent UI freeze."""
        threading.Thread(target=self.send_to_photoshop, daemon=True).start()

    def send_to_photoshop(self):
        raw_text = self.txt_input.get("0.0", "end").strip()
        if not raw_text:
            self.update_status("Error: Input is empty", "red")
            return

        self.update_status("Connecting to Photoshop...", "orange")

        try:
            # 1. Connect to Photoshop
            ps_app = None
            try:
                ps_app = win32com.client.GetActiveObject("Photoshop.Application")
            except Exception as e:
                print(f"GetActiveObject failed: {e}")
                try:
                    ps_app = win32com.client.Dispatch("Photoshop.Application")
                except Exception as e2:
                    print(f"Dispatch failed: {e2}")
                    self.update_status(f"Error: Could not connect to Photoshop. {e}", "red")
                    return

            if ps_app.Documents.Count == 0:
                self.update_status("Error: No document open", "red")
                return

            # 2. Convert Text
            font_mode = self.opt_font.get()
            final_text = raw_text
            
            if font_mode == "FM Abhaya":
                final_text = self.converter.process(raw_text)

            # 3. Add Layer
            doc = ps_app.Application.ActiveDocument
            art_layer = doc.ArtLayers.Add()
            art_layer.Kind = 2 # Text Layer
            
            text_item = art_layer.TextItem
            text_item.Contents = final_text
            text_item.Size = 36 # Default size
            
            # 4. Set Font (With Fallback)
            try:
                if font_mode == "FM Abhaya":
                    text_item.Font = "FMAbhaya"
                elif font_mode == "ISI Fonts":
                    text_item.Font = "IsiAbhaya" # Change to specific ISI font name if known
                else:
                    text_item.Font = "Iskoola Pota"
            except Exception:
                print("Font not found, defaulting")
                self.update_status("Warning: Target font missing", "orange")

            # Auto Copy check
            if self.chk_autocopy.get():
                pyperclip.copy(final_text)

            self.update_status("Success: Layer Added!", "green")

        except Exception as e:
            print(f"Critical Error: {e}")
            self.update_status(f"Error: {str(e)[:30]}...", "red")

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)

if __name__ == "__main__":
    # Resource path handling for PyInstaller
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    app = SinhalaTypeApp()
    app.mainloop()
