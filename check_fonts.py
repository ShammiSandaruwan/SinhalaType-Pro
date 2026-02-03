import win32com.client

def list_fonts():
    print("--- Checking Photoshop Fonts ---")
    try:
        ps_app = win32com.client.GetActiveObject("Photoshop.Application")
    except Exception:
        try:
            ps_app = win32com.client.Dispatch("Photoshop.Application")
        except Exception as e:
            print(f"Could not connect to Photoshop: {e}")
            return

    print(f"Photoshop Version: {ps_app.Version}")
    print("Scanning for FM/Abhaya fonts...")
    
    kms = ps_app.Fonts
    found = False
    for i in range(1, kms.Count + 1):
        font = kms.Item(i)
        name = font.Name
        family = font.Family
        postscript = font.PostScriptName
        
        # Check for target fonts
        if "FM" in name or "Abhaya" in name or "Isi" in name or "FM" in postscript:
            print(f"FOUND: Name='{name}', Family='{family}', PostScript='{postscript}'")
            found = True
            
    if not found:
        print("WARNING: No fonts with 'FM', 'Abhaya', or 'Isi' found in Photoshop font list.")
        print("Please ensure 'FMAbhaya' is installed.")

if __name__ == "__main__":
    list_fonts()
    input("\nPress Enter to exit...")
