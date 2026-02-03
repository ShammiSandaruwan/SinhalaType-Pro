import win32com.client
import unicodedata
import re

def test_photoshop_connection():
    print("--- Testing Photoshop Connection ---")
    try:
        # Try GetActiveObject
        print("Attempting GetActiveObject...")
        ps_app = win32com.client.GetActiveObject("Photoshop.Application")
        print("Success! Connected to Photoshop.")
        print(f"Photoshop Version: {ps_app.Version}")
        if ps_app.Documents.Count > 0:
            print(f"Open Documents: {ps_app.Documents.Count}")
            print(f"Active Document: {ps_app.ActiveDocument.Name}")
        else:
            print("No documents open.")
        return True
    except Exception as e:
        print(f"GetActiveObject Failed: {e}")
        
    try:
        # Try Dispatch
        print("\nAttempting Dispatch (CreateObject)...")
        ps_app = win32com.client.Dispatch("Photoshop.Application")
        print("Success! dispatched Photoshop.")
        print(f"Photoshop Version: {ps_app.Version}")
        return True
    except Exception as e:
        print(f"Dispatch Failed: {e}")
        
    print("\nCONCLUSION: Could not connect to Photoshop. Ensure it is running. Try running this script as Administrator.")
    return False

def test_conversion_logic(text):
    print(f"\n--- Testing Conversion Logic for '{text}' ---")
    
    # 1. Normalize
    normalized = unicodedata.normalize('NFC', text)
    print(f"Normalized (NFC): {[hex(ord(c)) for c in normalized]}")
    
    # 2. Rephaya
    text_reph = re.sub(r'ර්([ක-ෆ])', r'\1' + '¾', normalized) # using raw char for test
    if text_reph != normalized:
        print(f"Rephaya Logic Triggered: {text_reph}")
        
    # 3. Full Process (Simplified from main app)
    # Copying mapping from main app main logic for quick test
    MAPPING = {'අ': 'w', 'ම': 'u', '්': 'A', 'ා': 'd'} # Partial map for Amma
    
    chars = list(normalized)
    result = []
    i = 0
    while i < len(chars):
        char = chars[i]
        result.append(MAPPING.get(char, char))
        i += 1
    
    final = "".join(result)
    print(f"Simple Map Result: {final}")
    
    # Check expected for "අම්මා" -> w u A u d (wuAud)
    expected = "wuAmd" # Wait, al-lakuna + ma? NO.
    # අ = w
    # ම් = ම + ් = u + A
    # මා = ම + ා = u + d
    # wuAud
    
    print(f"Heuristic Check: If input was 'අම්මා', expected starts with 'w'. Got: {final[0] if final else 'Empty'}")
    
    # Check Clipboard
    import pyperclip
    pyperclip.copy(final)
    print(f"\n[INFO] Copied '{final}' to clipboard.")
    print("IMPORTANT: If you paste this into Notepad/Browser, it WILL look like English text (e.g. 'wuAud').")
    print("This is CORRECT. It will only look like Sinhala if you paste it into Photoshop and select 'FMAbhaya' font.")
    return final

if __name__ == "__main__":
    test_photoshop_connection()
    test_conversion_logic("අම්මා")
    input("\nPress Enter to exit...")
