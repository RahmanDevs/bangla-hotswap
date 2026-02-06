import keyboard
import win32com.client
import pythoncom

import requests
import urllib3
import bkit
urllib3.disable_warnings()



# ============================================================================
# CONVERTER CLASSES for Unicode → Bijoy
# ============================================================================
class UnicodeToBijoyConverter:

    @staticmethod
    def convert(input_text: str) -> str:
        normalized_text = bkit.transform.normalize_characters(input_text)

        response = requests.post(
            "https://fontengine.bangla.gov.bd/conversion/text/unicodeToAscii",
            json={
                "inputFont": "UTF-8",
                "outputFont": "SutonnyMj",
                "inputText": normalized_text
            },
            headers={"Content-Type": "application/json"},
            verify=False,
            timeout=10
        )

        data = response.json()
        return data["outputText"]
    
# ============================================================================
# CONVERTER CLASS for Bijoy → Unicode
# ============================================================================

class BijoyToUnicodeConverter:

    @staticmethod
    def convert(input_text: str) -> str:
        

        response = requests.post(
            "https://fontengine.bangla.gov.bd/conversion/text/asciiToUnicode",
            json={
                "inputFont": "SutonnyMj",
                "outputFont": "UTF-8",
                "inputText": input_text
            },
            headers={"Content-Type": "application/json"},
            verify=False,
            timeout=10
        )

        data = response.json()
        normalized_text = bkit.transform.normalize_characters(data["outputText"])
        return normalized_text




# ============================================================================
# SCRIPT INFORMATION
# ============================================================================


"""
Word Bangla Converter - Dual Converter (Unicode ↔ Bijoy)
Converts between Unicode and Bijoy formats in Microsoft Word

Hotkeys:
- Ctrl+Alt+Shift+B : Unicode → Bijoy (SutonnyMJ)
- Ctrl+Alt+Shift+U : Bijoy → Unicode (Kalpurush)
"""



# ============================================================================
# CONFIGURATION - Change these settings as needed
# ============================================================================

# Font settings
BIJOY_FONT = "SutonnyMJ"      # Font for Bijoy text
UNICODE_FONT = "Kalpurush"     # Font for Unicode text

# Hotkey settings
UNICODE_TO_BIJOY_HOTKEY = "ctrl+alt+shift+b"  # Convert Unicode → Bijoy
BIJOY_TO_UNICODE_HOTKEY = "ctrl+alt+shift+u"  # Convert Bijoy → Unicode

# Preserve original formatting (bold, italic, font size)
PRESERVE_FORMATTING = True

# ============================================================================

def convert_unicode_to_bijoy():
    """
    Converts selected Unicode text to Bijoy and applies Bijoy font.
    """
    try:
        print(f"Converting Unicode → Bijoy ({BIJOY_FONT})...")
        
        pythoncom.CoInitialize()
        
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except:
            print("❌ Error: Microsoft Word is not running.")
            return
        
        if word.Documents.Count == 0:
            print("❌ Error: No Word document is open.")
            return
        
        selection = word.Selection
        
        if selection.Text is None or len(selection.Text.strip()) == 0:
            print("⚠️  No text selected in Word.")
            return
        
        # Store original formatting
        original_size = selection.Font.Size
        original_bold = selection.Font.Bold
        original_italic = selection.Font.Italic
        original_underline = selection.Font.Underline
        original_color = selection.Font.Color
        
        # Get selected text
        selected_text = selection.Text
        if selected_text.endswith('\r'):
            selected_text = selected_text[:-1]
        
        original_start = selection.Start
        
        # Convert to Bijoy
        converter = UnicodeToBijoyConverter()
        bijoy_text = converter.convert(selected_text)
        
        if bijoy_text == selected_text:
            print("ℹ️  Text unchanged (may already be in Bijoy)")
            return
        
        # Replace text
        selection.TypeText(bijoy_text)
        
        # Select the new text
        new_range = word.ActiveDocument.Range(original_start, original_start + len(bijoy_text))
        new_range.Select()
        
        # Apply Bijoy font
        word.Selection.Font.Name = BIJOY_FONT
        
        # Restore formatting if configured
        if PRESERVE_FORMATTING:
            word.Selection.Font.Size = original_size
            word.Selection.Font.Bold = original_bold
            word.Selection.Font.Italic = original_italic
            word.Selection.Font.Underline = original_underline
            word.Selection.Font.Color = original_color
            print(f"✓ Converted to Bijoy (Font: {BIJOY_FONT}, Size: {original_size})")
        else:
            print(f"✓ Converted to Bijoy (Font: {BIJOY_FONT})")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def convert_bijoy_to_unicode():
    """
    Converts selected Bijoy text to Unicode and applies Unicode font.
    """
    try:
        print(f"Converting Bijoy → Unicode ({UNICODE_FONT})...")
        
        pythoncom.CoInitialize()
        
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except:
            print("❌ Error: Microsoft Word is not running.")
            return
        
        if word.Documents.Count == 0:
            print("❌ Error: No Word document is open.")
            return
        
        selection = word.Selection
        
        if selection.Text is None or len(selection.Text.strip()) == 0:
            print("⚠️  No text selected in Word.")
            return
        
        # Store original formatting
        original_size = selection.Font.Size
        original_bold = selection.Font.Bold
        original_italic = selection.Font.Italic
        original_underline = selection.Font.Underline
        original_color = selection.Font.Color
        
        # Get selected text
        selected_text = selection.Text
        if selected_text.endswith('\r'):
            selected_text = selected_text[:-1]
        
        original_start = selection.Start
        
        # Convert to Unicode
        converter = BijoyToUnicodeConverter()
        unicode_text = converter.convert(selected_text)
        
        if unicode_text == selected_text:
            print("ℹ️  Text unchanged (may already be in Unicode)")
            return
        
        # Replace text
        selection.TypeText(unicode_text)
        
        # Select the new text
        new_range = word.ActiveDocument.Range(original_start, original_start + len(unicode_text))
        new_range.Select()
        
        # Apply Unicode font
        word.Selection.Font.Name = UNICODE_FONT
        
        # Restore formatting if configured
        if PRESERVE_FORMATTING:
            word.Selection.Font.Size = original_size
            word.Selection.Font.Bold = original_bold
            word.Selection.Font.Italic = original_italic
            word.Selection.Font.Underline = original_underline
            word.Selection.Font.Color = original_color
            print(f"✓ Converted to Unicode (Font: {UNICODE_FONT}, Size: {original_size})")
        else:
            print(f"✓ Converted to Unicode (Font: {UNICODE_FONT})")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def main():
    """
    Main function to set up both hotkey listeners.
    """
    print("=" * 70)
    print(" " * 10 + "Word Bangla Converter - Dual Mode (Unicode ↔ Bijoy)")
    print("=" * 70)
    print(f"\n⚙️  Current Configuration:")
    print(f"   • Unicode Font: {UNICODE_FONT}")
    print(f"   • Bijoy Font: {BIJOY_FONT}")
    print(f"   • Preserve Formatting: {'Yes' if PRESERVE_FORMATTING else 'No'}")
    print("\n📋 Hotkeys:")
    print(f"   • {UNICODE_TO_BIJOY_HOTKEY.upper()} : Unicode → Bijoy ({BIJOY_FONT})")
    print(f"   • {BIJOY_TO_UNICODE_HOTKEY.upper()} : Bijoy → Unicode ({UNICODE_FONT})")
    print(f"   • ESC : Exit this script")
    print("\n💡 Usage:")
    print("   1. Select Bangla text in Microsoft Word")
    print("   2. Press the appropriate hotkey to convert")
    print("   3. Text will be converted and font will change automatically")
    print("\n💡 Features:")
    print("   • Bidirectional conversion (Unicode ↔ Bijoy)")
    print("   • Direct text replacement (no clipboard interference)")
    print("   • Automatic font switching")
    print("   • Undo support (Ctrl+Z works after conversion)")
    print("   • Preserves formatting (bold, italic, size, color)")
    print("\n⚠️  Requirements:")
    print("   • Microsoft Word must be running")
    print("   • A document must be open")
    print(f"   • {BIJOY_FONT} and {UNICODE_FONT} fonts must be installed")
    print("\n💡 Tip: Edit the CONFIGURATION section at the top of this")
    print("   script to change fonts or hotkeys")
    print("=" * 70)
    print(f"\n⏳ Waiting for hotkeys...\n")
    
    # Register both hotkeys
    keyboard.add_hotkey(UNICODE_TO_BIJOY_HOTKEY, convert_unicode_to_bijoy)
    keyboard.add_hotkey(BIJOY_TO_UNICODE_HOTKEY, convert_bijoy_to_unicode)
    
    # Keep the script running until ESC is pressed
    keyboard.wait('esc')
    
    print("\n👋 Script terminated. Goodbye!")

if __name__ == "__main__":
    main()
