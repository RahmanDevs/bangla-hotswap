# Ctrl Alt Bangla

Convert between Bangla Unicode and Bijoy formats directly in Microsoft Word with keyboard shortcuts.

## Features

- 🔄 **Bidirectional conversion**: Unicode ↔ Bijoy
- ⌨️ **Keyboard shortcuts**: Instant conversion with hotkeys
- 🎨 **Auto font switching**: Applies correct font automatically
- 💾 **No clipboard interference**: Your clipboard stays untouched
- ↩️ **Undo support**: Press Ctrl+Z to undo any conversion
- ✨ **Format preservation**: Keeps bold, italic, size, and color

## Installation

```bash
pip install -r requirements.txt
python -m pywin32_postinstall -install
```

## Usage

1. **Run the script** (as Administrator):
   ```bash
   python main.py
   ```

2. **In Microsoft Word**, select Bangla text and press:
   - **Ctrl+Alt+Shift+B** → Convert to Bijoy (SutonnyMJ font)
   - **Ctrl+Alt+Shift+U** → Convert to Unicode (Kalpurush font)

3. Press **ESC** to exit the script

## Configuration

Edit the top of `main.py` to customize:

```python
BIJOY_FONT = "SutonnyMJ"                       # Bijoy font
UNICODE_FONT = "Kalpurush"                     # Unicode font
UNICODE_TO_BIJOY_HOTKEY = "ctrl+alt+shift+b"   # Hotkey for Unicode → Bijoy
BIJOY_TO_UNICODE_HOTKEY = "ctrl+alt+shift+u"   # Hotkey for Bijoy → Unicode
PRESERVE_FORMATTING = True                     # Keep text formatting
```

## Requirements

- Windows OS
- Microsoft Word
- Python 3.7+
- SutonnyMJ and Kalpurush fonts installed

## Troubleshooting

**Hotkey not working?**
- Run Command Prompt/PowerShell as Administrator

**"Microsoft Word is not running"?**
- Open Word with a document before running the script

**Missing modules?**
```bash
pip install keyboard pywin32 banglaGovBD
```

