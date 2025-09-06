from tkinter import scrolledtext, messagebox
from difflib import get_close_matches
import tkinter as tk
import pandas as pd
import sys
import os
import re

# Helper to locate included files when packed into .exe
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # PyInstaller temporary folder
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Load abbreviations from Excel
def load_abbreviations_from_excel(file_path):
  df = pd.read_excel(file_path)
  df.columns = [c.strip().lower() for c in df.columns]

  if 'phrase' not in df.columns or 'abbreviation' not in df.columns:
      raise ValueError("Excel must contain 'phrase' and 'abbreviation' columns.")

  df['phrase'] = df['phrase'].astype(str).str.strip()
  df['abbreviation'] = df['abbreviation'].astype(str).str.strip()

  mapping = dict(zip(df['phrase'].str.lower(), df['abbreviation']))
  return mapping
  
# Convert plain NOTAM text into abbreviations
def encode_notam_text(plain_text, abbrev_dict):
  text = plain_text.lower()
  sorted_phrases = sorted(abbrev_dict.keys(), key=lambda x: -len(x))

  for phrase in sorted_phrases:
    abbrev = abbrev_dict[phrase]
    text = re.sub(rf"\b{re.escape(phrase)}\b", abbrev, text)

  # encoded_text = re.sub(r"\s+", " ", text).strip().upper()
  # Replace multiple spaces with one, but keep newlines
  encoded_text = re.sub(r"[ ]+", " ", text).strip().upper()

  if not encoded_text.endswith('.'):
    encoded_text += "."
  return encoded_text
  
# Convert NOTAM abbreviations back into plain text
def decode_notam_text(encoded_text, abbrev_dict):
    text = encoded_text.lower()

    # Invert dictionary: abbreviation -> list of phrases
    inv_dict = {}
    for phrase, abbr in abbrev_dict.items():
        abbr_l = abbr.lower()
        if abbr_l not in inv_dict:
            inv_dict[abbr_l] = []
        inv_dict[abbr_l].append(phrase)

    # Sort abbreviations by length (longest first)
    sorted_abbrevs = sorted(inv_dict.keys(), key=lambda x: -len(x))

    # Replace abbreviations with all matching phrases
    for abbrev in sorted_abbrevs:
        phrases = inv_dict[abbrev]
        replacement = "/".join(phrases)  # Join multiple phrases with a slash
        text = re.sub(rf"\b{re.escape(abbrev)}\b", replacement, text)

    # Replace multiple spaces with one, but keep newlines
    decoded_text = re.sub(r"[ ]+", " ", text).strip().capitalize()

    if not decoded_text.endswith('.'):
        decoded_text += "."
    return decoded_text

# Main GUI app
class NotamEncoderApp:
  def __init__(self, root):
    self.root = root
    self.root.title("NOTAM Encoder")
    self.abbrev_dict = {}

    # Automatically load Excel abbreviations
    try:
      file_path = resource_path("icao_abbreviations.xlsx")
      self.abbrev_dict = load_abbreviations_from_excel(file_path)
      # messagebox.showinfo("Success", "Abbreviations loaded successfully!")
    except Exception as e:
      messagebox.showerror("Error", f"Failed to load abbreviations: {e}")

    # Input text
    tk.Label(root, text="Enter plain NOTAM text:").pack(anchor='w', padx=10)
    self.input_text = scrolledtext.ScrolledText(root, width=100, height=15)
    self.input_text.pack(padx=10, pady=5)

	# Create a frame to hold the buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Create two buttons in the same frame
    encode_btn = tk.Button(button_frame, text="Encode", command=self.encode_text)
    encode_btn.pack(side="left", padx=5)

    decode_btn = tk.Button(button_frame, text="Decode", command=self.decode_text)
    decode_btn.pack(side="left", padx=5)

    # Output text
    tk.Label(root, text="Encoded NOTAM text:").pack(anchor='w', padx=10)
    self.output_text = scrolledtext.ScrolledText(root, width=100, height=15)
    self.output_text.pack(padx=10, pady=5)

    # Clear Button
    '''self.clear_btn = tk.Button(root, text="Clear", command=self.clear_text)
    self.clear_btn.pack(pady=5)'''

    # Copy Button
    '''self.copy_btn = tk.Button(root, text="Copy NOTAM", command=self.copy_to_clipboard)
    self.copy_btn.pack(pady=5)'''

    # Create a frame to hold the buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Create two buttons in the same frame
    clear_btn = tk.Button(button_frame, text="Clear", command=self.clear_text)
    clear_btn.pack(side="left", padx=5)

    copy_btn = tk.Button(button_frame, text="Copy NOTAM", command=self.copy_to_clipboard)
    copy_btn.pack(side="left", padx=5)

  def get_valid_words(self):
        return list(self.abbrev_dict.keys()) + list(self.abbrev_dict.values())
  
  def encode_text(self):
    plain = self.input_text.get("1.0", tk.END).strip()
    if not plain:
      messagebox.showwarning("Empty input", "Please enter some NOTAM text.")
      return
    encoded = encode_notam_text(plain, self.abbrev_dict)
    self.output_text.config(state='normal')
    self.output_text.delete("1.0", tk.END)
    self.output_text.insert(tk.END, encoded)
    self.output_text.config(state='disabled')

  def clear_text(self):
    self.input_text.delete("1.0", tk.END)
    self.output_text.config(state='normal')
    self.output_text.delete("1.0", tk.END)
    self.output_text.config(state='disabled')

  def copy_to_clipboard(self):
    encoded_text = self.output_text.get("1.0", tk.END).strip()
    if not encoded_text:
      messagebox.showwarning("No text", "There is no encoded NOTAM text to copy.")
      return
    self.root.clipboard_clear()
    self.root.clipboard_append(encoded_text)
    messagebox.showinfo("Copied", "Encoded NOTAM text copied to clipboard!")
  
  def decode_text(self):
    encoded = self.input_text.get("1.0", tk.END).strip()
    if not encoded:
        messagebox.showwarning("Empty input", "Please enter some NOTAM text.")
        return
    decoded = decode_notam_text(encoded, self.abbrev_dict)
    self.output_text.config(state='normal')
    self.output_text.delete("1.0", tk.END)
    self.output_text.insert(tk.END, decoded)
    self.output_text.config(state='disabled')

	
# Run app
if __name__ == "__main__":
  root = tk.Tk()
  root.title("NOTAM Encoder")
  app = NotamEncoderApp(root)
  root.mainloop()
