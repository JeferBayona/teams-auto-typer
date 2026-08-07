import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import openpyxl
import xlrd
import lxml
import html5lib
import bs4
import pyautogui
import threading
import time
import os
import sys

# Configure PyAutoGUI Fail-Safe: Moving mouse to any corner of the screen aborts the script.
pyautogui.FAILSAFE = True

def clean_name(full_name):
    """
    Cleans the name according to the rule:
    Dela Cruz, Juan Miguel Zerge -> Dela Cruz, Juan Miguel
    dela cruz, juan -> dela cruz, juan
    Also handles multi-word middle names like 'Dela Cruz', 'De Guzman', 'De Los Santos'
    """
    if pd.isna(full_name) or not isinstance(full_name, str):
        return None
    
    parts = full_name.split(',')
    if len(parts) < 2:
        return full_name.strip()
    
    surname = parts[0].strip()
    first_names_part = parts[1].strip()
    
    first_names = first_names_part.split()
    
    if len(first_names) >= 2:
        # Check for 3-word middle name (e.g., DE LOS SANTOS)
        if len(first_names) >= 4 and first_names[-3].upper() == "DE" and first_names[-2].upper() in ["LOS", "LAS", "LA"]:
            cleaned_first_names = " ".join(first_names[:-3])
        # Check for 2-word middle name (e.g., DELA CRUZ, DE GUZMAN, SAN JOSE)
        elif len(first_names) >= 3 and first_names[-2].upper() in ["DELA", "DE", "DEL", "SAN", "STA", "SANTA"]:
            cleaned_first_names = " ".join(first_names[:-2])
        else:
            # Standard 1-word middle name
            cleaned_first_names = " ".join(first_names[:-1])
            
        return f"{surname}, {cleaned_first_names}"
    else:
        # No middle name, keep as is
        return f"{surname}, {first_names_part}"

class TeamsBruteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Teams Auto-Typer (Brute Force)")
        self.geometry("600x650")
        
        self.filepath = None
        
        self.create_widgets()

    def create_widgets(self):
        # File Selection
        frame_file = tk.Frame(self)
        frame_file.pack(pady=10, padx=10, fill='x')
        
        self.lbl_file = tk.Label(frame_file, text="No Excel file selected", fg="blue")
        self.lbl_file.pack(side=tk.LEFT, expand=True, fill='x')
        
        btn_browse = tk.Button(frame_file, text="Browse Excel File", command=self.browse_file)
        btn_browse.pack(side=tk.RIGHT)

        # Settings
        frame_settings = tk.LabelFrame(self, text="Typing Settings")
        frame_settings.pack(pady=10, padx=10, fill='x')
        
        tk.Label(frame_settings, text="Excel Name Column Header:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_column = tk.Entry(frame_settings, width=40)
        self.entry_column.insert(0, "Name")
        self.entry_column.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(frame_settings, text="Wait Time (seconds) per name:\n(Time allowed for Teams to search)").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_delay = tk.Entry(frame_settings, width=10)
        self.entry_delay.insert(0, "2.5")
        self.entry_delay.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        # Auto-Add Settings
        frame_auto = tk.LabelFrame(self, text="Automation Actions")
        frame_auto.pack(pady=5, padx=10, fill='x')
        
        self.auto_add_var = tk.BooleanVar(value=True)
        chk_auto_add = tk.Checkbutton(frame_auto, text="Automatically press the 'Add' button when finished", variable=self.auto_add_var)
        chk_auto_add.grid(row=0, column=0, columnspan=2, sticky='w', padx=5, pady=5)
        
        tk.Label(frame_auto, text="Number of 'Tab' presses to reach Add button:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_tabs = tk.Entry(frame_auto, width=5)
        self.entry_tabs.insert(0, "1")
        self.entry_tabs.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        # Fail Safe Instructions
        frame_safe = tk.LabelFrame(self, text="⚠️ EMERGENCY STOP ⚠️", fg="red")
        frame_safe.pack(pady=10, padx=10, fill='x')
        tk.Label(frame_safe, text="If the robot types in the wrong place or goes out of control,\nSLAM YOUR MOUSE INTO ANY CORNER OF YOUR SCREEN\nto instantly abort the script.", fg="red", font=("Arial", 10, "bold")).pack(pady=10)

        # Log
        tk.Label(self, text="Process Log:").pack(anchor='w', padx=10)
        self.log_text = scrolledtext.ScrolledText(self, height=10, state='disabled')
        self.log_text.pack(padx=10, pady=5, fill='both', expand=True)
        
        # Action Button
        self.btn_run = tk.Button(self, text="START TYPING (5 Second Delay)", command=self.start_process, bg="green", fg="white", font=("Segoe UI", 12, "bold"))
        self.btn_run.pack(pady=10)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.update_idletasks()

    def browse_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.filepath:
            self.lbl_file.config(text=os.path.basename(self.filepath))

    def start_process(self):
        if not self.filepath:
            messagebox.showerror("Error", "Please select an Excel file.")
            return
            
        name_col = self.entry_column.get().strip()
        try:
            delay = float(self.entry_delay.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Wait time must be a number (e.g., 2.5).")
            return
            
        try:
            tabs_count = int(self.entry_tabs.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Number of tabs must be an integer.")
            return
            
        self.btn_run.config(state='disabled')
        self.log("Initializing...")
        
        # Run in a separate thread so UI doesn't freeze
        threading.Thread(target=self.run_process, args=(name_col, delay, self.auto_add_var.get(), tabs_count), daemon=True).start()

    def run_process(self, name_col, delay, auto_add, tabs_count):
        try:
            # 1. Read Excel
            self.log(f"Reading Excel file: {os.path.basename(self.filepath)}")
            try:
                if self.filepath.lower().endswith('.xls'):
                    try:
                        df = pd.read_excel(self.filepath, engine='xlrd')
                    except Exception as e:
                        # Fallback to HTML if it's an HTML file saved as .xls
                        self.log("xlrd failed, attempting to read as HTML table...")
                        dfs = pd.read_html(self.filepath)
                        df = dfs[0]
                else:
                    df = pd.read_excel(self.filepath, engine='openpyxl')
            except Exception as e:
                self.log(f"Failed to read Excel file: {e}")
                self.btn_run.config(state='normal')
                return
            
            if name_col not in df.columns:
                self.log(f"Error: Column '{name_col}' not found in the Excel file.")
                self.btn_run.config(state='normal')
                return

            names_to_type = []
            for index, row in df.iterrows():
                original_name = row[name_col]
                if pd.isna(original_name):
                    continue
                cleaned_name = clean_name(str(original_name))
                if cleaned_name:
                    names_to_type.append(cleaned_name)

            total_names = len(names_to_type)
            if total_names == 0:
                self.log("No valid names found in the Excel file.")
                self.btn_run.config(state='normal')
                return

            self.log(f"Found {total_names} names to type.")
            self.log("\n>>> YOU HAVE 5 SECONDS! <<<")
            self.log("Quickly click into the 'Add Member' search box in Teams...")
            
            # Countdown
            for i in range(5, 0, -1):
                self.log(f"Starting in {i}...")
                time.sleep(1)
                
            self.log("🚀 TAKING OVER KEYBOARD NOW! DO NOT TOUCH! 🚀")
            
            # Typing Loop
            for i, name in enumerate(names_to_type, 1):
                self.log(f"[{i}/{total_names}] Typing: {name}")
                
                # Type the name quickly
                pyautogui.write(name, interval=0.02)
                
                # Wait for Teams directory search to load the dropdown
                time.sleep(delay)
                
                # Press enter to select the top result
                pyautogui.press('enter')
                
                # Tiny wait before the next name to ensure Teams processes the selection
                time.sleep(0.5)
                
                # CLEAR the text box in case the name was NOT found.
                # If the name WAS found, the text box is empty and this does nothing.
                # If the name was NOT found, this selects the text and deletes it.
                # We use Shift+Home and Delete because Delete (forward delete) does not 
                # delete chips in Teams if the box is empty, whereas Backspace might.
                pyautogui.hotkey('shift', 'home')
                pyautogui.press('delete')
                time.sleep(0.2)
                
            # Auto Add logic
            if auto_add:
                self.log(f"\nTyping complete. Automatically pressing 'Add' (Tabbing {tabs_count} times)...")
                time.sleep(1.0) # Small pause before navigating
                for _ in range(tabs_count):
                    pyautogui.press('tab')
                    time.sleep(0.2)
                
                pyautogui.press('enter')
                self.log("✅ Pressed Add!")
                messagebox.showinfo("Done", "Typing complete and 'Add' was pressed automatically!")
            else:
                self.log("\n✅ ALL DONE!")
                self.log("Please review the selected names in Teams, and manually click the 'Add' button to finish.")
                messagebox.showinfo("Done", "Typing complete! Please review the names in Teams and click Add.")

        except pyautogui.FailSafeException:
            self.log("\n❌ EMERGENCY STOP TRIGGERED! ❌")
            self.log("You moved the mouse to the corner of the screen. Script aborted.")
            messagebox.showwarning("Aborted", "Emergency Stop Triggered. Process aborted.")
            
        except Exception as e:
            self.log(f"An unexpected error occurred: {str(e)}")
            
        finally:
            self.btn_run.config(state='normal')

if __name__ == "__main__":
    app = TeamsBruteApp()
    app.mainloop()
