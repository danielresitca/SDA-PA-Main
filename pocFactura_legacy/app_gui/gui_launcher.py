import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import threading
import sys


def run_standardization(xml_path):
    if not xml_path:
        messagebox.showerror("Error", "Please select an XML file first.")
        return

    def worker():
        # Get the directory where this script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(script_dir)  # Go up one level from app_gui

        # Get Python executable from virtual environment
        if sys.platform == "win32":
            python_exe = os.path.join(project_root, ".venv", "Scripts", "python.exe")
        else:
            python_exe = os.path.join(project_root, ".venv", "bin", "python")

        # Run the CLI script directly with Python - using absolute path
        cli_script = os.path.join(project_root, "src", "invoice_core", "cli_standardize.py")

        # Check if files exist
        if not os.path.exists(python_exe):
            messagebox.showerror("Error",
                                 f"Python not found at:\n{python_exe}\n\nPlease create virtual environment:\npython -m venv .venv")
            return

        if not os.path.exists(cli_script):
            messagebox.showerror("Error", f"CLI script not found at:\n{cli_script}")
            return

        # Output files in project root
        out_lines = os.path.join(project_root, "lines_raw.json")
        out_standard = os.path.join(project_root, "lines_standardized.json")
        out_csv = os.path.join(project_root, "lines_standardized.csv")
        codes_xlsx = os.path.join(project_root, "cod_vamal.xlsx")

        cmd = [
            python_exe,
            cli_script,
            "--xml", xml_path,
            "--codes-xlsx", codes_xlsx,
            "--out-lines", out_lines,
            "--out-standard", out_standard,
            "--out-csv", out_csv,
            "--min-score", "0.18"
        ]
        try:
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            messagebox.showinfo("Done", "✅ Processing complete!\nResults saved in your project folder.")
        except subprocess.CalledProcessError as e:
            error_msg = f"Processing failed!\n\nCommand: {' '.join(cmd)}\n\nError output:\n{e.stderr}\n\nOutput:\n{e.stdout}"
            messagebox.showerror("Error", error_msg)
        except FileNotFoundError as e:
            messagebox.showerror("Error",
                                 f"File not found:\n{e}\n\nMake sure your virtual environment exists and dependencies are installed.")

    threading.Thread(target=worker, daemon=True).start()  # Run in separate thread


def select_file():
    file_path = filedialog.askopenfilename(
        title="Select XML Invoice File",
        filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
    )
    if file_path:
        xml_var.set(file_path)


# --- GUI setup ---
root = tk.Tk()
root.title("Invoice Standardizer")
root.geometry("460x220")
root.resizable(False, False)

tk.Label(root, text="Select your XML invoice file:", font=("Arial", 12)).pack(pady=15)

xml_var = tk.StringVar()
entry = tk.Entry(root, textvariable=xml_var, width=55)
entry.pack(pady=5)

tk.Button(root, text="Browse", command=select_file, width=15).pack(pady=5)
tk.Button(root, text="Run Standardization", command=lambda: run_standardization(xml_var.get()), width=25).pack(pady=15)

tk.Label(root, text="Results: lines_raw.json, lines_standardized.json, lines_standardized.csv",
         font=("Arial", 9), fg="gray").pack(pady=5)

root.mainloop()