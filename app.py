import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import subprocess
import threading
import os
from datetime import datetime

# --- Backend Execution ---
def run_backend():
    symbol = symbol_entry.get().strip().upper()
    country = country_entry.get().strip().upper()
    portfolio = portfolio_combo.get().strip()
    start = start_date_entry.get_date().strftime("%Y%m%d")
    end = end_date_entry.get_date().strftime("%Y%m%d")
    stock_option = "Y" if stock_var.get() else "N"

    if not symbol or not country or not start or not end or not portfolio:
        messagebox.showerror("Input Error", "Please fill in all fields.")
        return

    progress_bar.grid(row=6, column=1, pady=10)
    progress_bar.start()

    def run():
        try:
            subprocess.run([
                "python",
                "main.py",       # chnage this string to the location of main.py
                "--symbol", symbol,
                "--country", country,
                "--instrument", instrument_var.get().strip(),
                "--start", start,
                "--end", end,
                "--stock", stock_option,
                "--portfolio", portfolio
            ], check=True)


            progress_bar.stop()
            progress_bar.grid_remove()

        except subprocess.CalledProcessError as e:
            progress_bar.stop()
            progress_bar.grid_remove()
            messagebox.showerror("Script Error", f"Failed to run backend:\n{str(e)}")

    threading.Thread(target=run).start()

# --- UI Setup ---
root = tk.Tk()
root.title("Bloomberg Analyzer")
root.attributes("-fullscreen", True)
root.configure(bg="#f2f2f2")

tk.Label(root, text="Bloomberg Analyzer", font=("Segoe UI", 26, "bold"), bg="#f2f2f2", fg="#333").pack(pady=30)
form_frame = tk.Frame(root, bg="#f2f2f2")
form_frame.pack()

def style_entry(entry):
    entry.config(font=("Segoe UI", 14), width=40, relief="groove", bd=2)

def style_button(btn, bg, fg):
    btn.config(font=("Segoe UI", 13, "bold"), bg=bg, fg=fg,
               relief="flat", padx=20, pady=10, bd=0)

# --- Inputs ---
tk.Label(form_frame, text="Company Symbol (e.g. HBM)", bg="#f2f2f2", font=("Segoe UI", 14)).grid(row=0, column=0, sticky="w", pady=10)
symbol_entry = tk.Entry(form_frame)
style_entry(symbol_entry)
symbol_entry.grid(row=0, column=1, padx=20)

tk.Label(form_frame, text="Underlying Country (e.g. CA)", bg="#f2f2f2", font=("Segoe UI", 14)).grid(row=1, column=0, sticky="w", pady=10)
country_entry = tk.Entry(form_frame)
style_entry(country_entry)
country_entry.grid(row=1, column=1, padx=20)


# --- Instrument Dropdown ---
tk.Label(form_frame, text="Instrument Type", bg="#f2f2f2", font=("Segoe UI", 14)).grid(row=2, column=0, sticky="w", pady=10)
instrument_var = tk.StringVar()
instrument_combo = ttk.Combobox(form_frame, textvariable=instrument_var, values=["Stock", "Index"], font=("Segoe UI", 14), width=37, state="readonly")
instrument_combo.grid(row=2, column=1, padx=20)
instrument_combo.set("Stock")

# --- Portfolio Dropdown ---
tk.Label(form_frame, text="Select Portfolio", bg="#f2f2f2", font=("Segoe UI", 14)).grid(row=2, column=0, sticky="w", pady=10)
portfolio_list = [
    "2561941 Ontario Inc",
    "Amplus Credit Income",
    "Amolino Holdings Inc",
    "Balsillie Family",
    "CANDESTRA HOLDINGS",
    "DK Holdings Ltd",
    "FORTIS INVESTMENTS",
    "Kavelman-Fon",
    "Lions Bay Fund",
    "Voyager Fund",
    "Jamie",
    "Mike",
    "Vektorium Master",
    "ALL"
]
portfolio_combo = ttk.Combobox(form_frame, values=portfolio_list, font=("Segoe UI", 14), width=37, state="readonly")
portfolio_combo.grid(row=2, column=1, padx=20)
portfolio_combo.set("Voyager Fund")

# --- Dates ---
tk.Label(form_frame, text="Start Date", bg="#f2f2f2", font=("Segoe UI", 14)).grid(row=3, column=0, sticky="w", pady=10)
start_date_entry = DateEntry(form_frame, width=37, font=("Segoe UI", 14), date_pattern='yyyy-mm-dd',
                             background='darkblue', foreground='white', borderwidth=2, state='readonly')
start_date_entry.set_date(datetime.today())
start_date_entry.grid(row=3, column=1, padx=20)

tk.Label(form_frame, text="End Date", bg="#f2f2f2", font=("Segoe UI", 14)).grid(row=4, column=0, sticky="w", pady=10)
end_date_entry = DateEntry(form_frame, width=37, font=("Segoe UI", 14), date_pattern='yyyy-mm-dd',
                           background='darkblue', foreground='white', borderwidth=2, state='readonly')
end_date_entry.set_date(datetime.today())
end_date_entry.grid(row=4, column=1, padx=20)

stock_var = tk.IntVar()
tk.Checkbutton(form_frame, text="Include Stock Gain & Yield", variable=stock_var, bg="#f2f2f2",
               font=("Segoe UI", 13)).grid(row=5, column=1, sticky="w", pady=15)

# --- Progress Bar ---
progress_bar = ttk.Progressbar(form_frame, mode="indeterminate", length=300, style="green.Horizontal.TProgressbar")
progress_bar.grid(row=6, column=1, pady=10)
progress_bar.grid_remove()

style = ttk.Style()
style.theme_use('default')
style.configure("green.Horizontal.TProgressbar", troughcolor="#ccc", background="green", thickness=20)

# --- Button ---
button_frame = tk.Frame(root, bg="#f2f2f2")
button_frame.pack(pady=40)

generate_btn = tk.Button(button_frame, text="Generate Excel", command=run_backend)
style_button(generate_btn, "#4682B4", "white")
generate_btn.grid(row=0, column=0, padx=30)

exit_btn = tk.Button(root, text="Exit", command=root.quit)
style_button(exit_btn, "#CEAA70", "black")
exit_btn.pack(side="bottom", pady=20)

root.mainloop()
