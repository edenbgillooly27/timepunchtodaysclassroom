#!/usr/bin/env python
import sys
import os
import json
import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import platform
import subprocess
import webbrowser
from pathlib import Path

# Only attempt auto–installation if NOT running as a frozen exe.
if not getattr(sys, 'frozen', False):
    try:
        from fpdf import FPDF
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "fpdf"])
        from fpdf import FPDF
    try:
        from PIL import Image, ImageTk
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
        from PIL import Image, ImageTk
else:
    try:
        from fpdf import FPDF
    except ImportError:
        raise ImportError("Missing 'fpdf' module. Please install before building or running the exe.")
    try:
        from PIL import Image, ImageTk
    except ImportError:
        raise ImportError("Missing 'Pillow' module. Please install before building or running the exe.")

# -------------------------------
# Configuration and Constants
# -------------------------------
DATA_FILE = "data.json"
FIRST_PAY_PERIOD_START = datetime.date(2025, 2, 17)  # Monday
FIRST_PAY_PERIOD_END = datetime.date(2025, 2, 28)    # Friday of second week

MAIN_COLOR = "#213b97"

# -------------------------------
# Helper: Generate Time Options (6:00 AM to 5:00 PM)
# -------------------------------
def generate_am_pm_time_options():
    options = []
    start = datetime.datetime(2023, 1, 1, 6, 0)  # 6:00 AM
    end = datetime.datetime(2023, 1, 1, 17, 0)  # 5:00 PM
    current = start
    while current <= end:
        time_str = current.strftime("%I:%M %p").lstrip("0")
        options.append(time_str)
        current += datetime.timedelta(minutes=15)
    return options

TIME_OPTIONS = generate_am_pm_time_options()

def parse_am_pm_time(t_str):
    try:
        return datetime.datetime.strptime(t_str, "%I:%M %p").time()
    except Exception:
        return None

# -------------------------------
# Helpers for Window Centering
# -------------------------------
def center_window_on_parent(child, parent):
    child.update_idletasks()
    w = child.winfo_width()
    h = child.winfo_height()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    x = px + (pw // 2) - (w // 2)
    y = py + (ph // 2) - (h // 2)
    child.geometry(f"{w}x{h}+{x}+{y}")

def center_window_on_screen(win):
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw // 2) - (w // 2)
    y = (sh // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

# -------------------------------
# Main Application Class
# -------------------------------
class TimePunchApp:
    def __init__(self, master):
        self.master = master

        # Configure TTK style
        self.style = ttk.Style(master)
        self.style.configure(
            "TNotebook.Tab",
            font=("Arial", 16, "bold"),
            padding=(15, 10),
            foreground=MAIN_COLOR,
            background="SystemButtonFace"
        )
        self.style.map("TNotebook.Tab", [])
        self.style.configure(
            "Treeview",
            font=("Arial", 14)
        )
        self.style.configure(
            "Treeview.Heading",
            font=("Arial", 14, "bold")
        )

        master.geometry("650x350")
        center_window_on_screen(master)
        master.resizable(True, True)
        self.master.title("Time Punch System")
        self.master.configure(bg="white")
        try:
            self.master.iconbitmap("myicon.ico")
        except Exception as e:
            print("Error setting icon:", e)

        self.username = None
        self.data = self.load_data()
        self.current_pay_period = None
        self.today = datetime.date.today()

        self.final_day_enabled = False  # for toggle button state
        self.login_frame = tk.Frame(master, bg="white")
        self.login_frame.pack(fill="both", expand=True)
        self.notebook = None
        self.create_login_ui()

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    return json.load(f)
            except Exception as e:
                messagebox.showerror("Error", f"Error reading data file: {e}")
                return {"users": {}}
        else:
            return {"users": {}}

    def save_data(self):
        try:
            with open(DATA_FILE, "w") as f:
                json.dump(self.data, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data: {e}")

    def create_login_ui(self):
        lbl = tk.Label(
            self.login_frame,
            text="Select or enter your name:",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        lbl.pack(pady=30)
        saved_names = list(self.data.get("users", {}).keys())
        self.name_var = tk.StringVar()
        self.name_combobox = ttk.Combobox(
            self.login_frame,
            textvariable=self.name_var,
            values=saved_names,
            font=("Arial", 16)
        )
        self.name_combobox.pack(pady=5)
        btn = tk.Button(
            self.login_frame,
            text="Login",
            command=self.login,
            font=("Arial", 16, "bold"),
            bd=4,
            relief="raised",
            width=10,
            bg=MAIN_COLOR,
            fg="white"
        )
        btn.pack(pady=20)

    def login(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Input Required", "Please enter your name.")
            return
        self.username = name
        if self.username not in self.data["users"]:
            self.data["users"][self.username] = {"pay_periods": []}
            self.save_data()
        self.initialize_current_pay_period()
        self.login_frame.pack_forget()
        self.master.geometry("900x635")
        center_window_on_screen(self.master)
        self.notebook = ttk.Notebook(self.master, style="TNotebook")
        self.notebook.pack(fill="both", expand=True)

        # Create tabs:
        self.create_punch_in_out_tab()
        self.create_weeks_punches_tab()  # Updated tab with scrollable area and full-width days
        self.create_past_pay_periods_tab()

        # Bind tab change (update week's punches when selected)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.update_ui()

    def on_tab_changed(self, event):
        selected_tab = event.widget.tab(event.widget.index("current"))["text"]
        if selected_tab == "Week's Punches":
            self.populate_weeks_punches()

    def initialize_current_pay_period(self):
        user_periods = self.data["users"][self.username]["pay_periods"]
        if not user_periods or user_periods[-1].get("finalized"):
            if user_periods:
                last_end = datetime.date.fromisoformat(user_periods[-1]["end_date"])
                new_start = last_end + datetime.timedelta(days=3)  # e.g. Friday -> Monday
            else:
                new_start = FIRST_PAY_PERIOD_START
            new_end = new_start + datetime.timedelta(days=11)  # 2 weeks total
            period = {
                "start_date": new_start.isoformat(),
                "end_date": new_end.isoformat(),
                "records": {},
                "custom_hours": None,
                "finalized": False
            }
            user_periods.append(period)
            self.save_data()
            self.current_pay_period = period
        else:
            self.current_pay_period = user_periods[-1]

    # -------------------------------
    # TAB 1: Punch In/Out (with Daily Punches Summary Above Last Punch)
    # -------------------------------
    def create_punch_in_out_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Punch In/Out")
        header_frame = tk.Frame(tab)
        header_frame.pack(pady=10, fill="x")
        tk.Label(
            header_frame,
            text="Today's Classroom",
            font=("Arial", 18, "bold"),
            fg=MAIN_COLOR
        ).pack(side="top", anchor="w", padx=5)
        self.welcome_label = tk.Label(
            header_frame,
            text=f"Welcome, {self.username}",
            font=("Arial", 16),
            fg=MAIN_COLOR
        )
        self.welcome_label.pack(side="left", padx=5)
        self.payperiod_label = tk.Label(
            header_frame,
            font=("Arial", 16),
            fg=MAIN_COLOR
        )
        self.payperiod_label.pack(side="right", padx=5)
        info_frame = tk.Frame(tab, bd=2, relief="groove", bg="white")
        info_frame.pack(pady=5, fill="x")
        # New daily punches line above "Last Punch"
        self.daily_punches_line = tk.Label(
            info_frame,
            text="Daily Punches: None",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        self.daily_punches_line.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.last_punch_label = tk.Label(
            info_frame,
            text="Last Punch: None",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        self.last_punch_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.daily_total_label = tk.Label(
            info_frame,
            text="Today's Hours: 0.00",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        self.daily_total_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.period_total_label = tk.Label(
            info_frame,
            text="Pay Period Total Hours: 0.00",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        self.period_total_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        button_frame = tk.Frame(tab)
        button_frame.pack(pady=15)
        self.punch_in_button = tk.Button(
            button_frame,
            text="Punch In",
            command=self.punch_in,
            font=("Arial", 16, "bold"),
            width=10,
            bd=4,
            relief="raised",
            bg=MAIN_COLOR,
            fg="white"
        )
        self.punch_in_button.grid(row=0, column=0, padx=10)
        self.punch_out_button = tk.Button(
            button_frame,
            text="Punch Out",
            command=self.punch_out,
            font=("Arial", 16, "bold"),
            width=10,
            bd=4,
            relief="raised",
            bg=MAIN_COLOR,
            fg="white"
        )
        self.punch_out_button.grid(row=0, column=1, padx=10)
        self.final_day_button = tk.Button(
            tab,
            text="Final Pay Period Day",
            command=self.toggle_final_day_state,
            font=("Arial", 16, "bold"),
            bd=4,
            relief="raised",
            width=18,
            bg=MAIN_COLOR,
            fg="white"
        )
        self.final_day_button.pack(pady=10)
        self.final_day_frame = tk.Frame(tab, bg="white", bd=2, relief="groove")
        self.final_pay_in1_var = tk.StringVar()
        self.final_pay_out1_var = tk.StringVar()
        self.final_pay_in2_var = tk.StringVar()
        self.final_pay_out2_var = tk.StringVar()
        pair1_frame = tk.Frame(self.final_day_frame, bg="white")
        pair1_frame.pack(pady=5)
        tk.Label(
            pair1_frame,
            text="Punch In 1:",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        ).grid(row=0, column=0, padx=10, pady=5)
        self.final_pay_in1 = ttk.Combobox(
            pair1_frame,
            textvariable=self.final_pay_in1_var,
            values=TIME_OPTIONS,
            font=("Arial", 16),
            state="readonly"
        )
        self.final_pay_in1.grid(row=0, column=1, padx=10, pady=5)
        tk.Label(
            pair1_frame,
            text="Punch Out 1:",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        ).grid(row=0, column=2, padx=10, pady=5)
        self.final_pay_out1 = ttk.Combobox(
            pair1_frame,
            textvariable=self.final_pay_out1_var,
            values=TIME_OPTIONS,
            font=("Arial", 16),
            state="readonly"
        )
        self.final_pay_out1.grid(row=0, column=3, padx=10, pady=5)
        pair2_frame = tk.Frame(self.final_day_frame, bg="white")
        pair2_frame.pack(pady=5)
        tk.Label(
            pair2_frame,
            text="Punch In 2:",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        ).grid(row=0, column=0, padx=10, pady=5)
        self.final_pay_in2 = ttk.Combobox(
            pair2_frame,
            textvariable=self.final_pay_in2_var,
            values=TIME_OPTIONS,
            font=("Arial", 16),
            state="readonly"
        )
        self.final_pay_in2.grid(row=0, column=1, padx=10, pady=5)
        tk.Label(
            pair2_frame,
            text="Punch Out 2:",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        ).grid(row=0, column=2, padx=10, pady=5)
        self.final_pay_out2 = ttk.Combobox(
            pair2_frame,
            textvariable=self.final_pay_out2_var,
            values=TIME_OPTIONS,
            font=("Arial", 16),
            state="readonly"
        )
        self.final_pay_out2.grid(row=0, column=3, padx=10, pady=5)
        self.finalize_button = tk.Button(
            self.final_day_frame,
            text="Finalize Hours",
            command=self.finalize_hours,
            font=("Arial", 16, "bold"),
            bd=4,
            relief="raised",
            state="disabled",
            bg=MAIN_COLOR,
            fg="white"
        )
        self.finalize_button.pack(pady=10)
        self.final_pay_in1_var.trace("w", self.update_final_pay_out1_options)
        self.final_pay_out1_var.trace("w", self.update_final_pay_in2_options)
        self.final_pay_in2_var.trace("w", self.update_final_pay_out2_options)
        self.final_pay_in1_var.trace("w", lambda *args: self.update_finalize_button_state())
        self.final_pay_out1_var.trace("w", lambda *args: self.update_finalize_button_state())

    def toggle_final_day_state(self):
        self.final_day_enabled = not self.final_day_enabled
        if self.final_day_enabled:
            self.final_day_frame.pack(pady=10, fill="x")
            self.final_day_button.config(relief="sunken")
        else:
            self.final_day_frame.pack_forget()
            self.final_day_button.config(relief="raised")

    def update_finalize_button_state(self):
        if self.final_pay_in1_var.get() and self.final_pay_out1_var.get():
            self.finalize_button.config(state="normal")
        else:
            self.finalize_button.config(state="disabled")

    def update_final_pay_out1_options(self, *args):
        selected_in1 = self.final_pay_in1_var.get()
        if selected_in1:
            base_time = parse_am_pm_time(selected_in1)
            new_options = [t for t in TIME_OPTIONS if parse_am_pm_time(t) > base_time]
            self.final_pay_out1.config(values=new_options)
        else:
            self.final_pay_out1.config(values=TIME_OPTIONS)

    def update_final_pay_in2_options(self, *args):
        selected_out1 = self.final_pay_out1_var.get()
        if selected_out1:
            base_time = parse_am_pm_time(selected_out1)
            new_options = [t for t in TIME_OPTIONS if parse_am_pm_time(t) > base_time]
            self.final_pay_in2.config(values=new_options)
        else:
            self.final_pay_in2.config(values=TIME_OPTIONS)

    def update_final_pay_out2_options(self, *args):
        selected_in2 = self.final_pay_in2_var.get()
        if selected_in2:
            base_time = parse_am_pm_time(selected_in2)
            new_options = [t for t in TIME_OPTIONS if parse_am_pm_time(t) > base_time]
            self.final_pay_out2.config(values=new_options)
        else:
            self.final_pay_out2.config(values=TIME_OPTIONS)

    # -------------------------------
    # TAB 2: Week's Punches (Scrollable, Full-Width Days, Mouse Wheel Support)
    # -------------------------------
    def create_weeks_punches_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Week's Punches")
        self.weeks_punches_label = tk.Label(
            tab,
            text="Punches for the current week",
            font=("Arial", 18, "bold"),
            fg=MAIN_COLOR
        )
        self.weeks_punches_label.pack(pady=10)
        # Create a canvas and attach a vertical scrollbar
        self.weeks_canvas = tk.Canvas(tab, bg="white")
        self.weeks_canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.weeks_canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.weeks_canvas.configure(yscrollcommand=scrollbar.set)
        # Bind mouse wheel scrolling (Windows/Mac and Linux)
        self.weeks_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.weeks_canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.weeks_canvas.bind_all("<Button-5>", self._on_mousewheel)
        # Create an internal frame inside the canvas
        self.punches_list_frame = tk.Frame(self.weeks_canvas, bg="white")
        # Place the frame in the canvas with full width
        self.canvas_window = self.weeks_canvas.create_window((0, 0), window=self.punches_list_frame, anchor="nw", width=self.weeks_canvas.winfo_width())
        self.punches_list_frame.bind("<Configure>", self._on_frame_configure)
        self.weeks_canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_frame_configure(self, event):
        self.weeks_canvas.configure(scrollregion=self.weeks_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # Ensure the internal frame width always matches the canvas width
        self.weeks_canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        if event.delta:
            self.weeks_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            self.weeks_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.weeks_canvas.yview_scroll(1, "units")

    def populate_weeks_punches(self):
        # Clear previous punches
        for widget in self.punches_list_frame.winfo_children():
            widget.destroy()
        # Determine current week (Monday to Sunday)
        monday = self.today - datetime.timedelta(days=self.today.weekday())
        days = [monday + datetime.timedelta(days=i) for i in range(7)]
        for day in days:
            day_str = day.isoformat()
            day_header = tk.Label(
                self.punches_list_frame,
                text=day_str,
                font=("Arial", 16, "bold"),
                fg=MAIN_COLOR,
                bg="white"
            )
            day_header.pack(fill="x", anchor="w", pady=(10, 2))
            records = self.current_pay_period.get("records", {}).get(day_str, [])
            if not records:
                tk.Label(
                    self.punches_list_frame,
                    text="No punches recorded.",
                    font=("Arial", 16),
                    fg=MAIN_COLOR,
                    bg="white"
                ).pack(fill="x", anchor="w", padx=10, pady=2)
            else:
                for idx, cycle in enumerate(records, start=1):
                    if "punch_in_1" in cycle:
                        text = (
                            f"Record (Finalized): In 1: {cycle.get('punch_in_1')} | "
                            f"Out 1: {cycle.get('punch_out_1')} | "
                            f"In 2: {cycle.get('punch_in_2')} | "
                            f"Out 2: {cycle.get('punch_out_2')} | "
                            f"Hours: {cycle.get('duration', 0.0):.2f}"
                        )
                    else:
                        text = (
                            f"Cycle {idx}: In: {cycle.get('punch_in', 'N/A')} | "
                            f"Out: {cycle.get('punch_out', 'N/A')} | "
                            f"Hours: {cycle.get('duration', 0.0):.2f}"
                        )
                    tk.Label(
                        self.punches_list_frame,
                        text=text,
                        font=("Arial", 16),
                        fg=MAIN_COLOR,
                        bg="white"
                    ).pack(fill="x", anchor="w", padx=20, pady=2)
            # Bold horizontal separator for each day
            separator = tk.Frame(self.punches_list_frame, height=2, bd=2, relief="ridge", bg=MAIN_COLOR)
            separator.pack(fill="x", pady=5)

    # -------------------------------
    # TAB 3: Past Pay Periods (Centered Table Columns)
    # -------------------------------
    def create_past_pay_periods_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Past Pay Periods")
        tree_frame = tk.Frame(tab, bg="white")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        columns = ("start", "end", "total")
        self.history_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8, style="Treeview")
        for col, txt in zip(columns, ["Start Date", "End Date", "Total Hours"]):
            self.history_tree.heading(col, text=txt)
            self.history_tree.column(col, width=120, anchor="center")  # Center each column
        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.history_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        detail_frame = tk.Frame(tab, bg="white", bd=2, relief="groove")
        detail_frame.pack(fill="x", padx=10, pady=5)
        self.detail_label_period = tk.Label(
            detail_frame,
            text="Pay Period: ",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        self.detail_label_period.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.detail_label_hours = tk.Label(
            detail_frame,
            text="Total Hours: ",
            font=("Arial", 16),
            fg=MAIN_COLOR,
            bg="white"
        )
        self.detail_label_hours.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        pdf_button_frame = tk.Frame(tab)
        pdf_button_frame.pack(pady=10)
        open_folder_button = tk.Button(
            pdf_button_frame,
            text="Open PDF Folder",
            command=self.open_pdf_folder_from_history,
            font=("Arial", 16, "bold"),
            bd=4,
            relief="raised",
            width=14,
            bg=MAIN_COLOR,
            fg="white"
        )
        open_folder_button.pack(side="left", padx=5)
        self.history_tree.bind("<Button-1>", self.on_single_click_period, add=True)
        self.history_tree.bind("<<TreeviewSelect>>", self.on_select_period)
        self.history_tree.bind("<Double-1>", self.on_double_click_period)
        self.populate_past_pay_periods()

    def on_single_click_period(self, event):
        rowid = self.history_tree.identify_row(event.y)
        if rowid:
            if rowid in self.history_tree.selection():
                self.history_tree.selection_remove(rowid)
                return "break"

    def populate_past_pay_periods(self):
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)
        user_periods = self.data["users"][self.username]["pay_periods"]
        finalized_list = []
        for idx, period in enumerate(user_periods):
            if period.get("finalized"):
                finalized_list.append((idx, period))
        finalized_list.sort(key=lambda item: item[1]["end_date"], reverse=True)
        self.finalized_indices = []
        for _, period in finalized_list:
            original_index = user_periods.index(period)
            self.finalized_indices.append(original_index)
            start = period["start_date"]
            end = period["end_date"]
            total = 0.0
            for day in period.get("records", {}):
                for cycle in period["records"][day]:
                    total += cycle.get("duration", 0.0)
            if period.get("custom_hours") is not None:
                total = float(period["custom_hours"])
            self.history_tree.insert(
                "",
                "end",
                iid=str(len(self.finalized_indices) - 1),
                values=(start, end, f"{total:.2f}")
            )

    def on_select_period(self, event):
        sel = self.history_tree.selection()
        if not sel:
            self.detail_label_period.config(text="Pay Period: ")
            self.detail_label_hours.config(text="Total Hours: ")
            return
        item = self.history_tree.item(sel)
        vals = item["values"]
        if len(vals) == 3:
            start, end, total = vals
            self.detail_label_period.config(text=f"Pay Period: {start} to {end}")
            self.detail_label_hours.config(text=f"Total Hours: {total}")

    def on_double_click_period(self, event):
        sel = self.history_tree.selection()
        if not sel:
            return
        tree_idx = int(sel[0])
        period_idx = self.finalized_indices[tree_idx]
        period = self.data["users"][self.username]["pay_periods"][period_idx]
        desktop_folder = Path.home() / "Desktop"
        folder = desktop_folder / "Hours" / self.username
        pdf_filename = f"{self.username} Pay Period {period['end_date']}.pdf"
        pdf_filepath = folder / pdf_filename
        self.open_pdf(str(pdf_filepath))

    def open_pdf_folder_from_history(self):
        sel = self.history_tree.selection()
        if not sel:
            messagebox.showwarning("Selection Required", "Please select a pay period.")
            return
        desktop_folder = Path.home() / "Desktop"
        folder = desktop_folder / "Hours" / self.username
        try:
            if platform.system() == "Windows":
                os.startfile(str(folder))
            elif platform.system() == "Darwin":
                subprocess.call(["open", str(folder)])
            else:
                subprocess.call(["xdg-open", str(folder)])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")

    # -------------------------------
    # Punch In/Out Logic
    # -------------------------------
    def punch_in(self):
        if self.today.weekday() >= 5:
            messagebox.showwarning("Unavailable", "Time punches are not allowed on weekends.")
            return
        today_str = self.today.isoformat()
        now_str = datetime.datetime.now().strftime("%I:%M %p").lstrip("0")
        records = self.current_pay_period.setdefault("records", {})
        cycles = records.setdefault(today_str, [])
        if cycles and ("punch_in" in cycles[-1] and "punch_out" not in cycles[-1]):
            messagebox.showwarning("Sequence Error", "You must punch out before punching in again.")
            return
        if len(cycles) >= 2:
            messagebox.showwarning("Limit Reached", "Maximum two punch cycles per day reached.")
            return
        cycles.append({"punch_in": now_str})
        self.save_data()
        self.update_ui()
        messagebox.showinfo("Punched In", "Thank you! You are now Punched In.")

    def punch_out(self):
        today_str = self.today.isoformat()
        records = self.current_pay_period.get("records", {})
        cycles = records.get(today_str, [])
        if not cycles or ("punch_in" in cycles[-1] and "punch_out" in cycles[-1]):
            messagebox.showwarning("Sequence Error", "No active punch in to punch out from.")
            return
        now = datetime.datetime.now()
        now_str = now.strftime("%I:%M %p").lstrip("0")
        punch_in_str = cycles[-1]["punch_in"]
        try:
            in_time = datetime.datetime.strptime(punch_in_str, "%I:%M %p")
            in_time = now.replace(hour=in_time.hour, minute=in_time.minute, second=0, microsecond=0)
            duration = (now - in_time).total_seconds() / 3600.0
        except Exception:
            duration = 0.0
        cycles[-1]["punch_out"] = now_str
        cycles[-1]["duration"] = duration if duration > 0 else 0.0
        self.save_data()
        self.update_ui()
        messagebox.showinfo("Punched Out", "Thank you! You are now punched out.")

    # -------------------------------
    # Finalize Hours
    # -------------------------------
    def finalize_hours(self):
        final_in1 = self.final_pay_in1_var.get()
        final_out1 = self.final_pay_out1_var.get()
        final_in2 = self.final_pay_in2_var.get()
        final_out2 = self.final_pay_out2_var.get()
        if not (final_in1 and final_out1):
            messagebox.showwarning("Input Required",
                                   "Please select Punch In 1 and Punch Out 1 for the final pay period day.")
            return
        if (bool(final_in2) != bool(final_out2)):
            messagebox.showerror("Input Error",
                                 "Please provide both Punch In 2 and Punch Out 2, or leave both empty.")
            return
        try:
            in1_time_obj = datetime.datetime.strptime(final_in1, "%I:%M %p")
            out1_time_obj = datetime.datetime.strptime(final_out1, "%I:%M %p")
        except Exception as e:
            messagebox.showerror("Error", f"Time parsing error for Punch In/Out 1: {e}")
            return
        if out1_time_obj <= in1_time_obj:
            messagebox.showerror("Time Error", "Punch Out 1 must be after Punch In 1.")
            return
        duration1 = (out1_time_obj - in1_time_obj).total_seconds() / 3600.0
        duration2 = 0.0
        if final_in2 and final_out2:
            try:
                in2_time_obj = datetime.datetime.strptime(final_in2, "%I:%M %p")
                out2_time_obj = datetime.datetime.strptime(final_out2, "%I:%M %p")
            except Exception as e:
                messagebox.showerror("Error", f"Time parsing error for Punch In/Out 2: {e}")
                return
            if out2_time_obj <= in2_time_obj:
                messagebox.showerror("Time Error", "Punch Out 2 must be after Punch In 2.")
                return
            duration2 = (out2_time_obj - in2_time_obj).total_seconds() / 3600.0
        total_duration = duration1 + duration2
        final_day_key = self.current_pay_period["end_date"]
        record = {
            "punch_in_1": final_in1,
            "punch_out_1": final_out1,
            "punch_in_2": final_in2 if final_in2 else "",
            "punch_out_2": final_out2 if final_out2 else "",
            "duration": total_duration
        }
        self.current_pay_period.setdefault("records", {})[final_day_key] = [record]
        self.current_pay_period["finalized"] = True
        self.save_data()
        self.update_ui()
        desktop_folder = Path.home() / "Desktop"
        folder = desktop_folder / "Hours" / self.username
        folder.mkdir(parents=True, exist_ok=True)
        pdf_filename = f"{self.username} Pay Period {self.current_pay_period['end_date']}.pdf"
        pdf_filepath = folder / pdf_filename
        self.generate_pdf(str(pdf_filepath))
        self.show_finalize_options_popup(str(pdf_filepath))
        self.populate_past_pay_periods()
        self.initialize_current_pay_period()
        self.update_ui()
        self.final_pay_in1_var.set("")
        self.final_pay_out1_var.set("")
        self.final_pay_in2_var.set("")
        self.final_pay_out2_var.set("")
        self.final_day_enabled = False
        self.final_day_button.config(relief="raised")
        self.final_day_frame.pack_forget()

    def generate_pdf(self, pdf_filepath):
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 18)
        pdf.set_text_color(33, 59, 151)
        pdf.cell(0, 10, "Today's Classroom", ln=True, align="C")
        pdf.set_font("Arial", "", 14)
        pdf.ln(5)
        pdf.cell(0, 10, f"Employee: {self.username}", ln=True, align="C")
        pdf.cell(
            0,
            10,
            f"Pay Period: {self.current_pay_period['start_date']} to {self.current_pay_period['end_date']}",
            ln=True,
            align="C"
        )
        pdf.ln(5)
        pdf.set_font("Arial", "B", 14)
        widths = [35, 30, 30, 30, 30, 40]
        headers = ["Date", "In 1", "Out 1", "In 2", "Out 2", "Day Total"]
        for w, h in zip(widths, headers):
            pdf.cell(w, 10, h, border=1, align="C")
        pdf.ln()
        pdf.set_font("Arial", "", 12)
        start_date = datetime.date.fromisoformat(self.current_pay_period["start_date"])
        end_date = datetime.date.fromisoformat(self.current_pay_period["end_date"])
        total_pay_period_hours = 0.0
        weekly_hours = {}
        current_day = start_date
        while current_day <= end_date:
            day_str = current_day.isoformat()
            if current_day.weekday() >= 5:
                pdf.set_fill_color(200, 200, 200)
                for w in widths:
                    pdf.cell(w, 10, "", border=1, align="C", fill=True)
                pdf.ln()
            else:
                records = self.current_pay_period.get("records", {}).get(day_str, [])
                punch_in_1 = ""
                punch_out_1 = ""
                punch_in_2 = ""
                punch_out_2 = ""
                day_hours = 0.0
                if records:
                    rec = records[0]
                    if "punch_in_1" in rec:
                        punch_in_1 = rec.get("punch_in_1", "")
                        punch_out_1 = rec.get("punch_out_1", "")
                        punch_in_2 = rec.get("punch_in_2", "")
                        punch_out_2 = rec.get("punch_out_2", "")
                        day_hours = rec.get("duration", 0.0)
                    else:
                        if len(records) >= 1:
                            c1 = records[0]
                            punch_in_1 = c1.get("punch_in", "")
                            punch_out_1 = c1.get("punch_out", "")
                            day_hours += c1.get("duration", 0.0)
                        if len(records) >= 2:
                            c2 = records[1]
                            punch_in_2 = c2.get("punch_in", "")
                            punch_out_2 = c2.get("punch_out", "")
                            day_hours += c2.get("duration", 0.0)
                total_pay_period_hours += day_hours
                week_num = current_day.isocalendar()[1]
                weekly_hours[week_num] = weekly_hours.get(week_num, 0.0) + day_hours
                def dash_if_empty(val):
                    return val if val else "-"
                row_data = [
                    day_str,
                    dash_if_empty(punch_in_1),
                    dash_if_empty(punch_out_1),
                    dash_if_empty(punch_in_2),
                    dash_if_empty(punch_out_2),
                    f"{day_hours:.2f}" if day_hours > 0 else "-"
                ]
                for w, cell in zip(widths, row_data):
                    pdf.cell(w, 10, cell, border=1, align="C")
                pdf.ln()
            current_day += datetime.timedelta(days=1)
        pdf.ln(5)
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Weekly Hours Summary:", ln=True, align="C")
        pdf.set_font("Arial", "", 12)
        for week, whours in weekly_hours.items():
            pdf.cell(0, 10, f"Week {week}: {whours:.2f} hours", ln=True, align="C")
        pdf.ln(5)
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, f"Total Pay Period Hours: {total_pay_period_hours:.2f}", ln=True, align="C")
        pdf.output(pdf_filepath)

    def show_finalize_options_popup(self, pdf_filepath):
        popup = tk.Toplevel(self.master)
        popup.title("Finalize Options")
        popup.configure(bg="white")
        popup.geometry("400x150")
        center_window_on_parent(popup, self.master)
        tk.Label(
            popup,
            text="Finalization Complete",
            font=("Arial", 16, "bold"),
            fg=MAIN_COLOR,
            bg="white"
        ).pack(pady=10)
        button_frame = tk.Frame(popup, bg="white")
        button_frame.pack(pady=10)
        def open_folder_containing(path):
            folder_path = os.path.dirname(path)
            if platform.system() == "Windows":
                os.startfile(folder_path)
            elif platform.system() == "Darwin":
                subprocess.call(["open", folder_path])
            else:
                subprocess.call(["xdg-open", folder_path])
        def view_and_close():
            self.open_pdf(pdf_filepath)
            popup.destroy()
        def send_and_close():
            subject = f"{self.username} Hours - {self.current_pay_period['end_date']}"
            import urllib.parse
            subject_encoded = urllib.parse.quote(subject)
            mailto_url = f"https://mail.google.com/mail/?view=cm&fs=1&to=rick@todaysclassroom.com&su={subject_encoded}"
            webbrowser.open(mailto_url)
            open_folder_containing(pdf_filepath)
            popup.destroy()
        view_button = tk.Button(
            button_frame,
            text="View PDF",
            command=view_and_close,
            font=("Arial", 14, "bold"),
            bd=4,
            relief="raised",
            width=10,
            bg=MAIN_COLOR,
            fg="white"
        )
        view_button.grid(row=0, column=0, padx=10)
        send_button = tk.Button(
            button_frame,
            text="Send to Rick",
            command=send_and_close,
            font=("Arial", 14, "bold"),
            bd=4,
            relief="raised",
            width=10,
            bg=MAIN_COLOR,
            fg="white"
        )
        send_button.grid(row=0, column=1, padx=10)

    def open_pdf(self, pdf_filepath):
        try:
            if platform.system() == "Windows":
                os.startfile(pdf_filepath)
            elif platform.system() == "Darwin":
                subprocess.call(["open", pdf_filepath])
            else:
                subprocess.call(["xdg-open", pdf_filepath])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {e}")

    def update_ui(self):
        if not self.current_pay_period:
            return
        period_start = self.current_pay_period["start_date"]
        period_end = self.current_pay_period["end_date"]
        if hasattr(self, "payperiod_label") and self.payperiod_label:
            self.payperiod_label.config(text=f"Current Pay Period: {period_start} to {period_end}")
        # Update punch button states based on today's records
        if self.today.weekday() >= 5:
            self.punch_in_button.config(state="disabled")
            self.punch_out_button.config(state="disabled")
        else:
            today_str = self.today.isoformat()
            records = self.current_pay_period.get("records", {})
            cycles = records.get(today_str, [])
            if cycles and ("punch_in" in cycles[-1] and "punch_out" not in cycles[-1]):
                self.punch_in_button.config(state="disabled")
                self.punch_out_button.config(state="normal")
            else:
                if len(cycles) >= 2:
                    self.punch_in_button.config(state="disabled")
                    self.punch_out_button.config(state="disabled")
                else:
                    self.punch_in_button.config(state="normal")
                    self.punch_out_button.config(state="disabled")
        # Update daily punches summary (horizontal line above Last Punch)
        today_str = self.today.isoformat()
        records = self.current_pay_period.get("records", {}).get(today_str, [])
        punch_texts = []
        for idx, cycle in enumerate(records, start=1):
            if "punch_in" in cycle:
                text = f"Cycle {idx}: {cycle.get('punch_in')} - {cycle.get('punch_out', '---')}"
            else:
                text = (f"Final: {cycle.get('punch_in_1')} - {cycle.get('punch_out_1')}, "
                        f"{cycle.get('punch_in_2', '')} - {cycle.get('punch_out_2', '')}")
            punch_texts.append(text)
        if punch_texts:
            self.daily_punches_line.config(text="Daily Punches: " + "  |  ".join(punch_texts))
        else:
            self.daily_punches_line.config(text="Daily Punches: None")
        # Update Last Punch
        last = None
        for day in sorted(self.current_pay_period.get("records", {})):
            for cycle in self.current_pay_period["records"][day]:
                t = cycle.get("punch_out") or cycle.get("punch_in")
                if t:
                    last = t
        if hasattr(self, "last_punch_label") and self.last_punch_label:
            self.last_punch_label.config(text=f"Last Punch: {last if last else 'None'}")
        # Daily total
        daily_total = sum(
            cycle.get("duration", 0.0)
            for cycle in self.current_pay_period.get("records", {}).get(self.today.isoformat(), [])
        )
        if hasattr(self, "daily_total_label") and self.daily_total_label:
            self.daily_total_label.config(text=f"Today's Hours: {daily_total:.2f}")
        # Pay period total
        period_total = 0.0
        for day in self.current_pay_period.get("records", {}):
            for cycle in self.current_pay_period["records"][day]:
                period_total += cycle.get("duration", 0.0)
        if hasattr(self, "period_total_label") and self.period_total_label:
            self.period_total_label.config(text=f"Pay Period Total Hours: {period_total:.2f}")

def main():
    root = tk.Tk()
    app = TimePunchApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
