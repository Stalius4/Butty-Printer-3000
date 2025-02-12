import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import win32com.client
import re
import csv
from ttkthemes import ThemedTk

# ------------------------------------------------------------------------
# Global variables
# ------------------------------------------------------------------------
BASE_DIR = r"C:\Users\Deivydas\Desktop\label_printer1"  # Replace with your actual path
current_price = "£0.00"
CSV_FILE = "Spalding_numbers.csv"         # For white and brown labels
CSV_PANINIS = "Spalding_paninis.CSV"        # For "other" (panini's) labels

# This dictionary maps each tab (ttk.Frame) to its folder path.
tab_folders = {}

# Stores references to our file widgets keyed by:
# (folder_path, (folder_type, filename)) -> (actual_subfolder, entry_widget)
file_widgets = {}

# Global dictionary to map (folder_path, row, col) -> entry widget for arrow-key navigation
grid_entries = {}

# Dictionary to store day buttons for each tab (keyed by folder_path)
day_buttons_by_tab = {}

# --- NEW: Dictionaries to store the canvas and its scrollable frame for each folder ---
canvas_by_folder = {}
scrollable_frames = {}

# Global references to main window objects
root = None
notebook = None
current_tab_total_label = None
price_label = None

# ------------------------------------------------------------------------
# Utility Functions
# ------------------------------------------------------------------------
def natural_key(text):
    """
    Splits the string into alpha and numeric parts so that sorting is done in natural order.
    """
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r'(\d+)', text)]

# ------------------------------------------------------------------------
# Price reading
# ------------------------------------------------------------------------
def get_price_from_label(file_path):
    """Read the 'Price' value from a label file."""
    try:
        bpac = win32com.client.Dispatch("bpac.Document")
        if bpac.Open(file_path):
            price_object = bpac.GetObject("Price")
            return price_object.Text if price_object else None
    except Exception as e:
        print(f"Error reading price from {file_path}: {e}")
    finally:
        bpac = None
    return None

# ------------------------------------------------------------------------
# Updating price display
# ------------------------------------------------------------------------
def update_price_display():
    """
    Uses the first label from the 'white' subfolder of the current folder to update the displayed price.
    """
    global current_price
    current_tab_name = notebook.select()
    current_tab = notebook.nametowidget(current_tab_name)
    folder_path = tab_folders.get(current_tab)
    if not folder_path:
        price_label.config(text="Current Price: £0.00")
        return

    price = "£0.00"
    white_path = os.path.join(folder_path, "white")
    if os.path.isdir(white_path):
        files = [f for f in os.listdir(white_path) if f.lower().endswith((".lbx", ".lbl"))]
        if files:
            first_label = os.path.join(white_path, files[0])
            price = get_price_from_label(first_label) or "£0.00"
    current_price = price
    price_label.config(text=f"Current Price: {current_price}")

# ------------------------------------------------------------------------
# Set Price
# ------------------------------------------------------------------------
def set_price():
    """
    Prompts the user for a new price and updates all 'white' and 'brown' labels in the current folder.
    """
    global current_price
    current_tab_name = notebook.select()
    current_tab = notebook.nametowidget(current_tab_name)
    folder_path = tab_folders.get(current_tab)
    if not folder_path:
        messagebox.showinfo("Info", "No folder selected.")
        return

    new_price = simpledialog.askstring("Set Price", "Enter new price (e.g., £2.50):")
    if not new_price:
        return
    if not new_price.startswith("£"):
        new_price = f"£{new_price}"

    current_price = new_price
    price_label.config(text=f"Current Price: {current_price}")

    # Only update white and brown labels
    files_to_process = []
    for (fpath, key_tuple), (actual_subfolder, entry_widget) in file_widgets.items():
        if fpath == folder_path:
            folder_type, filename = key_tuple
            if folder_type != "other":
                files_to_process.append((actual_subfolder, filename))

    if not files_to_process:
        messagebox.showinfo("Info", "No labels to update.")
        return

    total_files = len(files_to_process)
    progress_window = tk.Toplevel(root)
    progress_window.title("Updating Prices")
    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
    progress_bar["maximum"] = total_files
    progress_bar.pack(padx=20, pady=10)
    progress_label = tk.Label(progress_window, text=f"0/{total_files}")
    progress_label.pack(pady=5)
    progress_window.grab_set()
    progress_window.update()

    # Center the progress window
    progress_window.update_idletasks()
    pw_width = progress_window.winfo_width()
    pw_height = progress_window.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - pw_width) // 2
    y = (screen_height - pw_height) // 2
    progress_window.geometry(f"+{x}+{y}")

    for index, (folder, file) in enumerate(files_to_process):
        label_path = os.path.join(folder, file)
        try:
            bpac = win32com.client.Dispatch("bpac.Document")
            if bpac.Open(label_path):
                price_obj = bpac.GetObject("Price")
                if isinstance(price_obj, win32com.client.CDispatch):
                    price_obj.Text = new_price
                    if not bpac.Save():
                        print(f"Warning: Failed to save {label_path}")
                else:
                    print(f"Invalid Price object in {file}")
        except Exception as e:
            print(f"Error updating {file}: {e}")
        finally:
            bpac = None

        progress_bar["value"] = index + 1
        progress_label.config(text=f"{index+1}/{total_files}")
        progress_window.update()

    progress_window.destroy()
    update_price_display()

# ------------------------------------------------------------------------
# Print Labels
# ------------------------------------------------------------------------
def print_labels():
    """
    Prints all labels in the current folder that have a quantity > 0.
    """
    current_tab_name = notebook.select()
    current_tab = notebook.nametowidget(current_tab_name)
    folder_path = tab_folders.get(current_tab)
    if not folder_path:
        messagebox.showinfo("Info", "No folder selected.")
        return

    labels_to_print = []
    for (fpath, key_tuple), (actual_subfolder, entry_widget) in file_widgets.items():
        if fpath == folder_path:
            folder_type, file = key_tuple
            copies_str = entry_widget.get().strip()
            copies = int(copies_str) if copies_str.isdigit() else 0
            if copies > 0:
                file_path = os.path.join(actual_subfolder, file)
                labels_to_print.append((file_path, copies))

    if not labels_to_print:
        messagebox.showinfo("Info", "No labels selected.")
        return

    try:
        bpac = win32com.client.Dispatch("bpac.Document")
        if not bpac.StartPrint("", 0):
            messagebox.showerror("Error", "Failed to start printing.")
            return

        for file_path, copies in labels_to_print:
            try:
                if bpac.Open(file_path):
                    bpac.PrintOut(copies, 1)  # second argument indicates cut mode
                else:
                    messagebox.showerror("Error", f"Failed to open: {os.path.basename(file_path)}")
            except Exception as e:
                print(f"Error printing {file_path}: {e}")
                continue

        bpac.EndPrint()
        bpac.Close()
        messagebox.showinfo("Success", "All selected labels printed successfully!")
    except Exception as e:
        print(e)
    finally:
        bpac = None

# ------------------------------------------------------------------------
# Per-Tab Dynamic Sum Calculation
# ------------------------------------------------------------------------
def update_tab_total_display(folder_path):
    """
    Recalculates and displays the total for the current folder.
    """
    total = 0
    for (fpath, key_tuple), (actual_subfolder, entry_widget) in file_widgets.items():
        if fpath == folder_path:
            try:
                value = float(entry_widget.get())
                total += int(value)
            except ValueError:
                pass
    current_tab_total_label.config(text=f"Total: {total}")

def entry_update(event):
    """
    Called when an entry is updated; recalculates the total for the current folder.
    """
    current_tab_name = notebook.select()
    current_tab = notebook.nametowidget(current_tab_name)
    current_folder_path = tab_folders.get(current_tab)
    changed_entry = event.widget
    changed_entry_folder_path = None
    for (fpath, key_tuple), (actual_subfolder, entry_widget) in file_widgets.items():
        if entry_widget == changed_entry:
            changed_entry_folder_path = fpath
            break
    if changed_entry_folder_path == current_folder_path:
        update_tab_total_display(current_folder_path)

# ------------------------------------------------------------------------
# On Tab Change: Enable/Disable Day Buttons
# ------------------------------------------------------------------------
def on_tab_change(event):
    """
    Called whenever the user switches tabs. Enables all day buttons for the "Spalding" tab,
    and for other tabs enables only the "Reset" button.
    """
    update_price_display()
    current_tab_name = notebook.select()
    current_tab = notebook.nametowidget(current_tab_name)
    folder_path = tab_folders.get(current_tab)
    for fp, buttons in day_buttons_by_tab.items():
        if fp == folder_path:
            if os.path.basename(fp) == "Spalding":
                for btn in buttons:
                    btn.config(state="normal")
            else:
                for btn in buttons:
                    if btn['text'] == "Reset":
                        btn.config(state="normal")
                    else:
                        btn.config(state="disabled")
        else:
            for btn in buttons:
                btn.config(state="disabled")
    if folder_path:
        update_tab_total_display(folder_path)

# ------------------------------------------------------------------------
# Populate entries from CSV when a day button is pressed
# ------------------------------------------------------------------------
def populate_day(day, folder_path):
    """
    Reads data from two CSV files and populates the inputs:
      - White and brown inputs are updated from CSV_FILE.
      - Other inputs are updated from CSV_PANINIS.
    
    CSV_FILE is expected to have headers like:  
         Name,Monday white,Tuesday white,Wednesday white,...,Monday brown,Tuesday brown,...  
    CSV_PANINIS is expected to have headers like:  
         Name,Monday other,Tuesday other,Wednesday other,...,Saturday other
         
    This function strips extra spaces from keys to help match your CSV file.
    """
    # Load CSV data for white and brown
    data_main = {}
    try:
        with open(CSV_FILE, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                # Strip keys and values
                row = {k.strip(): (v.strip() if v is not None else "") for k, v in row.items() if k}
                key = row.get("Name", "").strip()
                if key:
                    data_main[key] = row
    except Exception as e:
        messagebox.showerror("Error", f"Could not read CSV file: {e}")
        return

    # Load CSV data for "other" (paninis)
    data_paninis = {}
    try:
        with open(CSV_PANINIS, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                row = {k.strip(): (v.strip() if v is not None else "") for k, v in row.items() if k}
                key = row.get("Name", "").strip()
                if key:
                    data_paninis[key] = row
    except Exception as e:
        messagebox.showerror("Error", f"Could not read CSV file ({CSV_PANINIS}): {e}")
        return

    # Update each input field in the current folder
    for (fpath, key_tuple), (actual_subfolder, entry_widget) in file_widgets.items():
        if fpath == folder_path:
            folder_type, filename = key_tuple
            base_name = os.path.splitext(filename)[0]
            if folder_type in ("white", "brown"):
                if base_name in data_main:
                    csv_row = data_main[base_name]
                    col_name = f"{day} {folder_type}".strip()  # e.g., "Monday white"
                    value = csv_row.get(col_name, "")
                    entry_widget.delete(0, tk.END)
                    entry_widget.insert(0, value)
            elif folder_type == "other":
                if base_name in data_paninis:
                    csv_row = data_paninis[base_name]
                    col_name = f"{day} {folder_type}".strip()  # e.g., "Monday other"
                    value = csv_row.get(col_name, "")
                    entry_widget.delete(0, tk.END)
                    entry_widget.insert(0, value)
    update_tab_total_display(folder_path)

# ------------------------------------------------------------------------
# Arrow Key Navigation Functionality with Auto-Scroll
# ------------------------------------------------------------------------
def navigate_arrow(event):
    """
    Enables navigation among input fields using arrow keys.
    After moving focus, if the target widget is not fully visible,
    scroll the canvas so that it becomes visible.
    """
    widget = event.widget
    folder_path = getattr(widget, 'folder_path', None)
    if folder_path is None:
        return
    row = getattr(widget, 'row', None)
    col = getattr(widget, 'col', None)
    if row is None or col is None:
        return

    new_row, new_col = row, col
    if event.keysym == "Up":
        new_row = row - 1
    elif event.keysym == "Down":
        new_row = row + 1
    elif event.keysym == "Left":
        new_col = col - 1
    elif event.keysym == "Right":
        new_col = col + 1

    target = grid_entries.get((folder_path, new_row, new_col))
    if target:
        target.focus_set()
        # --- NEW: Auto-scroll if the target is not visible ---
        if folder_path in canvas_by_folder and folder_path in scrollable_frames:
            canvas = canvas_by_folder[folder_path]
            scroll_frame = scrollable_frames[folder_path]
            # Get the target's top position relative to the screen and the canvas's top position.
            target_top = target.winfo_rooty()
            target_bottom = target_top + target.winfo_height()
            canvas_top = canvas.winfo_rooty()
            canvas_bottom = canvas_top + canvas.winfo_height()
            # If the target is above or below the visible area of the canvas, scroll.
            if target_top < canvas_top or target_bottom > canvas_bottom:
                # Compute the target's relative y-position inside the scrollable content.
                rel_y = target.winfo_rooty() - scroll_frame.winfo_rooty()
                total_height = scroll_frame.winfo_height()
                fraction = rel_y / total_height
                canvas.yview_moveto(fraction)

# ------------------------------------------------------------------------
# Global Mouse Wheel Handler for Scrolling
# ------------------------------------------------------------------------
def on_global_mousewheel(event):
    current_tab_name = notebook.select()
    current_tab = notebook.nametowidget(current_tab_name)
    canvas = None
    for child in current_tab.winfo_children():
        if isinstance(child, tk.Canvas):
            canvas = child
            break
    if canvas:
        canvas.yview_scroll(-1 * int(event.delta / 120), "units")

# ------------------------------------------------------------------------
# Build Tabs
# ------------------------------------------------------------------------
def build_tabs():
    """
    Creates a tab for each folder in BASE_DIR. Each tab includes sub-frames
    for white, brown, and other labels plus a side frame with day buttons.
    """
    folders = [f for f in os.listdir(BASE_DIR) if os.path.isdir(os.path.join(BASE_DIR, f))]
    
    for folder_name in folders:
        folder_path = os.path.join(BASE_DIR, folder_name)
        
        # Create a tab for this folder.
        folder_tab = ttk.Frame(notebook)
        notebook.add(folder_tab, text=folder_name)
        tab_folders[folder_tab] = folder_path

        # Configure grid layout for the tab.
        folder_tab.grid_rowconfigure(0, weight=1)
        folder_tab.grid_columnconfigure(0, weight=1)   # Canvas column
        folder_tab.grid_columnconfigure(1, weight=0)   # Scrollbar column
        folder_tab.grid_columnconfigure(2, weight=0)   # Days frame column

        # Create a canvas and vertical scrollbar.
        canvas = tk.Canvas(folder_tab, bd=0, highlightthickness=0, relief='flat')
        scrollbar = ttk.Scrollbar(folder_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # Attach the scrollable frame to the canvas.
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        # --- NEW: Save canvas and scrollable_frame for auto-scrolling ---
        canvas_by_folder[folder_path] = canvas
        scrollable_frames[folder_path] = scrollable_frame

        # Create the days frame on the right side.
        days_frame = tk.Frame(folder_tab, padx=5, pady=5)
        days_frame.grid(row=0, column=2, sticky="nsew")

        # Create day buttons.
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Reset"]
        tab_day_buttons = []
        for day in days:
            btn = ttk.Button(days_frame, text=day, width=12,
                             command=lambda d=day, fp=folder_path: populate_day(d, fp))
            btn.pack(side="top", fill="x", pady=2)
            tab_day_buttons.append(btn)
        day_buttons_by_tab[folder_path] = tab_day_buttons

        # Disable day buttons for non-Spalding tabs except Reset.
        if folder_name != "Spalding":
            for btn in tab_day_buttons:
                if btn['text'] != "Reset":
                    btn.config(state="disabled")
                else:
                    btn.config(state="normal")

        # Configure the canvas to update its scroll region.
        scrollable_frame.bind(
            "<Configure>",
            lambda e, canvas=canvas: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Create sub-frames for white, brown, and other labels.
        white_frame = tk.Frame(scrollable_frame, bd=0, relief="flat")
        brown_frame = tk.Frame(scrollable_frame, bd=0, relief="flat")
        other_frame = tk.Frame(scrollable_frame, bd=0, relief="flat")
        white_frame.pack(side="left", fill="both", expand=True, padx=0, pady=0)
        brown_frame.pack(side="left", fill="both", expand=True, padx=0, pady=0)
        other_frame.pack(side="left", fill="both", expand=True, padx=0, pady=0)

        # ---------------------------
        # WHITE LABELS
        # ---------------------------
        tk.Label(white_frame, text="W", font=("Calibri", 10, "bold")).pack(anchor="e", padx=0, pady=0)
        white_path = os.path.join(folder_path, "white")
        if os.path.isdir(white_path):
            files = sorted([f for f in os.listdir(white_path) if f.lower().endswith((".lbx", ".lbl"))], key=natural_key)
            white_row_index = 0
            for lbl_file in files:
                row_frame = tk.Frame(white_frame)
                row_frame.pack(anchor="w", padx=0, pady=0)
                bottom_border = tk.Frame(row_frame, bg="grey", height=0.5)
                bottom_border.pack(fill="x", side="bottom")
                tk.Label(row_frame, text=os.path.splitext(lbl_file)[0],
                         width=25, anchor="w").pack(side="left", padx=0, pady=0)
                entry = ttk.Entry(row_frame, width=3)
                entry.pack(side="left", padx=0, pady=0)
                entry.folder_path = folder_path
                entry.row = white_row_index
                entry.col = 0  # white column
                grid_entries[(folder_path, white_row_index, 0)] = entry
                entry.bind("<Up>", navigate_arrow)
                entry.bind("<Down>", navigate_arrow)
                entry.bind("<Left>", navigate_arrow)
                entry.bind("<Right>", navigate_arrow)
                entry.bind("<KeyRelease>", entry_update)
                file_widgets[(folder_path, ("white", lbl_file))] = (white_path, entry)
                white_row_index += 1
        else:
            tk.Label(white_frame, text="No 'white' folder").pack(anchor="w")

        # ---------------------------
        # BROWN LABELS
        # ---------------------------
        tk.Label(brown_frame, text="B", font=("Calibri", 10, "bold")).pack(anchor="w")
        brown_path = os.path.join(folder_path, "brown")
        if os.path.isdir(brown_path):
            files = sorted([f for f in os.listdir(brown_path) if f.lower().endswith((".lbx", ".lbl"))], key=natural_key)
            brown_row_index = 0
            for lbl_file in files:
                row_frame = tk.Frame(brown_frame)
                row_frame.pack(anchor="w", padx=2)
                bottom_border = tk.Frame(row_frame, bg="grey", height=0.5)
                bottom_border.pack(fill="x", side="bottom")
                entry = ttk.Entry(row_frame, width=3)
                entry.insert(0, "")
                entry.pack(side="left")
                entry.folder_path = folder_path
                entry.row = brown_row_index
                entry.col = 1  # brown column
                grid_entries[(folder_path, brown_row_index, 1)] = entry
                entry.bind("<Up>", navigate_arrow)
                entry.bind("<Down>", navigate_arrow)
                entry.bind("<Left>", navigate_arrow)
                entry.bind("<Right>", navigate_arrow)
                entry.bind("<KeyRelease>", entry_update)
                file_widgets[(folder_path, ("brown", lbl_file))] = (brown_path, entry)
                brown_row_index += 1
        else:
            tk.Label(brown_frame, text="No 'brown' folder").pack(anchor="w")

        # ---------------------------
        # OTHER LABELS
        # ---------------------------
        tk.Label(other_frame, text="Panini's", font=("Calibri", 10, "bold")).pack(anchor="w", padx=30)
        other_path = os.path.join(folder_path, "other")
        if os.path.isdir(other_path):
            files = [f for f in os.listdir(other_path) if f.lower().endswith((".lbx", ".lbl"))]
            for lbl_file in files:
                row_frame = tk.Frame(other_frame)
                row_frame.pack(anchor="w", padx=30)
                bottom_border = tk.Frame(row_frame, bg="grey", height=0.5)
                bottom_border.pack(fill="x", side="bottom")
                tk.Label(row_frame, text=os.path.splitext(lbl_file)[0],
                         width=15, anchor="w").pack(side="left")
                entry = ttk.Entry(row_frame, width=3)
                entry.insert(0, "")
                entry.pack(side="left")
                entry.bind("<KeyRelease>", entry_update)
                file_widgets[(folder_path, ("other", lbl_file))] = (other_path, entry)
        else:
            tk.Label(other_frame, text="No 'other' folder").pack(anchor="w")

# ------------------------------------------------------------------------
# Main GUI
# ------------------------------------------------------------------------
def main():
    global root, notebook, price_label, current_tab_total_label
    root = ThemedTk(theme="clearlooks")
    root.title("Butty Printer 3000")
    root.geometry("550x810")
    notebook = ttk.Notebook(root)
    notebook.grid(row=0, column=0, columnspan=3, sticky="nsew")
    controls = tk.Frame(root)
    controls.grid(row=1, column=0, columnspan=3, sticky="ew")
    ttk.Button(controls, text="Set Price", command=set_price).pack(side="left", padx=5, pady=3)
    price_label = ttk.Label(controls, text=f"Current Price: {current_price}", font=("Arial", 10, "bold"))
    price_label.pack(side="left", padx=5)
    ttk.Button(controls, text="Print Labels", command=print_labels).pack(side="right", padx=15, pady=3)
    current_tab_total_label = ttk.Label(controls, text="Total: 0", font=("Arial", 10, "bold"))
    current_tab_total_label.pack(side="right", padx=5)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    build_tabs()
    notebook.bind("<<NotebookTabChanged>>", on_tab_change)
    root.bind_all("<MouseWheel>", on_global_mousewheel)
    root.update_idletasks()
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"+{x}+{y}")
    root.mainloop()

if __name__ == "__main__":
    main()
