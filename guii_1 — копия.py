import tkinter as tk
from tkinter import ttk
from certification_many_gui import main_func
from guii_2 import print_func

root = tk.Tk()
root.title("Document Generator")
root.geometry("400x300")  # Set a fixed window size

# Create a container frame for the canvas and scrollbar
container = tk.Frame(root)
container.pack(fill="both", expand=True)

# Create a canvas for the notebook
canvas = tk.Canvas(container)
canvas.pack(side="left", fill="both", expand=True)

# Add a vertical scrollbar to the canvas
scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")
canvas.configure(yscrollcommand=scrollbar.set)

# Create a frame inside the canvas to hold the notebook
frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor="nw")

# Function to adjust the scroll region
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

frame.bind("<Configure>", on_frame_configure)

# Track sheet frames and available numbers
sheets = {}
available_numbers = []

def on_ok():
    """
    Collect data from all sheets and process it.
    """
    variables = {
        'worksheets_count': len(sheets),
    }

    for sheet_name, sheet_frame in sheets.items():
        # Retrieve entries for the current sheet
        entries = sheet_frame['entries']
        variables['{{CONTRACT_NUMBER}}'] = entries['contract_number'].get()
        variables['{{CONTRACT_YEAR}}'] = entries['contract_year'].get()
        variables['{{COMPANY_NAME}}'] = entries['company_name'].get()

    # main_func(variables)
    print_func(variables)

def create_sheet():
    # Determine the next available sheet number
    if available_numbers:
        sheet_number = min(available_numbers)
        available_numbers.remove(sheet_number)
    else:
        sheet_number = len(sheets) + 1

    sheet_name = f"Sheet {sheet_number}"

    # Create a new frame for each sheet
    frame = ttk.Frame(notebook)
    notebook.add(frame, text=sheet_name)

    # Dictionary to store Entry widgets for easy access
    entries = {}

    # Add labels and entry fields to the sheet
    tk.Label(frame, text="CONTRACT_NUMBER:").grid(row=0, column=0, padx=10, pady=5)
    entries['contract_number'] = tk.Entry(frame)
    entries['contract_number'].grid(row=0, column=1, padx=10, pady=5)

    tk.Label(frame, text="CONTRACT_YEAR:").grid(row=1, column=0, padx=10, pady=5)
    entries['contract_year'] = tk.Entry(frame)
    entries['contract_year'].grid(row=1, column=1, padx=10, pady=5)

    tk.Label(frame, text="COMPANY_NAME:").grid(row=2, column=0, padx=10, pady=5)
    entries['company_name'] = tk.Entry(frame)
    entries['company_name'].grid(row=2, column=1, padx=10, pady=5)

    # Store the frame and its entries in the sheets dictionary
    sheets[sheet_name] = {'frame': frame, 'entries': entries}

def remove_sheet():
    if notebook.tabs():  # Check if there are any tabs
        # Get the last tab
        last_tab = notebook.tabs()[-1]
        # Find the name associated with the last tab ID
        for name, sheet_data in list(sheets.items()):
            if str(sheet_data['frame']) == last_tab:
                # Extract sheet number from the name and add it back to available_numbers
                sheet_number = int(name.split()[-1])
                available_numbers.append(sheet_number)

                # Remove the tab from the notebook and delete it from the dictionary
                notebook.forget(sheet_data['frame'])
                del sheets[name]
                break

# Create the Notebook widget inside the frame
notebook = ttk.Notebook(frame)
notebook.pack(pady=10, padx=10, fill="both", expand=True)

# Initial setup: Add one default sheet
create_sheet()

# Buttons to add, remove sheets, and a global OK button
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

add_button = tk.Button(button_frame, text="Add Frame", command=create_sheet)
add_button.pack(side=tk.LEFT, padx=5)

remove_button = tk.Button(button_frame, text="Remove Frame", command=remove_sheet)
remove_button.pack(side=tk.LEFT, padx=5)

ok_button = tk.Button(button_frame, text="OK", command=on_ok)
ok_button.pack(side=tk.LEFT, padx=5)

# Run the main event loop
root.mainloop()
