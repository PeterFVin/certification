import tkinter as tk
from certification_many_gui import main_func
from tkinter import ttk

root = tk.Tk()
root.title("Document Generator")

def on_ok():
    variables1 = {}
    for i, tab in enumerate(tabs):
        contract_number = tab['CONTRACT_NUMBER'].get()
        contract_year = tab['CONTRACT_YEAR'].get()
        work_number = tab['WORK_NUMBER'].get()
        
        variables1[i] = {
            '{{CONTRACT_NUMBER}}': contract_number,
            '{{CONTRACT_YEAR}}': contract_year,
            '{{WORK_NUMBER}}': work_number,
            'worksheets_count': 1,
        }

    for i in range(len(variables1)):
        main_func(variables1[i])

# Create a style for the "+" button
style = ttk.Style()
style.configure('Plus.TButton', font=('Arial', 12, 'bold'), foreground='blue')

# Create a canvas to hold the notebook and the "+" button
canvas = tk.Canvas(root)
canvas.pack(expand=True, fill='both')

# Create a Notebook widget
notebook = ttk.Notebook(canvas)
notebook.pack(side='left', expand=True, fill='both')

# Create and place the "+" button
plus_button = ttk.Button(canvas, text="+", style='Plus.TButton', command=add_tab)
plus_button_id = canvas.create_window(10, 10, window=plus_button, anchor='nw')


# List to store tab data
tabs = []

def update_plus_button_position():
    canvas.update_idletasks()
    x = notebook.winfo_width()
    canvas.coords(plus_button_id, x + 5, 5)

# Function to add a new tab
def add_tab():
    tab_frame = ttk.Frame(notebook)
    tab_data = {}

    tk.Label(tab_frame, text="CONTRACT_NUMBER:").grid(row=0, column=0, padx=10, pady=5)
    contract_number_entry = tk.Entry(tab_frame)
    contract_number_entry.grid(row=0, column=1, padx=10, pady=5)
    tab_data['CONTRACT_NUMBER'] = contract_number_entry

    tk.Label(tab_frame, text="CONTRACT_YEAR:").grid(row=1, column=0, padx=10, pady=5)
    contract_year_entry = tk.Entry(tab_frame)
    contract_year_entry.grid(row=1, column=1, padx=10, pady=5)
    tab_data['CONTRACT_YEAR'] = contract_year_entry

    tk.Label(tab_frame, text="WORK_NUMBER:").grid(row=2, column=0, padx=10, pady=5)
    work_number_entry = tk.Entry(tab_frame)
    work_number_entry.grid(row=2, column=1, padx=10, pady=5)
    tab_data['WORK_NUMBER'] = work_number_entry

    tabs.append(tab_data)
    notebook.add(tab_frame, text=f'Tab {len(tabs)}')
    update_plus_button_position()

# Add initial tabs
add_tab()
add_tab()

# Create and place the "+" button
plus_button = ttk.Button(canvas, text="+", style='Plus.TButton', command=add_tab)
plus_button_id = canvas.create_window(10, 10, window=plus_button, anchor='nw')



# Add initial tabs
add_tab()
add_tab()

# Pack the canvas and place the "+" button in the correct position
canvas.pack(expand=True, fill='both')
root.update_idletasks()
update_plus_button_position()

# Create and place the OK button
ok_button = tk.Button(root, text="OK", command=on_ok)
ok_button.pack(pady=10)

# Run the main event loop
root.mainloop()