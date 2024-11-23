import tkinter as tk
from certification_many_gui import main_func

root = tk.Tk()
root.title("Document Generator")

def on_ok():

    variables = {
    'worksheets_count': 1,
    }
    
    variables['{{CONTRACT_NUMBER}}'] = contract_number_entry.get()
    variables['{{CONTRACT_YEAR}}'] = contract_year_entry.get()
    variables['{{WORK_NUMBER}}'] = work_number_entry.get()

    main_func(variables)

# Create and place the input fields
tk.Label(root, text="CONTRACT_NUMBER:").grid(row=0, column=0, padx=10, pady=5)
contract_number_entry = tk.Entry(root)
contract_number_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="CONTRACT_YEAR:").grid(row=1, column=0, padx=10, pady=5)
contract_year_entry = tk.Entry(root)
contract_year_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="WORK_NUMBER:").grid(row=2, column=0, padx=10, pady=5)
work_number_entry = tk.Entry(root)
work_number_entry.grid(row=2, column=1, padx=10, pady=5)

# Create and place the OK button
ok_button = tk.Button(root, text="OK", command=on_ok)
ok_button.grid(row=3, columnspan=2, pady=10)

# Run the main event loop
root.mainloop()