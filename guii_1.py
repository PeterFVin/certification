import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook

from guii_2 import print_func
from certification_many_gui import main_func


var_list = (
    ('CONTRACT_NUMBER', 'B2', 'Введите номер договора, например 19'),
    ('CONTRACT_YEAR', 'B3', 'Введите 2 последние цифры года договора, например 24'),
    ('WORK_NUMBER', 'B4', 'Введите номер работы - если несколько работ, то например 19-01, 19-02 и т.д.24'),
    ('BILL_NUMBER', 'B5', 'Введите номер счета'),
    ('CONTR_DATE', 'B6', 'Введите дату (день) договора, например 01'),
    ('CONTR_MONTH', 'B7', 'CONTR_MONTH', 'Введите месяц договора числом, например 01 если январь'),
    ('BUSINESS_FORM_FULL', 'B8', 'Введите организационно-правовую форму полностью, например общество с ограниченной ответственностью'),
    ('COMPANY_NAME_FULL', 'B9', 'Введите наименование организации (полное)'),
    ('BUSINESS_FORM', 'B10', 'Введите организационно-правовую форму кратко, например ООО'),
    ('COMPANY_NAME', 'B11', 'Введите сокращенное имя компании'),
    ('DIR_LASTNAME', 'B12', 'Введите фамилию директора'),
    ('DIR_FIRSTNAME', 'B13', 'Введите имя директора'),
    ('DIR_SECNAME', 'B14', 'Введите отчество директора'),
    ('GENDER', 'B15', 'Введите пол директора, строго М или Ж'),
    ('CERT_NAME', 'B16', 'Введите наименование сертификата'),
    ('OKPD', 'B17', 'Введите код ОКПД 2'),
    ('STANDART_MAIN', 'B18', 'Введите основной стандарт (указывается в акте отбора)'),
    ('STANDART_SHORT', 'B19', 'Введите перечень стандартов - без пунктов'),
    ('STANDART_FULL', 'B20', 'Введите перечень стандартов с пунктами (если есть отдельные пункты'),
    ('CONTRACT_SUM', 'B21', 'Введите полную сумму договора)'),
    ('CONTRACT_OS_FULL_SUM', 'B22', 'Введите общую сумму оплаты услуг по ОС'),
    ('CONTRACT_OS_SUM', 'B23', 'Введите общую сумму оплаты услуг по ОС по одному сертифицируемому объекту'),
    ('CONTRACT_IL_FULL_SUM', 'B24', 'Введите общую сумму оплаты услуг по ИЛ'),
    ('CONTRACT_IL_SUM', 'B25', 'Введите общую сумму оплаты услуг по ИЛ по одному сертифицируемому объекту'),
    ('CONTRACT_IK_SUM', 'B26', 'Введите сумму договора инспекционного контроля'),
    ('INN', 'B27', 'Введите ИНН организации'),
    ('KPP', 'B28', 'Введите КПП организации'),
    ('JUR_ADDRESS', 'B29', 'Введите юридический адрес организации'),
    ('PHYS_ADDRESS', 'B30', 'Введите физический адрес организации'),
    ('RAS_ACCOUNT', 'B31', 'Введите расчетный счет организации'),
    ('BANK_NAME', 'B32', 'Введите имя банка организации'),
    ('KORR_ACCOUNT', 'B33', 'Введите корреспондентский счет организации'),
    ('BIK', 'B34', 'Введите БИК банка организации'),
    ('OGRN', 'B35', 'Введите ЫЫЫЫЫЫЫЫЫЫЫЫЫЫ организации'),
    ('TEL', 'B36', 'Введите телефон организации'),
    ('E-MAIL', 'B37', 'Введите e-mail организации'),
    ('EXPERT_LASTNAME', 'B38', 'Введите фамилию эксперта по сертификации'),
    ('EXPERT_FIRSTNAME_SHORT', 'B39', 'Введите сокращенное имя эксперта по сертификации (с точкой) например П.'),
    ('EXPERT_SECNAME_SHORT', 'B40', 'Введите сокращенное отчество эксперта по сертификации (с точкой) например Ф.'),
    ('EXPERT_REG_NUMBER', 'B41', 'Введите номер эксперта в реестре МСС'),
    ('IL_NAME', 'B42', 'Введите наименование ИЛ'),
    ('IL_REG_NUMBER', 'B43', 'Введите номер ИЛ в реестре МСС'),
    ('IL_EXPIRE_DATE', 'B44', 'Введите дату окончания действия аккредитация ИЛ'),
    ('ISSUE_DECISION_DATE', 'B45', 'Введите дату (день) решения о выдаче сертификата, например 01'),
    ('ISSUE_DECISION_MONTH', 'B46', 'Введите месяц решения о выдаче сертификата - числом, например 01 если январь'),
    ('ISSUE_DECISION_YEAR', 'B47', 'Введите год решения о выдаче сертификата - четыре цифры'),
    ('CERTIFICARTE_START_DATE', 'B48', 'Введите дату (день) начала действия сертификата, например 01'),
    ('CERTIFICARTE_START_MONTH', 'B49', 'Введите Введите месяц начала действия сертификата - числом, например 01 если январь'),
    ('CERTIFICARTE_START_YEAR', 'B50', 'Введите год начала действия сертификата - четыре цифры'),
    ('CERTIFICARTE_DURATION', 'B51', 'Введите срок действия сертификата, например 3'),
    ('SAMPLE_ACT_DATE', 'B52', 'Введите дату (день) акта отбора образцов, например 01'),
    ('SAMPLE_ACT_MONTH', 'B53', 'Введите месяц акта отбора образцов - числом, например 01 если январь'),
    ('SAMPLE_ACT_YEAR', 'B54', 'Введите год акта отбора образцов - четыре цифры'),
    ('PRODUCTION_CREATED_QUARTER', 'B55', 'Введите квартал изготовления продукции, например IV'),
    ('PRODUCTION_CREATED_YEAR', 'B56', 'Введите год изготовления продукции - четыре цифры'),
    # ('ЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫ', 'B57', 'Введите ЫЫЫЫЫЫЫЫЫЫЫЫЫЫ'),
    # ('ЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫ', 'B58', 'Введите ЫЫЫЫЫЫЫЫЫЫЫЫЫЫ'),
    # ('ЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫ', 'B59', 'Введите ЫЫЫЫЫЫЫЫЫЫЫЫЫЫ'),






    # ('ЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫЫ', 'B5Ы', 'Введите ЫЫЫЫЫЫЫЫЫЫЫЫЫЫ'),
)

# Function to import data from Excel
def import_data():
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return  # If no file is selected, do nothing


    # Load the workbook and get sheet names
    workbook = load_workbook(file_path)
    sheet_names = workbook.sheetnames
    # Clear existing sheets
    global sheets
    sheets = {}
    for tab in notebook.tabs():
        notebook.forget(tab)
    # Loop through each sheet in the Excel file
    for idx, sheet_name in enumerate(sheet_names, start=1):
        worksheet = workbook[sheet_name]
        # Create a new sheet in the Tkinter notebook
        create_sheet()
        
        
        
        # Populate data from the Excel sheet
        entries = sheets[f"Серт {idx}"]['entries']
        for i in range(len(var_list)):
            entries[var_list[i][0]].delete(0, tk.END)
            entries[var_list[i][0]].insert(0, worksheet[var_list[i][1]].value)
            
            # entries['contract_year'].delete(0, tk.END)
            # entries['contract_year'].insert(0, worksheet["A2"].value)
            
            # entries['company_name'].delete(0, tk.END)
            # entries['company_name'].insert(0, worksheet["A3"].value)


def validate_fields(entries):
    """
    Validate the fields of a single sheet. Returns a list of error messages.
    """
    errors = []

    # Validate CONTRACT_NUMBER
    contract_number = entries[var_list[0][0]].get()
    if not contract_number.isdigit() or int(contract_number) > 100:
        errors.append("Field 'CONTRACT_NUMBER' should be an integer <= 100.")

    # Validate CONTRACT_YEAR
    contract_year = entries['CONTRACT_YEAR'].get()
    if not contract_year.isdigit():
        errors.append("Field 'CONTRACT_YEAR' should be an integer.")

    # # Validate COMPANY_NAME
    # company_name = entries['company_name'].get()
    # if not all(char in "AB" for char in company_name):
    #     errors.append("Field 'COMPANY_NAME' should contain only letters 'A' or 'B'.")

    return errors


root = tk.Tk()
root.title("Document Generator")
root.geometry("400x400")  # Set a fixed window size

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
    Validation, Collect data from all sheets and process it.
    """
    
    all_errors = []

    # Iterate through all sheets and validate their fields
    for sheet_name, sheet_data in sheets.items():
        entries = sheet_data['entries']
        errors = validate_fields(entries)
        if errors:
            all_errors.append(f"Errors in {sheet_name}:")
            all_errors.extend(errors)

    # If there are errors, show them in a messagebox
    if all_errors:
        messagebox.showerror("Validation Errors", "\n".join(all_errors))
        return

    variables = {}
    counter = 0

    for sheet_name, sheet_frame in sheets.items():
        counter += 1
        variables[counter] = {}

        entries = sheet_frame['entries']
        variables[counter]['worksheets_count'] = len(sheets)
        for i in range(len(var_list)):
            variables[counter]['{{' + var_list[i][0] + '}}'] = entries[var_list[i][0]].get()
            
            
            # variables[counter]['{{CONTRACT_NUMBER}}'] = entries[var_list[0][0]].get()
            # variables[counter]['{{CONTRACT_YEAR}}'] = entries['contract_year'].get()
            # variables[counter]['{{COMPANY_NAME}}'] = entries['company_name'].get()

    print_func(variables)
    main_func(variables)
    messagebox.showinfo('Программа выполнена!', 'Программа отработала успешно, все документы созданы!')

def create_sheet():
    # Determine the next available sheet number
    sheet_number = len(sheets) + 1
    sheet_name = f"Серт {sheet_number}"

    # Create a new frame for each sheet
    frame = ttk.Frame(notebook)
    notebook.add(frame, text=sheet_name)

    # Dictionary to store Entry widgets for easy access
    entries = {}

    # Retrieve data from the last sheet if it exists
    last_sheet_name = f"Серт {sheet_number - 1}"
    last_entries = sheets.get(last_sheet_name, {}).get('entries', {})

    # Add labels, entry fields, and tip text to the sheet
    for i in range(len(var_list)):
        tk.Label(frame, text=f'{var_list[i][0]}:').grid(row=i, column=0, padx=10, pady=5, sticky="e")
        entries[var_list[i][0]] = tk.Entry(frame)
        entries[var_list[i][0]].grid(row=i, column=1, padx=10, pady=5)
        if last_entries:
            entries[var_list[i][0]].insert(0, last_entries[var_list[i][0]].get())
        tk.Label(frame, text=var_list[i][2]).grid(row=i, column=2, padx=10, pady=5, sticky="w")
    


    # tk.Label(frame, text="CONTRACT_YEAR:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    # entries['contract_year'] = tk.Entry(frame)
    # entries['contract_year'].grid(row=1, column=1, padx=10, pady=5)
    # if last_entries:
    #     entries['contract_year'].insert(0, last_entries['contract_year'].get())
    # tk.Label(frame, text="Enter a valid year (e.g., 2024)").grid(row=1, column=2, padx=10, pady=5, sticky="w")

    # tk.Label(frame, text="COMPANY_NAME:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    # entries['company_name'] = tk.Entry(frame)
    # entries['company_name'].grid(row=2, column=1, padx=10, pady=5)
    # if last_entries:
    #     entries['company_name'].insert(0, last_entries['company_name'].get())
    # tk.Label(frame, text="Enter 'A' or 'B' only").grid(row=2, column=2, padx=10, pady=5, sticky="w")

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

import_button = tk.Button(button_frame, text="Import Data", command=import_data)
import_button.pack(side=tk.LEFT, padx=5)

ok_button = tk.Button(button_frame, text="OK", command=on_ok)
ok_button.pack(side=tk.LEFT, padx=5)

# Run the main event loop
root.mainloop()
