import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import openpyxl

def get_document_name():
    return input("Enter the name of the document: ")

def get_group_size():
    while True:
        try:
            size = int(input("Enter the number of people in the group: "))
            if size > 0:
                return size
            else:
                print("Please enter a valid positive number.")
        except ValueError:
            print("Please enter a valid number.")

def get_person_names(size):
    names = []
    for i in range(size):
        name = input(f"Enter the name of person {i + 1}: ")
        names.append(name)
    return names

def get_default_currency():
    return input("Enter the default currency code (e.g. DKK or EUR): ")

def create_expense_template(doc_name, people_names, default_currency):
    data = {'Paying person': [], 'Description': [], 'Amount': [], 'Currency': [], 'Shared with': []}
    
    # Add 10 blank lines
    for _ in range(10):
        data['Paying person'].append('')
        data['Description'].append('')
        data['Amount'].append('')
        data['Currency'].append(default_currency)
        data['Shared with'].append(', '.join(people_names))
    
    df = pd.DataFrame(data)
    
    # Create a new Workbook
    book = Workbook()
    
    # Remove the default sheet created and add a new one with the DataFrame
    book.remove(book.active)
    df.to_excel(f"{doc_name}.xlsx", index=False, engine='openpyxl')
    
    # Open the workbook again using openpyxl
    book = load_workbook(f"{doc_name}.xlsx")
    
    # Add a dropdown list for the "Paying person" column
    sheet = book.active
    dv = DataValidation(type="list", formula1=f'"{", ".join(people_names)}"', allow_blank=True)
    sheet.add_data_validation(dv)
    dv.add(f'A2:A{len(people_names)+1}')  # Assuming A is the column for "Paying person"
    
    # Set the width of the columns
    for i, column_cells in enumerate(sheet.columns):
        max_length = 0
        column = [str(cell.value) for cell in column_cells]
        for cell in column:
            try:  # Necessary to avoid error on empty cells
                if len(cell) > max_length:
                    max_length = len(cell)
            except:
                pass
        adjusted_width = (max_length + 2)
        
        # Set the width of the 'Shared with' column to the width of all names
        if openpyxl.utils.get_column_letter(i+1) == 'E':  # Assuming E is the column for "Shared with"
            adjusted_width = len(', '.join(people_names)) + 2
            
        sheet.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = adjusted_width
    
    # Save the Excel file
    book.save(f"{doc_name}.xlsx")
    
    print(f"Expense template '{doc_name}.xlsx' created successfully!")

def main():
    document_name = get_document_name()
    group_size = get_group_size()
    people_names = get_person_names(group_size)
    default_currency = get_default_currency()
    create_expense_template(document_name, people_names, default_currency)

if __name__ == "__main__":
    main()
