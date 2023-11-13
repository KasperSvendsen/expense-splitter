# Expense Consolidation Scripts

This repository contains two Python scripts that help groups of people to consolidate their shared expenses:

1. `generate_template.py`: This script generates an Excel template for recording shared expenses.
2. `consolidate.py`: This script reads the filled-in Excel template and calculates how much each person owes to each other.

## Usage

### Step 1: Install the Required Packages

This project requires Python 3 and some Python packages. You can install these packages using the provided `requirements.txt` file by running the following command in your terminal:

```bash
pip install -r requirements.txt
```

### Step 2: Generate the Expense Template

Run the `generate_template.py` script first. It will ask for the following inputs:

- The name of the document (Excel file)
- The number of people in the group
- The names of the people in the group

The script will then create an Excel file with the specified name. The file will contain a table with columns for the 'Paying person', 'Description', 'Amount', and 'Shared with'. The 'Paying person' column will have a dropdown list containing the names of the people in the group. The 'Shared with' column will be pre-filled with the names of all people in the group.

### Step 3: Fill in the Expenses

Open the generated Excel file and fill in the expenses. For each expense, select the paying person from the dropdown list, enter a description and the amount, and specify with whom the expense was shared.

### Step 4: Consolidate the Debts

After all expenses have been entered, run the `consolidate.py` script. It will ask you to select the Excel file. The script will then calculate how much each person owes to each other and print the debts.