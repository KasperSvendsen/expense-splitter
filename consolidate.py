import os
import pandas as pd
import requests
import numpy as np

# Function to convert a given amount to DKK using exchange rates
def convert_to_dkk(amount, currency, exchange_rates):
    if pd.isna(currency) or currency == 'DKK':
        return amount
    try:
        exchange_rate = exchange_rates.get(currency)
        if exchange_rate is None:
            print(f"Exchange rate for {currency} not found.")
            return amount  # Return original amount if exchange rate not found
        amount_in_dkk = amount / exchange_rate
        return amount_in_dkk
    except Exception as e:
        print(f"An error occurred: {e}")
        return amount  # Return original amount on error

# Function to safely split the 'Shared with' values
def split_shared_with(x):
    if pd.isna(x) or not isinstance(x, str):
        return []  # Return empty list for non-string values
    return [person.strip() for person in x.split(', ') if person.strip()]  # Filter out empty names

# Function to load and preprocess data from an Excel file
def load_and_preprocess_data(file_name):
    df = pd.read_excel(file_name)
    
    # Drop rows with NaN in 'Paying person' or where it's an empty string
    df = df.dropna(subset=['Paying person'])
    df = df[df['Paying person'].str.strip() != '']
    
    # Ensure 'Paying person' is properly formatted as string
    df['Paying person'] = df['Paying person'].astype(str).str.strip()
    
    # Handle the 'Shared with' column safely
    df['Shared with'] = df['Shared with'].apply(split_shared_with)
    
    # Set default currency to DKK if missing
    df['Currency'] = df['Currency'].fillna('DKK')

    # Convert 'Amount' column to float to avoid dtype issues
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0).astype(float)

    # Fetch the exchange rates once
    try:
        response = requests.get("https://open.er-api.com/v6/latest/DKK")
        if response.status_code != 200:
            print("Error fetching exchange rates. Using only DKK values.")
            return df
        data = response.json()
        exchange_rates = data['rates']
    except Exception as e:
        print(f"An error occurred while fetching exchange rates: {e}")
        print("Using only DKK values.")
        return df

    # Convert non-DKK currencies to DKK for all relevant columns
    for index, row in df.iterrows():
        if row['Currency'] != 'DKK':
            # Convert the main amount
            converted_amount = convert_to_dkk(row['Amount'], row['Currency'], exchange_rates)
            if converted_amount is not None:
                df.at[index, 'Amount'] = float(round(converted_amount, 2))  # Cast to float and round to 2 decimal places

            # Convert each person's share
            for person in row['Shared with']:
                share_column = f"{person}'s share"
                if share_column in df.columns and pd.notna(row[share_column]):
                    converted_share = convert_to_dkk(row[share_column], row['Currency'], exchange_rates)
                    if converted_share is not None:
                        df.at[index, share_column] = float(round(converted_share, 2))  # Cast to float and round

            df.at[index, 'Currency'] = 'DKK'

    return df

# Function to calculate the total expenses paid by each individual
def calculate_individual_expenses(df):
    individual_expenses = {}
    for _, row in df.iterrows():
        payer = row['Paying person']
        amount = row['Amount']  # Assumed to be in DKK after preprocessing
        if payer and not pd.isna(payer):  # Skip empty payers
            individual_expenses[payer] = round(individual_expenses.get(payer, 0) + amount, 2)
    return individual_expenses

# Function to calculate the total shares owed by each individual
def calculate_total_shares(df):
    total_shares = {}
    for _, row in df.iterrows():
        shared_with = row['Shared with']
        amount = row['Amount']
        
        # Skip if shared_with is empty
        if not shared_with:
            continue
            
        explicit_shares_provided = False
        total_explicit_shares = 0

        # First, use specific shares if provided
        for person in shared_with:
            share_column = f"{person}'s share"
            if share_column in df.columns and pd.notna(row[share_column]):
                share = row[share_column]
                total_shares[person] = total_shares.get(person, 0) + share
                total_explicit_shares += share
                explicit_shares_provided = True

        # If the expense is not fully covered by specific shares, divide the remainder equally
        if not explicit_shares_provided:
            equal_share = amount / len(shared_with)
            for person in shared_with:
                total_shares[person] = total_shares.get(person, 0) + equal_share
        elif total_explicit_shares < amount:
            # Count people without explicit shares
            people_without_shares = sum(1 for person in shared_with if 
                                      f"{person}'s share" not in df.columns or 
                                      pd.isna(row[f"{person}'s share"]))
            
            # Avoid division by zero
            if people_without_shares > 0:
                remainder = amount - total_explicit_shares
                equal_share = remainder / people_without_shares
                
                for person in shared_with:
                    share_column = f"{person}'s share"
                    if share_column not in df.columns or pd.isna(row[share_column]):
                        total_shares[person] = total_shares.get(person, 0) + equal_share

    return total_shares

# Function to calculate the net balance for each individual
def calculate_net_balances(individual_expenses, total_shares):
    net_balances = {}
    for person in set(individual_expenses.keys()).union(set(total_shares.keys())):
        if not person or pd.isna(person):  # Skip empty or NaN persons
            continue
        paid_amount = individual_expenses.get(person, 0)
        share_amount = total_shares.get(person, 0)
        net_balances[person] = round(paid_amount - share_amount, 2)
    return net_balances

# Function to simplify debts between individuals
def simplify_debts(net_balances):
    # Create a list to store simplified debts
    simplified_debts = []

    # Create two lists to store people who owe money (debtors) and who are owed money (creditors)
    debtors = [(person, -balance) for person, balance in net_balances.items() if balance < 0]
    creditors = [(person, balance) for person, balance in net_balances.items() if balance > 0]

    # Iterate until all debts are settled
    while debtors and creditors:
        debtor, debt_amount = debtors.pop(0)
        creditor, credit_amount = creditors.pop(0)

        # Determine the transaction amount
        transaction_amount = min(debt_amount, credit_amount)
        simplified_debts.append((debtor, creditor, transaction_amount))

        # Update remaining amounts
        debt_amount -= transaction_amount
        credit_amount -= transaction_amount

        # Re-add debtor or creditor to the list if they still owe money or are still owed money
        if debt_amount > 0.01:  # Use small threshold to avoid floating point issues
            debtors.insert(0, (debtor, debt_amount))
        if credit_amount > 0.01:  # Use small threshold to avoid floating point issues
            creditors.insert(0, (creditor, credit_amount))

    return simplified_debts

# Function to select an Excel file for processing
def select_file():
    files = [f for f in os.listdir() if f.endswith('.xlsx')]
    if not files:
        print("No .xlsx files found in the current directory.")
        return None
    for i, file in enumerate(files, start=1):
        print(f"{i}. {file}")
    file_number = int(input("Please enter the number of the file you want to select: ")) - 1
    if file_number < 0 or file_number >= len(files):
        print("Invalid selection")
        return None
    return files[file_number]

# Main function to execute the script
def main():
    file_name = select_file()
    if file_name is None:
        return

    print(f"Processing {file_name}...")
    df = load_and_preprocess_data(file_name)
    
    if df.empty:
        print("No valid data found in the file after preprocessing.")
        return
        
    individual_expenses = calculate_individual_expenses(df)
    total_shares = calculate_total_shares(df)
    net_balances = calculate_net_balances(individual_expenses, total_shares)
    simplified_debts = simplify_debts(net_balances)

    print("\nNet Balances in DKK:")
    for person, balance in net_balances.items():
        if balance > 0:
            print(f"{person} is owed {balance:.2f} DKK")
        elif balance < 0:
            print(f"{person} owes {abs(balance):.2f} DKK")
        else:
            print(f"{person} is settled up")

    # Print simplified debts
    print("\nSimplified Debts:")
    for debtor, creditor, amount in simplified_debts:
        print(f"{debtor} owes {creditor} {amount:.2f} DKK")

if __name__ == "__main__":
    main()