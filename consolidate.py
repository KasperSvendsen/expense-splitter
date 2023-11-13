import os
import pandas as pd

def calculate_debts(file_name):
    # Read the Excel file
    df = pd.read_excel(file_name)

    # Initialize a dictionary to store the debts
    debts = {}

    # Iterate over the rows of the DataFrame
    for _, row in df.iterrows():
        # Get the paying person, amount, and the people with whom the expense is shared
        paying_person = row['Paying person']
        amount = row['Amount']
        shared_with = row['Shared with'].split(', ')

        # Calculate the amount each person owes for this expense
        owed_amount = amount / len(shared_with)

        # Update the debts
        for person in shared_with:
            if person != paying_person:
                if (person, paying_person) not in debts:
                    debts[(person, paying_person)] = owed_amount
                else:
                    debts[(person, paying_person)] += owed_amount

    # Consolidate debts
    final_debts = {}
    for (debtor, creditor), amount in debts.items():
        if (creditor, debtor) in final_debts:
            final_debts[(creditor, debtor)] -= amount
            if final_debts[(creditor, debtor)] < 0:
                final_debts[(debtor, creditor)] = -final_debts[(creditor, debtor)]
                del final_debts[(creditor, debtor)]
        else:
            final_debts[(debtor, creditor)] = amount

    return final_debts

def main():
    # Get a list of all .xlsx files in the current directory
    files = [f for f in os.listdir() if f.endswith('.xlsx')]

    # If there are no .xlsx files, print a message and return
    if not files:
        print("No .xlsx files found in the current directory.")
        return

    # Print the available files
    for i, file in enumerate(files, start=1):
        print(f"{i}. {file}")

    # Ask the user to select a file
    file_number = int(input("Please enter the number of the file you want to select: ")) - 1
    file_name = files[file_number]

    # Calculate and print the debts
    debts = calculate_debts(file_name)

    # Sort the debts by amount owed
    sorted_debts = sorted(debts.items(), key=lambda x: x[1], reverse=True)

    # Group the debts by debtor
    grouped_debts = {}
    for people, amount in sorted_debts:
        debtor = people[0]
        if debtor not in grouped_debts:
            grouped_debts[debtor] = []
        grouped_debts[debtor].append((people[1], amount))

    # Print the debts
    print("")
    for debtor, debts in grouped_debts.items():
        print(f"{debtor} owes:")
        for creditor, amount in debts:
            print(f"  {creditor}: {amount:.2f},-")
        print("")

if __name__ == "__main__":
    main()
