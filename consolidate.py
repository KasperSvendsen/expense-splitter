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
    for (debtor, creditor), amount in   debts.items():
        if (creditor, debtor) in final_debts:
            final_debts[(creditor, debtor)] -= amount
            if final_debts[(creditor, debtor)] < 0:
                final_debts[(debtor, creditor)] = -final_debts[(creditor, debtor)]
                del final_debts[(creditor, debtor)]
        else:
            final_debts[(debtor, creditor)] = amount

    return final_debts

def main():
    file_name = "Test.xlsx"
    debts = calculate_debts(file_name)
    for people, amount in debts.items():
        print(f"{people[0]} owes {people[1]} {amount:.2f},-")

if __name__ == "__main__":
    main()
