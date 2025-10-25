import openpyxl
from openpyxl.styles import Font, Alignment
import pandas as pd

# Load the Excel file (or CSV)
parents_df = pd.read_excel("parents_payments.xlsx")
kids_df = pd.read_excel("kids_list.xlsx")

monthly_fee_per_kid = 20.0  # Example monthly fee per kid

print("Data loaded successfully.")

print("Showing data samples:")
print("Parents Payments Data:")
print(parents_df.head())
print("\nKids Full Names Data:")
print(kids_df.head())


# find parents names in kids full names
def find_kids_of_parrents(parents_df, kids_df):
    distinct_parents = parents_df['parents_name'].dropna().unique()
    distinct_kids = kids_df['kid_name'].dropna().unique()
    # put each parents name as key and value as list of kids names that contain the parents name
    parent_kid_map = {}
    for parent in distinct_parents:
        last_name_parent = parent.split()[-1]
        # print(f"Parent: {parent}, Last Name: {last_name_parent}")
        matched_kids = [kid for kid in distinct_kids if last_name_parent == kid.split()[-1]]
        if matched_kids:
            parent_kid_map[parent] = matched_kids
            # print(f"Parent: {parent} has kids: {matched_kids}")
    return parent_kid_map

def getting_mount_from_string(amount_str):
    try:
        amount = int(''.join(filter(str.isdigit, amount_str)))
        return amount
    except:
        return 0.0

def calculate_months_paid(parents_df = parents_df, parent_kid_map = {} , monthly_fee_per_kid=20.0):
    parents_mount = dict(zip(parents_df['parents_name'], parents_df['amount'].apply(getting_mount_from_string)))
    print(parents_mount)
    print("Calculating months paid for each kid...")
    kids_months_paid = {}
    
    return kids_months_paid

print("Finding distinct parents...")
parent_kid_map = find_kids_of_parrents(parents_df, kids_df)

print("Print parents to their kids:")
for parent, kids in parent_kid_map.items():
    print(f"Parent: {parent} -> Kids: {', '.join(kids)}")

calculate_months_paid(parents_df , parent_kid_map , monthly_fee_per_kid )