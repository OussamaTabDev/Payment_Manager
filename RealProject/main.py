import openpyxl
from openpyxl.styles import Font, Alignment
import pandas as pd

# Load the Excel file (or CSV)
parents_df = pd.read_excel("parents_payments.xlsx" , header=None,)
kids_df = pd.read_excel("kids_list.xlsx",header=None,)
# Split the first 3 rows
kids_first_rows = kids_df.iloc[:3]   # first 3 rows
kids_df = kids_df.iloc[3:]    # all rows after the first 3
kids_df = kids_df.reset_index(drop=True)

kids_last_rows = kids_df
# refinning kids_df , parents_df columns names
parents_df.columns = [
    "Account_Number",            # Auftragskonto
    "Booking_Date",              # Buchungstag
    "Value_Date",                # Valutadatum
    "Transaction_Text",          # Buchungstext
    "Usage_Purpose",             # Verwendungszweck
    "parent_name",        # Beguenstigter/Zahlungspflichtiger
    "Account_or_IBAN",           # Kontonummer/IBAN
    "BIC_SWIFT_Code",            # BIC (SWIFT-Code)
    "Amount",                    # Betrag
    "Currency",                  # Waehrung
    "Info"                       # Info
]

kids_df.columns = [
    "kid_id",'kid_name', 'parent_name', 
    '9', '10', '11', '12', '1', '2', '3',
    '4', '5', '6', '7', '8','9_next',
    '10_next', '11_next', '12_next', '1_next',
    'class' , 'priceOn','book_taken','nabil_liste'
    ]

backup_kids_df = kids_df.copy()



print("Data loaded successfully.")

print("Showing data samples:")
print("Parents Payments Data:")
print(parents_df.head())
print("\nKids Full Names Data:")
print(kids_df.head())

# Choose mode: "test" or "prod"
mode = "prod"  # change to "test" when testing

if mode == "test":
    # Limit to first 100 rows
    kids_df = kids_df.head(100)
    print(f"🧪 Testing mode: Limited to {len(kids_df)} rows.")
else:
    # Production mode: stop at first empty kid_id
    if "kid_id" in kids_df.columns:
        # Find first row where kid_id is empty
        stop_index = kids_df["kid_id"].isna().idxmax() if kids_df["kid_id"].isna().any() else None
        
        if stop_index is not None and stop_index > 0:
            kids_last_rows = kids_df.iloc[stop_index:]
            kids_df = kids_df.iloc[:stop_index]
            print(f"✅ Production mode: Stopped at first empty kid_id (row {stop_index}).")
        else:
            print("✅ Production mode: No missing kid_id found. Using all rows.")
    else:
        print("⚠️ Column 'kid_id' not found in kids_df!")



# kids_df.to_excel("kids_parents_from_kids_debug.xlsx", index=False)
# find parents names in kids full names
def find_kids_of_parrents(parents_df, kids_df):
    #step 1: find distinct parents names
    # 1. Distinct parents from parents_df
    distinct_parents = parents_df['parent_name'].dropna().unique()

    # 2. Distinct kid-parent pairs from kids_df
    # Example DataFrame
    kids_parents_from_kids = kids_df[['kid_id','kid_name', 'parent_name' , 'class']]

    # Extract content inside parentheses to new column
    kids_parents_from_kids['phone_number'] = kids_parents_from_kids['parent_name'].str.extract(r'\(([^)]*)\)')

    # Remove parentheses and content from parent_name
    kids_parents_from_kids['parent_name'] = kids_parents_from_kids['parent_name'].str.replace(r'\s*\([^\)]*\)', '', regex=True)

    # Remove leading/trailing spaces
    kids_parents_from_kids['parent_name'] = kids_parents_from_kids['parent_name'].str.strip()
    kids_parents_from_kids.to_excel("kids_parents_from_kids_debug21.xlsx", index=False)

    # Drop empty strings
    empty_kids_parents_from_kids = kids_parents_from_kids[kids_parents_from_kids['parent_name'].isna()]
    kids_parents_from_kids = kids_parents_from_kids[kids_parents_from_kids['parent_name'].notna()]
    other_distinct_parents = kids_parents_from_kids['parent_name'].dropna().unique()
    print(f"Distinct parents from parents list: {distinct_parents}")
    
    # completing empty parent names from distinct parents
    # Loop over the rows with missing parent names
    for index, row in empty_kids_parents_from_kids.iterrows():
        kid_name = row['kid_name']
        matched_parent = None

        # First, check other distinct parents
        for parent in other_distinct_parents:
            if not isinstance(parent, str) or not parent.strip():
                continue  # skip empty or non-string parents
            last_name_parent = parent.split()[-1]
            if kid_name and kid_name.split() and last_name_parent.lower() == kid_name.split()[0].lower():
                matched_parent = parent
                break

        # If no match found, check main distinct parents
        if not matched_parent:
            for parent in distinct_parents:
                if not isinstance(parent, str) or not parent.strip():
                    continue
                last_name_parent = parent.split()[-1]
                if kid_name and kid_name.split() and last_name_parent.lower() == kid_name.split()[0].lower():
                    matched_parent = parent
                    break



        # Update the main DataFrame if a match was found
        if matched_parent:
            empty_kids_parents_from_kids.at[index, 'parent_name'] = matched_parent
            print(f"Completed missing parent name for kid '{kid_name}' with parent '{matched_parent}'")
        else:
            print(f"No matching parent found for kid '{kid_name}' with missing parent name will put the original")
            # Make sure both DataFrames have 'kid_id' column
            for kid_id in empty_kids_parents_from_kids['kid_id']:
                # Find the row in empty_kids_parents_from_kids
                idx = empty_kids_parents_from_kids.index[empty_kids_parents_from_kids['kid_id'] == kid_id][0]
                
                # Only update if parent_name is empty
                if pd.isna(empty_kids_parents_from_kids.at[idx, 'parent_name']) or empty_kids_parents_from_kids.at[idx, 'parent_name'] == '':
                    # Get the parent_name from backup_kids_df using kid_id
                    parent_name = backup_kids_df.loc[backup_kids_df['kid_id'] == kid_id, 'parent_name'].values
                    if len(parent_name) > 0:
                        empty_kids_parents_from_kids.at[idx, 'parent_name'] = str(parent_name[0])


    empty_kids_parents_from_kids.to_excel("kids_parents_from_kids_debug23.xlsx", index=False)
    # After filling missing names, sort the DataFrame by kid_id
    # kids_parents_from_kids = kids_parents_from_kids.sort_values(by='kid_id').reset_index(drop=True)
    # replaceing parents names in  kids_parents_from_kids with distinct parents names if match found or in
    for index, row in kids_parents_from_kids.iterrows():
        kid_name = row['kid_name']
        current_parent_name = row['parent_name']
        matched_parent = None
        for parent in distinct_parents:
            try:
                parent_str = str(parent)
                current_parent_name_str = str(current_parent_name)
                if parent_str in current_parent_name_str or current_parent_name_str in parent_str and len(parent_str) > 3:
                    matched_parent = parent
                    break
            except TypeError:
                print("TypeError encountered!")
                print(f"parent: {parent} ({type(parent)})")
                print(f"current_parent_name: {current_parent_name} ({type(current_parent_name)})")
                continue


        if matched_parent:
            kids_parents_from_kids.at[index, 'parent_name'] = matched_parent
            print(f"Replaced parent name for kid '{kid_name}' with parent New:'{matched_parent}' , Old:'{current_parent_name}'")
        else:
            print(f"No matching parent found to replace for kid '{kid_name}' with parent name '{current_parent_name}'")

    # export excel file for kids_parents_from_kids for debugging
    # Concatenate the two DataFrames
    combined = pd.concat([kids_parents_from_kids, empty_kids_parents_from_kids], ignore_index=True)

    # Sort by 'kid_id'
    combined = combined.sort_values(by='kid_id', ignore_index=True)
    combined.to_excel("kids_parents_from_kids_debug.xlsx", index=False)
    # print(f"Distinct parents from kids list: {kids_parents_from_kids['parent_name'].unique()}")

    # distinct_kids = kids_df['kid_name'].dropna().unique()
    # # put each parents name as key and value as list of kids names that contain the parents name, and it's class
    return combined

def get_parent_kid_map(combined_df):
    parent_kid_map = {}

    # Step 1: Remove rows with empty or missing parent_name
    df_valid = combined_df[combined_df['parent_name'].notna() & (combined_df['parent_name'].str.strip() != '')]

    # Step 2: Group by parent_name and filter those with >=2 kids
    parent_groups = df_valid.groupby('parent_name').filter(lambda x: len(x) >= 2)

    # Step 3: Build the nested dictionary
    result = {}
    for parent, group in parent_groups.groupby('parent_name'):
        result[parent] = dict(zip(group['kid_name'], group['class']))

    print(result)
    return result


data_map =  get_parent_kid_map(find_kids_of_parrents(parents_df, kids_df))

def getting_mount_from_string(amount_str):
    try:
        print(f"Extracting amount from string: {amount_str}")
        amount = int(''.join(filter(str.isdigit, amount_str)))
        return amount
    except:
        return 0.0

def calculate_months_paid(parents_df=parents_df):
    # Ensure 'Amount' is numeric
    parents_df['Amount'] = pd.to_numeric(parents_df['Amount'], errors='coerce')
    
    # Skip the first ROW (index 0)
    df_filtered = parents_df.iloc[1:].copy()
    
    parents_mount = dict(zip(df_filtered['parent_name'], df_filtered['Amount']))
    print(parents_mount)
    print("Calculating months paid for each kid...")
    return parents_mount


amount_map = calculate_months_paid(parents_df)


# limiting row of kids_df
# infos
monthly_fee_per_kid_A = 25.0  # Example monthly fee per kid 25 for A5,.. and 15 for B0,...
monthly_fee_per_kid_B = 15.0  # Example monthly fee per kid 25 for A5,.. and 15 for B0,...
A5_names = ["A5","A6","A7","A8","A9","A10","A11","A12"] # premair school, collige names 
B0_names = ["B0","B1","B2","B3","G1","G2"] # before primary school names , Sunday school

def get_monthly_fee_for_class(class_name):
    if class_name in A5_names:
        return monthly_fee_per_kid_A
    
    if class_name in B0_names:
        return monthly_fee_per_kid_B
    
    return monthly_fee_per_kid_A  # default to A fee



def determine_status_and_color(months_paid, monthly_fee, allocated_amount, class_name):
    # Handle edge cases first
    if monthly_fee <= 0:
        return "Not yet registered", "#ffffff"  # white or undefined

    # Check common patterns
    if allocated_amount == 0:
        return "Nothing paid.", "#ff0000"
    
    # Check if fully paid (at least 1 full month)
    if months_paid >= 1.0:
        return "Fully paid.", "#92d050"
    
    # Special case: G1/G2 paying €15 instead of €25 → assuming A5 = G1/G2
    if class_name.strip() == 'A5' and abs(allocated_amount - 15) < 0.01:
        return "G1 and G2 paid €15 instead of €25.", "#ffff00"

    # Check partial transfers: only 10, 15, or 20 instead of 25
    if class_name.strip() == 'A5':
        if allocated_amount in [10, 15, 20]:
            return "Transfers only €10, €15, or €20 instead of €25.", "#c65911"
    elif class_name.strip() == 'B0':
        if allocated_amount in [10, 15]:  # since full is 20
            return "Transfers only €10 or €15 instead of €20.", "#c65911"

    # Default partial payment
    return f"Partial payment: {allocated_amount:.2f}€ ({months_paid:.2f} months)", "#ffc000"



'''
#c65911 :Transfers only €10, €15, or €20 instead of €25.
#ff0000 :Nothing paid.
#92d050 :Fully paid.
#ffff00 : G1 and G2 paid €15 instead of €25.
#f4b084 :Unenrolled kids.
# '''

def calculate_kid_payments(data_map, amount_map):
    kid_payment_status = {}

    for parent, kids in data_map.items():
        total_amount = float(amount_map.get(parent, 0.0))
        #convert to float
        total_amount = total_amount if isinstance(total_amount, float) else 0
        print(f"\nParent: {parent}, Total Amount Paid: €{total_amount}")

        # Get monthly fees for all kids
        kid_list = list(kids.items())  # preserve order
        kid_fees = {}
        total_monthly_fee = 0.0
        for kid, cls in kid_list:
            fee = get_monthly_fee_for_class(cls)
            kid_fees[kid] = fee
            total_monthly_fee += fee

        print(f"Total Monthly Fee for {parent}: €{total_monthly_fee}")

        # Handle zero fee case
        if total_monthly_fee <= 0:
            for kid_name, class_name in kid_list:
                status_msg, color = determine_status_and_color(0, 0, 0, class_name)
                kid_payment_status[kid_name] = {
                    'parent': parent,
                    'class': class_name.strip(),
                    'monthly_fee': 0.0,
                    'allocated_amount': 0.0,
                    'months_paid': 0.0,
                    'status': status_msg,
                    'color': color,
                    'extras': 0.0,
                    'extras_color': color
                }
            continue

        # Step 1: Compute how many FULL months are covered in total
        print(f"  → Total Amount: €{total_amount}, Total Monthly Fee: €{total_monthly_fee}")
        full_months_total = int(total_amount // total_monthly_fee)
        remainder = total_amount - (full_months_total * total_monthly_fee)

        # Step 2: Base allocation = full_months_total * individual fee
        base_allocations = {
            kid_name: full_months_total * kid_fees[kid_name]
            for kid_name, _ in kid_list
        }

        # Step 3: Assign entire remainder to the FIRST child only
        allocations = {}
        for i, (kid_name, _) in enumerate(kid_list):
            if i == 0:
                allocations[kid_name] = base_allocations[kid_name] + remainder
            else:
                allocations[kid_name] = base_allocations[kid_name]

        # Step 4: Build result per kid
        for kid_name, class_name in kid_list:
            monthly_fee = kid_fees[kid_name]
            allocated = allocations[kid_name]

            # Compute months paid (for status logic)
            months_paid = allocated / monthly_fee if monthly_fee > 0 else 0.0

            # Determine main status and color
            status_msg, color = determine_status_and_color(
                months_paid, monthly_fee, allocated, class_name
            )

            # Compute extras (amount beyond last full month)
            if monthly_fee > 0:
                full_months_for_kid = int(allocated // monthly_fee)
                extras = allocated - (full_months_for_kid * monthly_fee)
                # Determine extras color: orange if partial, green if zero
                extras_color = "#ffc000" if extras > 1e-2 else "#92d050"
                extras = round(extras, 2)
            else:
                extras = 0.0
                extras_color = color

            # Final rounding
            allocated = round(allocated, 2)
            months_paid = round(months_paid, 2)
            monthly_fee = round(monthly_fee, 2)

            kid_payment_status[kid_name] = {
                'parent': parent,
                'class': class_name.strip(),
                'monthly_fee': monthly_fee,
                'allocated_amount': allocated,
                'months_paid': months_paid,
                'status': status_msg,
                'color': color,
                'extras': extras,
                'extras_color': extras_color
            }

            print(f"  → {kid_name}: €{allocated:.2f} allocated → {months_paid:.2f} months → {status_msg}")
            if extras > 0:
                print(f"      (Extras: €{extras:.2f})")

    return kid_payment_status

print("\nCalculating kid payment statuses...-------------------------")
print(calculate_kid_payments(data_map, amount_map))
