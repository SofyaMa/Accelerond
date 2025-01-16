import numpy as np
import pandas as pd
import os 
import re

# Lese die Excel-Datei
 
data = pd.read_excel('/home/marchenko/accelerond/prostata-2023.xlsx') 

#Extract Block Nrs
data['NUMBERS_AFTER_P'] = data['NAME_LABEL_IH'].apply(
    lambda x: re.findall(r'P(\d+)', x) if isinstance(x, str) else []
)
#Get unique Block Nrs
data['UNIQUE_NUMBERS_PER_ROW'] = data['NUMBERS_AFTER_P'].apply(lambda x: sorted(set(x)))

#Für Ethikantrag
#Anzahl Objektträger in Datensatz 
sum_count_ot = data['COUNT_OT'].sum()
# Anzahl IHC 
sum_total_elements = sum(len(row) for row in data['UNIQUE_NUMBERS_PER_ROW'])
# Compute the percentage
percentage = (sum_total_elements / sum_count_ot ) * 100
# Print the result
print(f"Percentage: {percentage:.2f}%")
# Output the results
print("Sum of COUNT_OT:", sum_count_ot)
print("Sum of unique numbers across all UNIQUE_NUMBERS_PER_ROW:", sum_total_elements)

# Create the new column as a vector of length COUNT_IH for each row
def create_vector(row):
    # Initialize a vector of zeros with length equal to COUNT_IH
    vector = np.zeros(row['COUNT_OT'], dtype=int)
    
    # If COUNT_IH is greater than zero and there are unique numbers
    if row['COUNT_IH'] > 0 and row['UNIQUE_NUMBERS_PER_ROW']:
        # Iterate over the unique numbers and set corresponding indices to 1
        for num in row['UNIQUE_NUMBERS_PER_ROW']:
            if 0 <= num - 1 < len(vector):  # Ensure the index is valid
                vector[num - 1] = 1
    return vector

# Apply the function to create the new column
data['VECTOR_BASED_ON_COUNT_IH'] = data.apply(create_vector, axis=1)

# Display the relevant columns for verification
data[['COUNT_IH', 'UNIQUE_NUMBERS_PER_ROW']].head()


# Function to extract unique IH tests from each row
def extract_unique_ihcs(row):
    if pd.notnull(row):
        split_tests = row.split("###")
        unique_tests = pd.Series(split_tests).str.extract(r'(\w+-IH-\w+)')[0].dropna().unique()
        return ", ".join(unique_tests)
    return None

# Apply the function to create a new column with unique IH tests
data['Unique_IHCs'] = data['NAME_LABEL_IH'].apply(extract_unique_ihcs)

# Extract text after "Mikroskopie" and save in a new column
def extract_after_mikroskopie(text):
    if pd.notnull(text):
        split_text = text.split("Mikroskopie", 1)
        if len(split_text) > 1:
            return split_text[1].strip()
    return None

data['After_Mikroskopie'] = data['DOCTEXT_LAST'].apply(extract_after_mikroskopie)


def extract_and_sort_numbers_with_points_updated(text):
    if pd.notnull(text):
        result = []
        # Find standalone numbers and ranges with "-" or "bis"
        matches = re.findall(r'(\d+)\.?-?bis?-?(\d+)?\.', text)
        for match in matches:
            if match[1]:  # If there's a range
                # Generate the range of numbers and append
                result.extend([int(i) for i in range(int(match[0]), int(match[1]) + 1)])
            else:  # If it's a standalone number
                result.append(int(match[0]))
        # Sort the numbers and format them as "X."
        sorted_result = sorted(result)
        return ", ".join([f"{num}." for num in sorted_result])
    return None

# Apply the function to the 'After_Mikroskopie' column
data['Sorted_Extracted_Numbers'] = data['After_Mikroskopie'].apply(extract_and_sort_numbers_with_points)


def check_karzinom_or_gleason(text, after_mikroskopie):
    if pd.notnull(text) and pd.notnull(after_mikroskopie):
        numbers = text.split(", ")
        updated_numbers = []
        for number in numbers:
            # Escape the number for a regex match
            number_match = re.escape(number)
            # Check for "karzinom" or "Gleason" after the number
            if re.search(f"{number_match}.*?(karzinom|gleason)", after_mikroskopie, re.IGNORECASE):
                updated_numbers.append(f"{number} TRUE")
            else:
                updated_numbers.append(f"{number} FALSE")
        return ", ".join(updated_numbers)
    return None

# Apply the function to create updated column with TRUE/FALSE for each number
data['Sorted_Extracted_Numbers'] = data.apply(
    lambda row: check_karzinom_or_gleason(row['Sorted_Extracted_Numbers'], row['After_Mikroskopie']),
    axis=1
)


# Define the output file path
output_file_path = '/home/marchenko/accelerond/2023_table.xlsx'

# Save the selected columns to an Excel file
data.to_excel(output_file_path, index=False)

