import numpy as np
import pandas as pd
import os 
import re

# Lese die Excel-Datei
 
data = pd.read_excel(r'C:\VMscope\Projekte\ACELEROND\Parsing\for_test_prostata.xlsx') 

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
    # Initialize a vector of zeros with length equal to COUNT_OT
    vector = np.zeros(row['COUNT_OT'], dtype=int)
    
    # If COUNT_IH is greater than zero and there are unique numbers
    if row['COUNT_IH'] > 0 and row['UNIQUE_NUMBERS_PER_ROW']:
        # Iterate over the unique numbers and ensure they are integers
        for num in row['UNIQUE_NUMBERS_PER_ROW']:
            try:
                num = int(num)  # Konvertiere num zu einer Zahl
                if 0 <= num - 1 < len(vector):  # Ensure the index is valid
                    vector[num - 1] = 1
            except ValueError:
                continue  # Überspringe fehlerhafte Werte
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


def extract_and_sort_numbers_with_points_and_update_text(text):
    if pd.notnull(text):
        text = text.replace("--", "-")
        result = []
        updated_text = text

        # Match standalone numbers (e.g., "9.") and ranges (e.g., "9 bis 12.")
        matches = re.finditer(r'(\d+)\.(?:\s*(?:-|bis)\s*(\d+)\.)?', text)
        for match in matches:
            start = int(match.group(1))
            end = int(match.group(2)) if match.group(2) else start  # Use range if "bis" is present
            expanded_range = ", ".join([f"{num}." for num in range(start, end + 1)])  # Expand the range
            
            # Replace the original range in the text with the expanded range
            original_range = match.group(0)
            updated_text = updated_text.replace(original_range, expanded_range)
            
            # Add numbers to result
            result.extend(range(start, end + 1))  # Add the numbers to the result
        
        # Sort the numbers and format them as "X."
        sorted_result = sorted(set(result))  # Deduplicate and sort
        return ", ".join([f"{num}." for num in sorted_result]), updated_text

    return None, text  # If no valid text, return original text unchanged


# Apply the function to update both the extracted numbers and the modified text
data[['Sorted_Extracted_Numbers', 'Updated_After_Mikroskopie']] = data['After_Mikroskopie'].apply(
    lambda x: pd.Series(extract_and_sort_numbers_with_points_and_update_text(x))
)


def check_karzinom_or_gleason(text, after_mikroskopie):
    if pd.notnull(text) and pd.notnull(after_mikroskopie):
        # Liste für die Ergebnisse
        updated_numbers = []

        # Liste der Blocknummern aus dem Text extrahieren
        numbers = text.split(", ")

        # Alle Blocknummern und deren Positionen finden
        block_matches = list(re.finditer(r'(\d+)\.', after_mikroskopie))

        # Wenn keine Blocknummern vorhanden sind, gibt es nichts zu tun
        if not block_matches:
            return None

        # Textsegmente zwischen den Blocknummern definieren
        segments = []
        for i in range(len(block_matches)):
            # Startposition der aktuellen Blocknummer
            start_index = block_matches[i].end()
            # Endposition der nächsten Blocknummer oder des gesamten Textes
            end_index = block_matches[i + 1].start() if i + 1 < len(block_matches) else len(after_mikroskopie)
            # Segment erstellen
            segments.append((block_matches[i].group(1), after_mikroskopie[start_index:end_index].strip()))

        # Iteriere über die Segmente und prüfe, ob "karzinom" oder "gleason" vorkommen
        for block_number, segment in segments:
            if re.search(r'\b(karzinom|gleason)\b', segment, re.IGNORECASE):
                updated_numbers.append(f"{block_number}. TRUE")
            else:
                updated_numbers.append(f"{block_number}. FALSE")

        # Bearbeite zusammenhängende Blocknummern
        for number in numbers:
            if not any(number.strip(".") == entry.split(". ")[0] for entry in updated_numbers):
                updated_numbers.append(f"{number} FALSE")

        # Ergebnisse sortieren
        results_dict = {entry.split(". ")[0]: entry for entry in updated_numbers}
        sorted_results = [results_dict.get(number.strip("."), f"{number} FALSE") for number in numbers]

        return ", ".join(sorted_results)

    return None


# Apply the function to create updated column with TRUE/FALSE for each number
data['Sorted_Extracted_Numbers_2'] = data.apply(
    lambda row: check_karzinom_or_gleason(row['Sorted_Extracted_Numbers'], row['Updated_After_Mikroskopie']),
    axis=1
)


# Define the output file path
output_file_path = 'C:\VMscope\Projekte\ACELEROND\Parsing/out.xlsx'

# Save the selected columns to an Excel file
data.to_excel(output_file_path, index=False)

