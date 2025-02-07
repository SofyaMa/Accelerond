import pandas as pd
import os 
import re
# Lese die Excel-Datei
df = pd.read_excel('accelerond/query_accelerond_test.xlsx') 
df2 = pd.read_excel('accelerond/qu__prostata-2024-100__SM.xlsx') 

# Funktion zum Parsen der Abschnitte zwischen den Zahlen
def extract_psa_score(text):
    # Adjusted pattern to recognize "PSA", "PSA-Wert", "PSA aktuell", "PSA-Anstieg", and "Anstieg...auf" with numbers
    psa_pattern = r'(?:PSA[\s-]*(?:Wert|aktuell|Anstieg)?)[\s:]*([0-9]+(?:,[0-9]+)?)|Anstieg.*?auf\s([0-9]+(?:,[0-9]+)?)'
    
    # Search for all matches and select the last one
    matches = re.findall(psa_pattern, text)
    if matches:
        # Retrieve the last matching value, convert comma to period, and parse as float
        last_match = matches[-1]
        # Either from PSA score or the Anstieg section
        score_str = last_match[0] or last_match[1]
        return float(score_str.replace(',', '.'))
    return None
  
  
# Anwenden der Funktion auf die DOCTEXT_LAST-Spalte
df['PSA_Score'] = df['DOCTEXT_LAST'].apply(extract_psa_score)
df['PSA_Found'] = df['DOCTEXT_LAST'].apply(lambda x: 'PSA' in x)

df2['PSA_Score'] = df2['(Kein Spaltenname)'].apply(extract_psa_score)
df2['PSA_Found'] = df2['(Kein Spaltenname)'].apply(lambda x: 'PSA' in x)
df2.to_excel('accelerond/psa_score_not_detected.xlsx', index=False)
# Ausgabe des DataFrame
print(df2[[ 'PSA_Score',"PSA_Found"]])

# Count rows where 'PSA_Score' is NaN and 'PSA_Found' is False
count = df2[df2['PSA_Score'].isna() & (df2['PSA_Found'] == True)].shape[0] #18/100
# Filter rows where 'PSA_Score' is NaN and 'PSA_Found' is False
filtered_df = df2[df2['PSA_Score'].isna() & (df2['PSA_Found'] == False)]

# Print the 'DOC_TEXT' column of the filtered rows
print(filtered_df['DOC_TEXT'])

print(f"The number of rows with no PSA_Score and PSA_Found as True is: {count}")

def parse_karzinom(text):
    # Dictionary to store results
    results = {}
    
    # Split the text into lines
    lines = text.split('\n')
    
    # Regular expression to match the diagnosis lines
    pattern = r'(\d+)\.\s*(Azinäres Adenokarzinom der Prostata|Tumorfreie Prostatastanzzylinder|Prostatadrüsen- und Stromagewebe mit hochgradiger prostatischer intraepithelialer Neoplasie)'
    
    for line in lines:
        match = re.match(pattern, line)
        if match:
            number = int(match.group(1))
            diagnosis = match.group(2)
            if 'Adenokarzinom' in diagnosis:
                results[number] = 'Karzinom'
            elif 'Tumorfreie' in diagnosis:
                results[number] = 'Kein Karzinom'
            elif 'hochgradiger prostatischer intraepithelialer Neoplasie' in diagnosis:
                results[number] = 'Hochgradige Neoplasie'
            else:
                results[number] = 'Nicht angegeben'
    
    # Ensure all numbers from 1 to 11 are present
    for i in range(1, 12):
        if i not in results:
            results[i] = 'Nicht angegeben'
    return results

# Assuming your DataFrame is named 'df' and the column with the text is named 'biopsy_text'
df['Parsed_Sections'] = df['DOCTEXT_LAST'].apply(parse_karzinom)


# Create separate columns for each number
for i in range(1, 12):
    df[f'Sample_{i}'] = df['Parsed_Sections'].apply(lambda x: x.get(i, 'Nicht angegeben'))

# Drop the intermediate 'parsed_results' column if you don't need it
df = df.drop(columns=['Parsed_Sections'])

# Display the results
print(df[[f'Sample_{i}' for i in range(1, 12)]])
