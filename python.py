import os
import pandas as pd
import re
from bs4 import BeautifulSoup

# --- Configuration ---
FOLDER_NAME = 'allen_results'
OUTPUT_FILE = 'Allen_Dark_Analysis.xlsx'

def parse_allen_result(file_path):
    """Parses a single HTML file and returns a dictionary of data."""
    with open(file_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    filename = os.path.basename(file_path)
    # Initialize with Date column (empty placeholder for new files)
    row_data = {'Date': '', 'File Name': filename}

    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 1. Global Info
    title_div = soup.find('div', {'data-testid': 'test-title'})
    if title_div:
        full_text = " ".join(title_div.get_text(separator=" ").split())
        row_data['Test Name'] = full_text.replace("Result:", "").strip()

    # Scores
    score_matches = re.findall(r'text-3xl lg:text-4xl leading-10 lg:leading-12">(\d+)</div>', html_content)
    if len(score_matches) >= 2:
        row_data['Glb Score'] = int(score_matches[0])
        row_data['Max Marks'] = int(score_matches[1])

    # Percentile
    percentile_match = re.search(r'(\d+\.\d+)\s*Percentile', html_content)
    if percentile_match:
        row_data['Percentile'] = float(percentile_match.group(1))

    # Global Stats
    correct_match = re.search(r'text-success">(\d+)</p>', html_content)
    if correct_match:
        row_data['Glb C'] = int(correct_match.group(1))

    incorrect_match = re.search(r'text-error">(\d+)</p>', html_content)
    if incorrect_match:
        row_data['Glb W'] = int(incorrect_match.group(1))

    unattempted_match = re.search(r'Unattempted.*?font-bold.*?text-default-body">(\d+)</p>', html_content, re.DOTALL)
    if unattempted_match:
        row_data['Glb U'] = int(unattempted_match.group(1))

    # 2. Subject Stats
    subject_matches = re.finditer(r'(PART-\d+\s*:\s*[A-Z]+)', html_content)
    subj_map = {"PHYSICS": "Phy", "CHEMISTRY": "Chem", "MATHEMATICS": "Math"}

    for match in subject_matches:
        raw_name = match.group(1)
        short_subj = "Unknown"
        for key, val in subj_map.items():
            if key in raw_name:
                short_subj = val
                break

        start_idx = match.end()
        search_window = html_content[start_idx:start_idx+1500]
        numbers = re.findall(r'class="col-span-2 text-center">(\d+)</div>', search_window)
        
        if len(numbers) >= 3:
            s_score = int(numbers[0])
            s_correct = int(numbers[1])
            s_incorrect = int(numbers[2])
            s_unattempted = 25 - (s_correct + s_incorrect)
            
            row_data[f'{short_subj} S'] = s_score
            row_data[f'{short_subj} C'] = s_correct
            row_data[f'{short_subj} W'] = s_incorrect
            row_data[f'{short_subj} U'] = s_unattempted

    return row_data

def calculate_accuracy(df):
    """Calculates accuracy percentages for Global and Subjects."""
    # Global Accuracy
    if 'Glb C' in df.columns and 'Glb W' in df.columns:
        attempts = df['Glb C'] + df['Glb W']
        # Avoid division by zero
        df['Glb Acc%'] = df.apply(lambda x: (x['Glb C'] / (x['Glb C'] + x['Glb W']) * 100) if (x['Glb C'] + x['Glb W']) > 0 else 0, axis=1)

    # Subject Accuracy
    for subj in ['Phy', 'Chem', 'Math']:
        c_col = f'{subj} C'
        w_col = f'{subj} W'
        acc_col = f'{subj} Acc%'
        
        if c_col in df.columns and w_col in df.columns:
             df[acc_col] = df.apply(lambda x: (x[c_col] / (x[c_col] + x[w_col]) * 100) if (x[c_col] + x[w_col]) > 0 else 0, axis=1)
            
    return df

def apply_styling(df, output_file):
    """Writes the DataFrame to Excel with Dark Mode and Conditional Formatting."""
    
    # Organize Columns
    cols = ['Date', 'File Name', 'Test Name', 'Glb Score', 'Percentile', 'Glb Acc%']
    cols += ['Glb C', 'Glb W', 'Glb U']
    for subj in ['Phy', 'Chem', 'Math']:
        cols += [f'{subj} S', f'{subj} Acc%', f'{subj} C', f'{subj} W', f'{subj} U']
    
    # Keep only columns that exist (in case a subject is missing in a test)
    final_cols = [c for c in cols if c in df.columns]
    df = df[final_cols]

    # Initialize Writer
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Define Formats
    dark_bg = workbook.add_format({'bg_color': '#000000', 'font_color': '#FFFFFF', 'border': 1, 'border_color': '#444444'})
    header_fmt = workbook.add_format({'bg_color': '#222222', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center'})
    percent_fmt = workbook.add_format({'num_format': '0.0', 'bg_color': '#000000', 'font_color': '#FFFFFF'})
    
    # Apply Dark Background to Data Area
    (max_row, max_col) = df.shape
    worksheet.set_column(0, max_col-1, 10, dark_bg) # Set default width 10 and dark format
    
    # Widen Title Columns
    worksheet.set_column('B:C', 25) 
    
    # Write Headers
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # Conditional Formatting
    for i, col_name in enumerate(df.columns):
        col_letter = chr(ord('A') + i)
        if i >= 26: col_letter = 'A' + chr(ord('A') + (i-26)) # Basic support for AA, AB...
        
        rng = f"{col_letter}2:{col_letter}{max_row+1}"
        
        # Green (High) -> Red (Low) [Score, Accuracy, Correct]
        if any(x in col_name for x in ['Score', 'Percentile', 'Acc%', ' C', ' S']):
            worksheet.conditional_format(rng, {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'})
            if 'Acc%' in col_name or 'Percentile' in col_name:
                 worksheet.set_column(i, i, 12, percent_fmt)

        # Red (High) -> Green (Low) [Wrong, Incorrect]
        elif ' W' in col_name:
            worksheet.conditional_format(rng, {'type': '3_color_scale', 'min_color': '#63BE7B', 'mid_color': '#FFEB84', 'max_color': '#F8696B'})

    writer.close()
    print(f"Successfully updated and saved to '{output_file}'")

def update_excel_sheet():
    # 1. Check for existing data
    if os.path.exists(OUTPUT_FILE):
        try:
            print("Found existing Excel file. Reading data...")
            df_existing = pd.read_excel(OUTPUT_FILE)
            existing_files = df_existing['File Name'].tolist()
        except Exception as e:
            print(f"Error reading existing file: {e}")
            existing_files = []
            df_existing = pd.DataFrame()
    else:
        print("No existing Excel file found. Creating new...")
        df_existing = pd.DataFrame()
        existing_files = []

    # 2. Scan folder for NEW files
    if not os.path.exists(FOLDER_NAME):
        print(f"Error: Folder '{FOLDER_NAME}' not found.")
        return

    all_files = [f for f in os.listdir(FOLDER_NAME) if f.endswith('.html')]
    new_files = [f for f in all_files if f not in existing_files]

    if not new_files:
        print("No new HTML files found. Excel is up to date.")
        return

    print(f"Found {len(new_files)} new files. Parsing...")

    # 3. Parse New Files
    new_results = []
    for file in new_files:
        try:
            data = parse_allen_result(os.path.join(FOLDER_NAME, file))
            new_results.append(data)
            print(f" -> Parsed: {file}")
        except Exception as e:
            print(f" -> Error parsing {file}: {e}")

    if new_results:
        df_new = pd.DataFrame(new_results)
        df_new = calculate_accuracy(df_new)
        
        # 4. Concatenate Old + New
        if not df_existing.empty:
            # Align columns and append
            df_final = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_final = df_new
            
        # 5. Save with Styling (Overwrites file but keeps old data + adds new)
        apply_styling(df_final, OUTPUT_FILE)

if __name__ == "__main__":
    update_excel_sheet()