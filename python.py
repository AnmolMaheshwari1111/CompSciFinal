import os
import pandas as pd
import re
from bs4 import BeautifulSoup
import email.parser
import email.policy
import io
import mimetypes

# --- Configuration ---
FOLDER_NAME = 'allen_results'
OUTPUT_FILE = 'Allen_Dark_Analysis.xlsx'
ALLOWED_EXTENSIONS = ('.html', '.mht', '.mhtml') 

# --- MHTML Utility Function ---
def read_mhtml(file_path):
    """
    Reads an MHTML file and extracts the main HTML content.
    MHTML is a multipart/related MIME message.
    """
    try:
        # Read the file content as bytes
        with open(file_path, 'rb') as fp:
            msg = email.parser.BytesParser(policy=email.policy.default).parse(fp)

        # Helper to safely decode payload
        def decode_payload(part):
            payload = part.get_payload(decode=True)
            if not payload:
                return None
            charset = part.get_content_charset() or 'utf-8'
            try:
                return payload.decode(charset, errors='replace')
            except LookupError:
                return payload.decode('utf-8', errors='replace')

        # MHTML is typically 'multipart/related'
        if msg.is_multipart():
            # 1. Look for explicit text/html part
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    return decode_payload(part)
            
            # 2. Fallback: Look for text content that resembles HTML
            for part in msg.walk():
                if part.get_content_maintype() == 'text':
                    content = decode_payload(part)
                    if content and '<!DOCTYPE html' in content[:1000]:
                        return content
        else:
            # Handle case where it might not be multipart
            return decode_payload(msg)
        
        print(f"Warning: Could not find main HTML part in MHTML file: {file_path}")
        return None

    except Exception as e:
        print(f"Error processing MHTML file {file_path}: {e}")
        return None

# --- Main Parsing Function ---
def parse_allen_result(file_path, html_content):
    """Parses HTML/MHTML content and returns a dictionary of data."""
    
    filename = os.path.basename(file_path)
    row_data = {'Date': '', 'File Name': filename}

    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 1. Global Info
    title_div = soup.find('div', {'data-testid': 'test-title'})
    if title_div:
        full_text = " ".join(title_div.get_text(separator=" ").split())
        row_data['Test Name'] = full_text.replace("Result:", "").strip()

    # Scores
    score_element = soup.find('div', class_=re.compile(r'text-3xl.*leading-10'))
    if score_element:
        row_data['Glb Score'] = int(score_element.get_text(strip=True))
        # Try to find max marks (usually the next similar element)
        score_matches = re.findall(r'text-3xl.*?leading-10.*?>(\d+)</div>', html_content)
        if len(score_matches) >= 2:
            row_data['Max Marks'] = int(score_matches[1])

    # Percentile
    percentile_match = re.search(r'(\d+\.\d+)\s*Percentile', html_content)
    if percentile_match:
        row_data['Percentile'] = float(percentile_match.group(1))

    # Predictive AIR
    air_element = soup.find(attrs={"data-testid": "rank-range"})
    if air_element:
        row_data['Predictive AIR'] = air_element.get_text(strip=True)

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


# --- Helper Functions ---
def calculate_accuracy(df):
    """Calculates accuracy percentages for Global and Subjects."""
    if 'Glb C' in df.columns and 'Glb W' in df.columns:
        df['Glb Acc%'] = df.apply(lambda x: (x['Glb C'] / (x['Glb C'] + x['Glb W']) * 100) if (x['Glb C'] + x['Glb W']) > 0 else 0, axis=1)

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
    cols = ['Date', 'File Name', 'Test Name', 'Glb Score', 'Percentile', 'Predictive AIR', 'Glb Acc%']
    cols += ['Glb C', 'Glb W', 'Glb U']
    for subj in ['Phy', 'Chem', 'Math']:
        cols += [f'{subj} S', f'{subj} Acc%', f'{subj} C', f'{subj} W', f'{subj} U']
    
    # Keep only columns that exist
    final_cols = [c for c in cols if c in df.columns]
    df = df[final_cols]

    # Initialize Writer
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Apply Dark Background
    (max_row, max_col) = df.shape
    worksheet.set_column(0, max_col-1, 10)
    
    # Widen Title Columns
    worksheet.set_column('B:C', 25) 
    worksheet.set_column('F:F', 18) # Predictive AIR
    
    # Write Headers
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value)

    # Conditional Formatting
    for i, col_name in enumerate(df.columns):
        col_letter = chr(ord('A') + i)
        if i >= 26: col_letter = 'A' + chr(ord('A') + (i-26))
        
        rng = f"{col_letter}2:{col_letter}{max_row+1}"
        
        if any(x in col_name for x in ['Score', 'Percentile', 'Acc%', ' C', ' S']):
            worksheet.conditional_format(rng, {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'})
            if 'Acc%' in col_name or 'Percentile' in col_name:
                 worksheet.set_column(i, i, 12)

        elif ' W' in col_name:
            worksheet.conditional_format(rng, {'type': '3_color_scale', 'min_color': '#63BE7B', 'mid_color': '#FFEB84', 'max_color': '#F8696B'})

    writer.close()
    print(f"Successfully updated and saved to '{output_file}'")

# --- Main Logic ---
def update_excel_sheet():
    # 1. Check for existing data
    if os.path.exists(OUTPUT_FILE):
        try:
            print("Found existing Excel file. Reading data...")
            df_existing = pd.read_excel(OUTPUT_FILE)
            existing_files = df_existing['File Name'].tolist() if 'File Name' in df_existing.columns else []
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

    all_files = [f for f in os.listdir(FOLDER_NAME) if f.lower().endswith(ALLOWED_EXTENSIONS)]
    new_files = [f for f in all_files if f not in existing_files]

    # 3. Parse New Files
    new_results = []
    if new_files:
        print(f"Found {len(new_files)} new files. Parsing...")
        for file in new_files:
            file_path = os.path.join(FOLDER_NAME, file)
            
            html_content = None
            if file.lower().endswith(('.mht', '.mhtml')):
                html_content = read_mhtml(file_path)
            elif file.lower().endswith('.html'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        html_content = f.read()
                except Exception as e:
                    print(f"Error reading HTML file {file}: {e}")

            if html_content:
                try:
                    data = parse_allen_result(file_path, html_content)
                    new_results.append(data)
                    print(f" -> Parsed: {file}")
                except Exception as e:
                    print(f" -> Error parsing content of {file}: {e}")
            else:
                print(f" -> Skipped {file} (Could not extract HTML)")
    else:
        print("No new files found to parse.")

    # 4. Combine Data
    if new_results:
        df_new = pd.DataFrame(new_results)
        df_new = calculate_accuracy(df_new)
        
        if not df_existing.empty:
            df_final = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_final = df_new
    else:
        df_final = df_existing

    if df_final.empty:
        print("No data available to save.")
        return

    # 5. Sort by File Name (Natural Sort)
    # This logic ensures "test1", "test2", "test10" are sorted numerically (1, 2, 10) instead of (1, 10, 2)
    try:
        def natural_keys(text):
            return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(text))]
        
        df_final = df_final.sort_values(
            by='File Name', 
            key=lambda x: x.map(natural_keys),
            ignore_index=True
        )
        print("Sorted data by File Name (natural order).")
    except Exception as e:
        print(f"Warning: Natural sort failed ({e}). Falling back to standard sort.")
        df_final = df_final.sort_values(by='File Name', ignore_index=True)

    # 6. Save with Styling
    apply_styling(df_final, OUTPUT_FILE)

if __name__ == "__main__":
    update_excel_sheet()