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
<<<<<<< HEAD
ALLOWED_EXTENSIONS = ('.html', '.mht', '.mhtml') 
=======
ALLOWED_EXTENSIONS = ('.html', '.mht', '.mhtml') # Added MHTML extensions
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f

# --- MHTML Utility Function ---
def read_mhtml(file_path):
    """
    Reads an MHTML file and extracts the main HTML content.
    MHTML is a multipart/related MIME message.
    """
    try:
<<<<<<< HEAD
        with open(file_path, 'rb') as fp:
            msg = email.parser.BytesParser(policy=email.policy.default).parse(fp)

        # Helper to decode payload with fallback
        def decode_payload(part):
            payload = part.get_payload(decode=True)
            if not payload:
                return None
            
            charset = part.get_content_charset() or 'utf-8'
            try:
                return payload.decode(charset, errors='replace')
            except LookupError:
                # Fallback to utf-8 if charset is invalid
                return payload.decode('utf-8', errors='replace')

        if msg.is_multipart():
            # 1. Try to find part with text/html
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    return decode_payload(part)
            
            # 2. If no explicit text/html, checking for first part that looks like HTML
            # (Sometimes content-type might be missing or generic)
            for part in msg.walk():
                if part.get_content_maintype() == 'text':
                    content = decode_payload(part)
                    if content and '<!DOCTYPE html' in content[:1000]:
                        return content

        else:
            # Not multipart, maybe just a plain HTML file saved with .mhtml extension
            return decode_payload(msg)
        
=======
        # Read the file content as bytes
        with open(file_path, 'rb') as fp:
            msg = email.parser.BytesParser(policy=email.policy.default).parse(fp)

        # MHTML is typically 'multipart/related'
        if msg.is_multipart():
            # Find the main HTML part, usually the first part or the one with content-type text/html
            for part in msg.walk():
                content_type = part.get_content_type()
                
                # We are looking for the main HTML file
                if content_type == 'text/html':
                    # Decode the payload (main HTML content)
                    payload = part.get_payload(decode=True)
                    # Get the charset from the part headers, default to UTF-8
                    charset = part.get_content_charset() or 'utf-8'
                    
                    # Return the decoded HTML as a string
                    return payload.decode(charset)
        
        # If it's not multipart or we couldn't find text/html part
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
        print(f"Warning: Could not find main HTML part in MHTML file: {file_path}")
        return None

    except Exception as e:
        print(f"Error processing MHTML file {file_path}: {e}")
        return None

<<<<<<< HEAD
# --- Main Parsing Function ---
=======
# --- Main Parsing Function (Unchanged, now accepts content string) ---
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
def parse_allen_result(file_path, html_content):
    """Parses HTML/MHTML content and returns a dictionary of data."""
    
    filename = os.path.basename(file_path)
    row_data = {'Date': '', 'File Name': filename}

    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 1. Global Info
    # ... (rest of the parsing logic is unchanged)
    # 1. Global Info
    title_div = soup.find('div', {'data-testid': 'test-title'})
    if title_div:
        full_text = " ".join(title_div.get_text(separator=" ").split())
        row_data['Test Name'] = full_text.replace("Result:", "").strip()

    # Scores (Updated regex to be slightly more flexible if classes change order)
    # Looking for the specific structure of the score circle text
    score_element = soup.find('div', class_=re.compile(r'text-3xl.*leading-10'))
    if score_element:
        row_data['Glb Score'] = int(score_element.get_text(strip=True))
        # Usually the sibling or next element contains max marks, but relying on regex for safety
        score_matches = re.findall(r'text-3xl.*?leading-10.*?>(\d+)</div>', html_content)
        if len(score_matches) >= 2:
            row_data['Max Marks'] = int(score_matches[1])

    # Percentile
    percentile_match = re.search(r'(\d+\.\d+)\s*Percentile', html_content)
    if percentile_match:
        row_data['Percentile'] = float(percentile_match.group(1))

    # Predictive AIR (Added Logic)
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
    # Using regex to find subject sections then parsing the nearby numbers
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
        # Look ahead a bit to find the scores
        search_window = html_content[start_idx:start_idx+1500]
        # Matches <div class="col-span-2 text-center">87</div>
        numbers = re.findall(r'class="col-span-2 text-center">(\d+)</div>', search_window)
        
        if len(numbers) >= 3:
            s_score = int(numbers[0])
            s_correct = int(numbers[1])
            s_incorrect = int(numbers[2])
            s_unattempted = 25 - (s_correct + s_incorrect) # Assuming 25 qs per subject
            
            row_data[f'{short_subj} S'] = s_score
            row_data[f'{short_subj} C'] = s_correct
            row_data[f'{short_subj} W'] = s_incorrect
            row_data[f'{short_subj} U'] = s_unattempted

    return row_data


<<<<<<< HEAD
# --- Helper Functions ---
=======
# --- Unchanged Helper Functions ---
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
def calculate_accuracy(df):
    """Calculates accuracy percentages for Global and Subjects."""
    # Global Accuracy
    if 'Glb C' in df.columns and 'Glb W' in df.columns:
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
    # ... (Styling logic is unchanged)
    
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
    
    # Define Formats
    dark_bg = workbook.add_format({'bg_color': '#000000', 'font_color': '#FFFFFF', 'border': 1, 'border_color': '#444444'})
    header_fmt = workbook.add_format({'bg_color': '#222222', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center'})
    percent_fmt = workbook.add_format({'num_format': '0.0', 'bg_color': '#000000', 'font_color': '#FFFFFF'})
    
    # Apply Dark Background to Data Area
    (max_row, max_col) = df.shape
    worksheet.set_column(0, max_col-1, 10, dark_bg) 
    
    # Widen Title Columns and AIR
    worksheet.set_column('B:C', 25) 
    worksheet.set_column('F:F', 18)
    
    # Write Headers
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # Conditional Formatting
    for i, col_name in enumerate(df.columns):
        col_letter = chr(ord('A') + i)
        if i >= 26: col_letter = 'A' + chr(ord('A') + (i-26))
        
        rng = f"{col_letter}2:{col_letter}{max_row+1}"
        
        if any(x in col_name for x in ['Score', 'Percentile', 'Acc%', ' C', ' S']):
            worksheet.conditional_format(rng, {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'})
            if 'Acc%' in col_name or 'Percentile' in col_name:
                 worksheet.set_column(i, i, 12, percent_fmt)

        elif ' W' in col_name:
            worksheet.conditional_format(rng, {'type': '3_color_scale', 'min_color': '#63BE7B', 'mid_color': '#FFEB84', 'max_color': '#F8696B'})

    writer.close()
    print(f"Successfully updated and saved to '{output_file}'")

<<<<<<< HEAD
# --- Main Logic ---
=======
# --- Main Logic (Updated to handle file extensions) ---
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
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

<<<<<<< HEAD
    # 2. Scan folder
=======
    # 2. Scan folder for NEW files (HTML/MHTML)
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
    if not os.path.exists(FOLDER_NAME):
        print(f"Error: Folder '{FOLDER_NAME}' not found.")
        return

<<<<<<< HEAD
=======
    # Check for all allowed extensions
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
    all_files = [f for f in os.listdir(FOLDER_NAME) if f.lower().endswith(ALLOWED_EXTENSIONS)]
    new_files = [f for f in all_files if f not in existing_files]

    if not new_files:
<<<<<<< HEAD
        print("No new files found. Refreshing styling...")
        if not df_existing.empty:
            apply_styling(df_existing, OUTPUT_FILE)
=======
        print("No new HTML/MHTML files found. Excel is up to date.")
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f
        return

    print(f"Found {len(new_files)} new files. Parsing...")

    # 3. Parse New Files
    new_results = []
    for file in new_files:
        file_path = os.path.join(FOLDER_NAME, file)
        
<<<<<<< HEAD
        html_content = None
        if file.lower().endswith(('.mht', '.mhtml')):
            html_content = read_mhtml(file_path)
        elif file.lower().endswith('.html'):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
            except Exception as e:
                print(f"Error reading HTML file {file}: {e}")
=======
        # Determine the file type and read content
        if file.lower().endswith(('.mht', '.mhtml')):
            html_content = read_mhtml(file_path)
            file_type = "MHTML"
        elif file.lower().endswith('.html'):
            with open(file_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
            file_type = "HTML"
        else:
            # Should not happen if ALLOWED_EXTENSIONS check is correct
            continue

        if html_content:
            try:
                # Pass the content string to the parser
                data = parse_allen_result(file_path, html_content)
                new_results.append(data)
                print(f" -> Parsed ({file_type}): {file}")
            except Exception as e:
                print(f" -> Error parsing {file} content: {e}")
>>>>>>> 2341f6404cd8f88ca57dc2a3104f060e4c80609f

        if html_content:
            try:
                data = parse_allen_result(file_path, html_content)
                new_results.append(data)
                print(f" -> Parsed: {file}")
            except Exception as e:
                print(f" -> Error parsing content of {file}: {e}")
        else:
            print(f" -> Skipped {file} (Could not extract HTML)")

    # 4. Concatenate and Save
    if new_results:
        df_new = pd.DataFrame(new_results)
        df_new = calculate_accuracy(df_new)
        
        if not df_existing.empty:
            df_final = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_final = df_new
            
        apply_styling(df_final, OUTPUT_FILE)
    else:
        # Just re-save to ensure headers/styling are updated if code changed but no new files
        if not df_existing.empty:
            apply_styling(df_existing, OUTPUT_FILE)

if __name__ == "__main__":
    update_excel_sheet()
