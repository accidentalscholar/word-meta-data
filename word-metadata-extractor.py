# ==========================================
# Word Metadata Extractor 
# Version: 4.0
# Citation: Pundir, V. (2026, May 27). Word Meta Data ExtractorVersion (4.0). Retrieved from https://github.com/accidentalscholar/word-meta-data. 
# Citation: RIS and BibTeX files included for referencing software.
# Tested in: Python 3.10.9 64 bit packaged by Anaconda, Inc.
# Reporsitory: https://github.com/accidentalscholar/word-meta-data
# Provided under: GNU AFFERO GENERAL PUBLIC LICENSE (see accompanying license file)
# ==========================================

import sys
import subprocess
import importlib
import os
import glob
import datetime
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog

# --- 1. Fault Tolerance: Auto-Install Missing Libraries ---
REQUIRED_PACKAGES = {
    'pandas': 'pandas',
    'xlsxwriter': 'xlsxwriter'
}

print("Checking dependencies...")
for import_name, pip_name in REQUIRED_PACKAGES.items():
    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Library '{import_name}' not found. Attempting to install '{pip_name}' via pip...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            print(f"Successfully installed {pip_name}.")
            importlib.invalidate_caches()
        except subprocess.CalledProcessError as e:
            print(f"CRITICAL: Failed to install {pip_name}. Please install manually. Error: {e}")
            sys.exit(1)

import pandas as pd

# --- 2. Configuration ---
VERSION = "4.0"
OUTPUT_FILENAME = f"Word_Metadata_Extracted_v{VERSION}.xlsx"

def get_xml_text(tree, tag_name):
    """Finds an XML element by its tag name, completely ignoring namespaces."""
    for elem in tree.iter():
        if elem.tag.endswith(f'}}{tag_name}') or elem.tag == tag_name:
            return elem.text
    return None

def format_editing_time(minutes_str):
    """Converts raw minutes into the requested Word Properties dialog string format."""
    if not minutes_str:
        return "0 days : 0 hours : 0 minutes : 0 seconds", 0.0
    try:
        total_mins = int(float(minutes_str))
        days = total_mins // 1440
        hours = (total_mins % 1440) // 60
        mins = total_mins % 60
        formatted_string = f"{days} days : {hours} hours : {mins} minutes : 0 seconds"
        return formatted_string, float(total_mins)
    except ValueError:
        return "0 days : 0 hours : 0 minutes : 0 seconds", 0.0

def extract_docx_metadata(file_path):
    """Extracts metadata directly from the XML inside a .docx zip archive."""
    metadata = {
        'Title': None, 'Authors': None, 'Last saved by': None, 'Revision Number': 0,
        'Content created (date-time)': None, 'Last date saved (date-time)': None,
        'Total editing time': "0 days : 0 hours : 0 minutes : 0 seconds", 
        '_RawMinutes': 0.0, 
        'Pages': None, 'Word count': None, 
        'Character count': None, 'Paragraph count': None, 'Template': None
    }
    
    try:
        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            # 1. Parse core.xml (Authors, Dates, Revisions)
            if 'docProps/core.xml' in docx_zip.namelist():
                core_xml = docx_zip.read('docProps/core.xml')
                core_tree = ET.fromstring(core_xml)
                
                metadata['Title'] = get_xml_text(core_tree, 'title')
                metadata['Authors'] = get_xml_text(core_tree, 'creator')
                metadata['Last saved by'] = get_xml_text(core_tree, 'lastModifiedBy')
                
                revision = get_xml_text(core_tree, 'revision')
                metadata['Revision Number'] = int(revision) if revision else 0
                
                metadata['Content created (date-time)'] = get_xml_text(core_tree, 'created')
                metadata['Last date saved (date-time)'] = get_xml_text(core_tree, 'modified')

            # 2. Parse app.xml (Pages, Words, Template, Editing Time)
            if 'docProps/app.xml' in docx_zip.namelist():
                app_xml = docx_zip.read('docProps/app.xml')
                app_tree = ET.fromstring(app_xml)
                
                editing_time_raw = get_xml_text(app_tree, 'TotalTime')
                formatted_time, raw_mins = format_editing_time(editing_time_raw)
                
                metadata['Total editing time'] = formatted_time
                metadata['_RawMinutes'] = raw_mins
                
                pages = get_xml_text(app_tree, 'Pages')
                words = get_xml_text(app_tree, 'Words')
                chars = get_xml_text(app_tree, 'Characters')
                paras = get_xml_text(app_tree, 'Paragraphs')
                
                metadata['Pages'] = int(pages) if pages else None
                metadata['Word count'] = int(words) if words else None
                metadata['Character count'] = int(chars) if chars else None
                metadata['Paragraph count'] = int(paras) if paras else None
                metadata['Template'] = get_xml_text(app_tree, 'Template')
                
    except zipfile.BadZipFile:
        print(f"Skipping {os.path.basename(file_path)}: Not a readable .docx archive.")
    except Exception as e:
        print(f"Error reading {os.path.basename(file_path)}: {e}")
        
    return metadata

def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    folder_path = filedialog.askdirectory(title="Select Folder containing Word Documents")
    
    if not folder_path:
        print("No folder selected. Exiting script.")
        return
        
    print(f"\nScanning folder: {folder_path}")
    
    search_pattern = os.path.join(folder_path, "**", "*.docx")
    all_files = glob.glob(search_pattern, recursive=True)
    
    word_files = [f for f in all_files if not os.path.basename(f).startswith('~')]
    
    if not word_files:
        print("No .docx documents found in the selected folder.")
        return
        
    print(f"Found {len(word_files)} .docx document(s). Extracting metadata...")

    metadata_list = []
    
    for file_path in word_files:
        filename = os.path.basename(file_path)
        os_created = datetime.datetime.fromtimestamp(os.path.getctime(file_path))
        extracted_data = extract_docx_metadata(file_path)
        
        row_data = {
            'Filename': filename,
            'File created (date-time)': os_created,
            'Title': extracted_data['Title'],
            'Authors': extracted_data['Authors'],
            'Last saved by': extracted_data['Last saved by'],
            'Revision Number': extracted_data['Revision Number'],
            'Content created (date-time)': extracted_data['Content created (date-time)'],
            'Last date saved (date-time)': extracted_data['Last date saved (date-time)'],
            'Total editing time': extracted_data['Total editing time'],
            'Pages': extracted_data['Pages'],
            'Word count': extracted_data['Word count'],
            'Character count': extracted_data['Character count'],
            'Paragraph count': extracted_data['Paragraph count'],
            'Template': extracted_data['Template'],
            '_RawMinutes': extracted_data['_RawMinutes'] 
        }
        metadata_list.append(row_data)
        
    df = pd.DataFrame(metadata_list)
    
    for col in ['Content created (date-time)', 'Last date saved (date-time)']:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.tz_localize(None)
    
    df['File created (date-time)'] = pd.to_datetime(df['File created (date-time)'], errors='coerce').dt.tz_localize(None)

    output_path = os.path.join(folder_path, OUTPUT_FILENAME)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter', datetime_format='yyyy-mm-dd hh:mm:ss')
    df.to_excel(writer, index=False, sheet_name='Metadata')
    
    workbook = writer.book
    worksheet = writer.sheets['Metadata']
    
    light_orange = workbook.add_format({'bg_color': '#FFD580'})
    light_red    = workbook.add_format({'bg_color': '#FFCCCC'})
    light_green  = workbook.add_format({'bg_color': '#CCFFCC'})
    
    max_row = len(df) + 1
    
    if max_row > 1:
        # Cols D & E (Authors vs Last saved by): Light orange if different
        worksheet.conditional_format(f'D2:E{max_row}', {'type': 'formula', 'criteria': '=$D2<>$E2', 'format': light_orange})
        
        # Col F (Revision Number): Light orange if less than 2
        worksheet.conditional_format(f'F2:F{max_row}', {'type': 'cell', 'criteria': 'less than', 'value': 2, 'format': light_orange})
        
        # Col I (Total editing time): 3-tier formatting using the hidden _RawMinutes column (Column O)
        # We use PERCENTILE (without the .INC) for strict legacy Excel backwards compatibility.
        worksheet.conditional_format(f'I2:I{max_row}', {'type': 'formula', 'criteria': f'=$O2<=PERCENTILE($O$2:$O${max_row}, 0.33)', 'format': light_red})
        worksheet.conditional_format(f'I2:I{max_row}', {'type': 'formula', 'criteria': f'=$O2>=PERCENTILE($O$2:$O${max_row}, 0.67)', 'format': light_green})
        worksheet.conditional_format(f'I2:I{max_row}', {'type': 'formula', 'criteria': f'=AND($O2>PERCENTILE($O$2:$O${max_row}, 0.33), $O2<PERCENTILE($O$2:$O${max_row}, 0.67))', 'format': light_orange})
        
        # Col N (Template): Light orange if not Normal, Normal.dot, or Normal.dotm
        worksheet.conditional_format(f'N2:N{max_row}', {'type': 'formula', 'criteria': '=AND($N2<>"Normal", $N2<>"Normal.dot", $N2<>"Normal.dotm")', 'format': light_orange})

    for i, col in enumerate(df.columns):
        if col == '_RawMinutes':
            continue 
        column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, column_len)
        
    # Hide Column O securely using standard notation
    worksheet.set_column('O:O', None, None, {'hidden': True})

    writer.close()
    print(f"\nSuccess! Metadata extracted and saved to:\n{output_path}")

if __name__ == "__main__":
    main()