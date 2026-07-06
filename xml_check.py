import zipfile
import xml.etree.ElementTree as ET
import os
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

def parse_xlsx_fast(path):
    t0 = time.time()
    with zipfile.ZipFile(path, 'r') as z:
        # Step 1: Read workbook.xml to map sheet names to relation IDs
        wb_data = z.read('xl/workbook.xml')
        root_wb = ET.fromstring(wb_data)
        
        # We need namespace map
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
              'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        
        target_sheet_rid = None
        for sheet in root_wb.findall('.//ns:sheet', ns):
            name = sheet.attrib.get('name')
            if name == 'Chi tiết data':
                target_sheet_rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                break
        
        if not target_sheet_rid:
            return "Sheet 'Chi tiết data' not found"
            
        # Step 2: Read workbook.xml.rels to map relation ID to sheet file path
        rels_data = z.read('xl/_rels/workbook.xml.rels')
        root_rels = ET.fromstring(rels_data)
        r_ns = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        
        sheet_path_in_zip = None
        for rel in root_rels.findall('.//rels:Relationship', r_ns):
            rid = rel.attrib.get('Id')
            if rid == target_sheet_rid:
                # Target path is usually relative to xl/
                target = rel.attrib.get('Target')
                sheet_path_in_zip = f"xl/{target}"
                break
                
        if not sheet_path_in_zip:
            return "Sheet relation path not found"
            
        print(f"Reading worksheet from {sheet_path_in_zip}")
        
        # Step 3: Stream parse the worksheet XML using iterparse
        # We are looking for cells in column AF.
        # Column AF cells look like <c r="AF3" ...>
        # We only care about rows >= 3 (since row 1 is metadata, row 2 is header)
        # Wait, the cell reference is r="AF3", "AF4", etc.
        # We want to check if they have a child <v> tag (or any value) that is not empty.
        
        total_rows_with_af = 0
        
        sheet_file = z.open(sheet_path_in_zip)
        
        # We use iterparse to avoid loading the whole XML into memory
        context = ET.iterparse(sheet_file, events=('start', 'end'))
        
        # Fast helper: check if string starts with "AF" and is followed by digits representing row >= 3
        # e.g. "AF3", "AF10", etc.
        for event, elem in context:
            if event == 'end' and elem.tag.endswith('}c'):
                r = elem.attrib.get('r')
                if r and r.startswith('AF'):
                    row_num_str = r[2:]
                    if row_num_str.isdigit():
                        row_num = int(row_num_str)
                        if row_num >= 3: # 1-based index (row 1 is 1st row, row 2 is 2nd row)
                            # Check if cell has value/content
                            # In openxml, if the cell is present but empty, it might not have a <v> tag
                            # or the <v> tag might be empty.
                            v_elem = elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                            if v_elem is not None and v_elem.text is not None and v_elem.text.strip() != "":
                                # Check if it is a formula that evaluated to blank (in some cases <v> is empty string)
                                total_rows_with_af += 1
                elem.clear() # Free memory
                
        t1 = time.time()
        print(f"Fast parsing took: {t1-t0:.2f} seconds")
        print(f"Total rows with AF data (rows >= 3): {total_rows_with_af}")
        return total_rows_with_af

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
parse_xlsx_fast(os.path.join(folder, files[0]))
