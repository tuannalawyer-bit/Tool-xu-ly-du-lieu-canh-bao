import zipfile
import xml.etree.ElementTree as ET
import os
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Bùi Anh Tuấn.xlsx"

t0 = time.time()
with zipfile.ZipFile(path, 'r') as z:
    # Read shared strings
    print("Reading shared strings...")
    sst_file = z.open('xl/sharedStrings.xml')
    # Using iterparse for sharedStrings to save memory
    shared_strings = []
    context = ET.iterparse(sst_file, events=('start', 'end'))
    
    # In sharedStrings.xml, each string is inside a <si> element, usually in a <t> child
    current_si = []
    in_t = False
    for event, elem in context:
        if event == 'start' and elem.tag.endswith('}t'):
            in_t = True
        elif event == 'end' and elem.tag.endswith('}t'):
            in_t = False
        elif event == 'end' and elem.tag.endswith('}si'):
            # Join all text inside this si element
            # There can be multiple <t> elements inside <si> if there is rich text formatting
            # But usually it's just one <t>. Let's do a quick gather.
            text = "".join([t.text for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text is not None])
            shared_strings.append(text)
            elem.clear()
            
    print(f"Loaded {len(shared_strings)} shared strings. Time: {time.time()-t0:.2f}s")
    
    # Let's print some of them
    for idx in [0, 7524, 10073, 10077]:
        if idx < len(shared_strings):
            print(f"Index {idx}: '{shared_strings[idx]}'")
            
    # Now let's parse sheet3.xml and check values
    print("Parsing sheet3.xml...")
    sheet_file = z.open('xl/worksheets/sheet3.xml')
    context = ET.iterparse(sheet_file, events=('start', 'end'))
    
    total_data_rows = 0
    non_empty_af = 0
    
    for event, elem in context:
        if event == 'end' and elem.tag.endswith('}c'):
            r = elem.attrib.get('r')
            if r and r.startswith('AF'):
                row_num_str = r[2:]
                if row_num_str.isdigit():
                    row_num = int(row_num_str)
                    if row_num >= 3: # 1-based index, row 1 is metadata, row 2 is header
                        total_data_rows += 1
                        
                        # Get the cell value
                        t_type = elem.attrib.get('t')
                        v_elem = elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        val = None
                        if v_elem is not None and v_elem.text is not None:
                            val_text = v_elem.text
                            if t_type == 's': # shared string
                                idx = int(val_text)
                                if idx < len(shared_strings):
                                    val = shared_strings[idx]
                            else:
                                val = val_text
                                
                        if val is not None and str(val).strip() != "":
                            non_empty_af += 1
            elem.clear()
            
    print(f"Done. Time: {time.time()-t0:.2f}s")
    print(f"Total rows checked: {total_data_rows}")
    print(f"Non-empty AF: {non_empty_af}")
