import zipfile
import xml.etree.ElementTree as ET

file_path = 'SRS HR COPY.xlsx'

def get_sheet_data(zip_ref, strings, sheet_file, rows_limit=100):
    try:
        with zip_ref.open(sheet_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            rows = root.findall('.//main:row', ns)
            data = []
            for i, row in enumerate(rows[:rows_limit]):
                cells = row.findall('main:c', ns)
                row_cells = []
                for c in cells:
                    v = c.find('main:v', ns)
                    t = c.attrib.get('t')
                    val = v.text if v is not None else ""
                    if t == 's' and val:
                        val = strings[int(val)]
                    row_cells.append((c.attrib.get('r'), val))
                data.append(row_cells)
            return data
    except Exception as e:
        return f"Error: {e}"

try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        # Get shared strings
        strings = []
        try:
            with zip_ref.open('xl/sharedStrings.xml') as f:
                root = ET.parse(f).getroot()
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in root.findall('.//main:si', ns):
                    parts = si.findall('.//main:t', ns)
                    strings.append("".join([p.text for p in parts if p.text]))
        except: pass

        # Get sheet files from workbook.xml.rels or workbook.xml
        # For simplicity, we assume sheet1 is Internal, sheet2 is External, etc.
        # But let's check workbook.xml for the mapping
        with zip_ref.open('xl/workbook.xml') as f:
            root = ET.parse(f).getroot()
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = root.findall('.//main:sheet', ns)
            sheet_map = {s.attrib['name']: f"xl/worksheets/sheet{i+1}.xml" for i, s in enumerate(sheets)}

        print("Internal Sheet Structure (Rows 1-60):")
        internal_data = get_sheet_data(zip_ref, strings, sheet_map['Internal'], 60)
        for row in internal_data:
            print(" | ".join([f"{r}: {v}" for r, v in row]))

        print("\nRetirement Sheet Structure (Rows 1-30):")
        ret_data = get_sheet_data(zip_ref, strings, sheet_map['Retirement'], 30)
        for row in ret_data:
            print(" | ".join([f"{r}: {v}" for r, v in row]))

except Exception as e:
    print(f"Error: {e}")
