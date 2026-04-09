import zipfile
import xml.etree.ElementTree as ET

file_path = 'SRS HR COPY.xlsx'

try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        # 1. Get shared strings
        strings = []
        try:
            with zip_ref.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in root.findall('.//main:si', ns):
                    t = si.find('main:t', ns)
                    if t is not None:
                        strings.append(t.text)
                    else:
                        # Handle rich text strings?
                        parts = si.findall('.//main:t', ns)
                        strings.append("".join([p.text for p in parts if p.text]))
        except:
            strings = []

        # 2. Get sheet names
        with zip_ref.open('xl/workbook.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = root.findall('.//main:sheet', ns)
            print("Sheet names:")
            for sheet in sheets:
                print(f" - {sheet.attrib['name']}")
        
        # 3. Peak into the first sheet
        try:
            with zip_ref.open('xl/worksheets/sheet1.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                rows = root.findall('.//main:row', ns)
                print("\nPeaking into Sheet 1 (first 10 rows):")
                for i, row in enumerate(rows[:10]):
                    cells = row.findall('main:c', ns)
                    row_data = []
                    for c in cells:
                        v = c.find('main:v', ns)
                        t = c.attrib.get('t')
                        val = v.text if v is not None else ""
                        if t == 's' and val:
                            val = strings[int(val)]
                        row_data.append(f"{c.attrib.get('r')}: {val}")
                    print(" | ".join(row_data))
        except Exception as e:
            print(f"Could not read sheet1: {e}")

except Exception as e:
    print(f"Error: {e}")
