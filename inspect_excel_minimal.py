import zipfile
import xml.etree.ElementTree as ET

file_path = 'SRS HR COPY.xlsx'

try:
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        # Load workbook to get sheet names
        with zip_ref.open('xl/workbook.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            # Namespaces
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = root.findall('.//main:sheet', ns)
            print("Sheet names:")
            for sheet in sheets:
                print(f" - {sheet.attrib['name']} (rId: {sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')})")
        
        # Try to peak into the first sheet to see headers
        try:
            with zip_ref.open('xl/worksheets/sheet1.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                rows = root.findall('.//main:row', ns)
                print("\nPeaking into Sheet 1:")
                for i, row in enumerate(rows[:20]):
                    cells = row.findall('main:c', ns)
                    row_data = []
                    for c in cells:
                        v = c.find('main:v', ns)
                        t = c.attrib.get('t')
                        val = v.text if v is not None else ""
                        row_data.append(f"{c.attrib.get('r')}: {val} (type: {t})")
                    print(" | ".join(row_data))
        except Exception as e:
            print(f"Could not read sheet1: {e}")

except Exception as e:
    print(f"Error: {e}")
