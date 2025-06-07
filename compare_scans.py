import zipfile
import xml.etree.ElementTree as ET
import sys

# Usage: python compare_scans.py old_file.xlsx new_file.xlsx output.txt

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def load_rows(path):
    """Return list of rows (as lists of cell values) from sheet1."""
    with zipfile.ZipFile(path) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = z.read('xl/sharedStrings.xml')
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root:
                text = ''.join(t.text or '' for t in si.iter('{%s}t' % NS['m']))
                strings.append(text)
        sheet_xml = z.read('xl/worksheets/sheet1.xml')
        root = ET.fromstring(sheet_xml)
        sheet_data = root.find('m:sheetData', NS)
        rows = []
        for row in sheet_data:
            row_values = []
            for c in row.findall('m:c', NS):
                v = c.find('m:v', NS)
                if v is None:
                    value = ''
                else:
                    if c.get('t') == 's':
                        idx = int(v.text)
                        value = strings[idx]
                    else:
                        value = v.text
                row_values.append(value)
            rows.append(row_values)
        return rows

def build_dict(path):
    rows = load_rows(path)
    header = rows[0]
    try:
        idx_scan = header.index('Отсканировано')
        idx_req = header.index('Реквизиты')
    except ValueError:
        raise ValueError('Required columns not found')
    data = {}
    for row in rows[1:]:
        if len(row) <= max(idx_scan, idx_req):
            continue
        key = row[idx_req]
        val = row[idx_scan]
        data[key] = val
    return data

def main(old_file, new_file, output_file):
    old = build_dict(old_file)
    new = build_dict(new_file)
    with open(output_file, 'w', encoding='utf-8') as out:
        for req, scan_val in new.items():
            old_val = old.get(req, '')
            if (not old_val or old_val == '0') and scan_val and scan_val != '0':
                out.write(req + '\n')

if __name__ == '__main__':
    if len(sys.argv) != 4:
        print('Usage: python compare_scans.py old_file.xlsx new_file.xlsx output.txt')
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])
