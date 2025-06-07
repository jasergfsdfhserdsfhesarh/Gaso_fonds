import zipfile
import xml.etree.ElementTree as ET
import sys

# Usage: python compare_xls.py old_file.xlsx new_file.xlsx output.txt

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
        # Store the whole row so we can output every column later
        data[key] = row

    # Return mapping and the index of the scan column for further checks
    return data, idx_scan

def main(old_file, new_file, output_file):
    # Build dictionaries for old and new files. Each dictionary maps the
    # "Реквизиты" value to the entire row. The returned index points to the
    # "Отсканировано" column within the row.
    old_data, old_scan_idx = build_dict(old_file)
    new_data, new_scan_idx = build_dict(new_file)

    with open(output_file, 'w', encoding='utf-8') as out:
        for req, new_row in new_data.items():
            # Determine the scanning value in the new file
            new_scan = new_row[new_scan_idx] if len(new_row) > new_scan_idx else ''

            old_row = old_data.get(req)
            old_scan = ''
            if old_row and len(old_row) > old_scan_idx:
                old_scan = old_row[old_scan_idx]

            # If the case was not scanned in the old file but has a value in the
            # new file, output the entire row separated by tabs.
            if (not old_scan or old_scan == '0') and new_scan and new_scan != '0':
                out.write('\t'.join(new_row) + '\n')

if __name__ == '__main__':
    if len(sys.argv) != 4:
        print('Usage: python compare_xls.py old_file.xlsx new_file.xlsx output.txt')
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])