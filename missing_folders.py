import os
import sys
from compare_xls import load_rows


def main(xls_path, output_path="result.txt"):
    rows = load_rows(xls_path)
    missing = []
    for row in rows[1:]:
        if len(row) < 8:
            continue
        h_val = row[8]
        if h_val and h_val != '0':
            req = f"{row[0]}-{row[1]}-{row[2]}"
            if not os.path.isdir(req):
                missing.append(req)
    with open(output_path, "w", encoding="utf-8") as f:
        for r in missing:
            f.write(r + "\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python missing_folders.py input.xlsx [output.txt]")
        sys.exit(1)
    xls = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "result.txt"
    main(xls, out)
