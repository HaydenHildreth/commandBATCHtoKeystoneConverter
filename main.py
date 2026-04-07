"""
convert_mix_listing.py



Convert CommandAlkon's commandBATCH mix design listing into a nice
and readable format which can be imported into Keystone.

Output:
  Excel Column A (col 0): Mix Design Name  (blank on the header row, 0.0 sentinel on Name rows)
  Excel Column B (col 1): Ingredient name  (or 'Name:' / 'Mix Design Listing' on header rows)
  Excel Column C (col 2): Mix Design Name  (on Name rows only)
  Excel Column D (col 3): Amount           (numeric string)
  Excel Column E (col 4): Unit             (lb / oz / etc.)

Usage:
  python convert_mix_listing.py [input.xls] [output.xls] [plant_separator]

Defaults (these are used if script is ran without parameters at runtime):
  input  = MixListing.xls
  output = MixListingEdit_converted.xls
  plant_separator = ""
"""

import sys
import xlrd
import xlwt
from datetime import datetime

INPUT  = sys.argv[1] if len(sys.argv) > 1 else "MixListing.xls"
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "MixListingEdit_converted.xls"
PLANT_SEPARATOR= sys.argv[3] if len(sys.argv) > 3 else ""



def parse_mixes(path):
    """
    Read original file into dictionary
      { 'name': str, 'unit': str, 'ingredients': [(name, amount, unit), ...] }
    """
    wb  = xlrd.open_workbook(path, formatting_info=True)
    sh  = wb.sheet_by_index(0)

    mixes       = []
    current_mix = None
    in_ingredients = False

    for r in range(sh.nrows):
        row = sh.row_values(r)

        val0 = str(row[0]).strip()

        # Skip header
        if val0 == "Mix Design Listing":
            continue

        # Skip long dashes
        if val0.startswith("-----") and len(val0) > 20:
            in_ingredients = False
            continue

        # Skip short dashes
        if val0 == "------------":
            continue

        # Name row means new mix begins
        if val0 == "Name:":
            if current_mix:
                mixes.append(current_mix)
            name = str(row[2]).strip() if row[2] != '' else ''
            unit = str(row[6]).strip() if len(row) > 6 and row[6] != '' else 'yd'
            current_mix    = {'name': name, 'unit': unit, 'ingredients': []}
            in_ingredients = False
            continue

        # Skip ingredient row
        if val0 == "Ingredient":
            in_ingredients = True
            continue

        # Skip separating dashes
        if val0 == "------------":
            continue

        # Skip description, mix yield, dates and blanks
        if val0 in ("Description:", "Mix Yield:", "Print Date:", ""):
            continue

        # Read in mix information
        if in_ingredients and current_mix and val0 != '':
            ingredient = val0
            # Amount is in column index 6, unit in 7
            amount = row[6] if len(row) > 6 else ''
            unit   = row[7].upper() if len(row) > 7 else ''
            if amount != '':
                try:
                    amount_str = f"{float(amount):.3f}"
                except (ValueError, TypeError):
                    amount_str = str(amount)
                current_mix['ingredients'].append((ingredient, amount_str, str(unit).strip()))

    # GRAB LAST MIX
    if current_mix:
        mixes.append(current_mix)

    return mixes


def write_output(mixes, path):
    wb  = xlwt.Workbook()
    ws  = wb.add_sheet("Sheet1")

    # Header row (row 0)
    ws.write(0, 1, "Mix Design Listing")

    out_row = 1
    first   = True

    for mix in mixes:
        name = mix['name']
        unit = mix['unit']

        # Name row
        ws.write(out_row, 0, 0.0 if not first else "")
        ws.write(out_row, 1, "Name:")
        ws.write(out_row, 2, name+PLANT_SEPARATOR)
        ws.write(out_row, 3, unit)
        out_row += 1
        first    = False

        # Ingredient rows
        for (ingredient, amount, ing_unit) in mix['ingredients']:
            ws.write(out_row, 0, name+PLANT_SEPARATOR)
            ws.write(out_row, 1, ingredient+PLANT_SEPARATOR)
            ws.write(out_row, 3, amount)
            ws.write(out_row, 4, ing_unit)
            out_row += 1

    wb.save(path)
    print(f"Written {out_row} rows across {len(mixes)} mixes → {path}")


if __name__ == "__main__":
    print(f"Parsing {INPUT} ...")
    mixes = parse_mixes(INPUT)
    print(f"Found {len(mixes)} mix designs.")
    write_output(mixes, OUTPUT)
    # DEBUGGING print(f"{PLANT_SEPARATOR}")
  
