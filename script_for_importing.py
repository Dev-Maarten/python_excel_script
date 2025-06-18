import pandas as pd
import re
import os
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook


# Load the Excel file
excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') or f.endswith('.xls')]
if not excel_files:
    raise FileNotFoundError("No Excel files found in the current directory.")

file_path = 'Meerendael-leden (1).xlsx'
xls = pd.ExcelFile(file_path)

# Load the 'Eigenaren' sheet
df = pd.read_excel(xls, sheet_name='Eigenaren')


def parse(full_name):
    if pd.isna(full_name):
        return '', '', ''
    
    parts = full_name.strip().split()
    if not parts:
        return '', '', ''
    
    voorletters = ""
    tussenvoegsel = ""
    achternaam = ""
    
    
    if "." in parts[0]:
        voorletters = parts[0]
        rest = parts[1:]
    else:
        rest = parts

    tussenvoegsels = {"van", "van de", "van der", "de", "den", "ter", "ten", "het", "der"}
    
    # Check for multi-word tussenvoegsel
    if len(rest) >= 2:
        for i in range(len(rest)-1):
            possible_tv = " ".join(rest[:i+1]).lower()
            if possible_tv in tussenvoegsels:
                tussenvoegsel = possible_tv
                achternaam = " ".join(rest[i+1:])
                break
        else:
            achternaam = " ".join(rest)
    elif rest:
        achternaam = rest[0]
    
    return voorletters.strip(), tussenvoegsel.strip(), achternaam.strip()

def parse_address(address):
    """
    Split an address like 'Street 123, 1234AB City' into components.
    """
    if pd.isna(address):
        return None, None, None, None

    try:
        street_part, rest = address.split(',', 1)
        postcode_match = re.search(r'\d{4}\s?[A-Z]{2}', rest)
        postcode = postcode_match.group(0) if postcode_match else ""
        city = rest.replace(postcode, '').strip() if postcode else rest.strip()

        street_name = re.sub(r'\d.*', '', street_part).strip()
        house_number = street_part.replace(street_name, '').strip()

        return street_name, house_number, postcode.strip(), city
    except Exception:
        return None, None, None, None

# Group rows by Index nr. to collect cohabitants
grouped = df.groupby('Index nr.')

# Define the target structure
output_columns = [
    "(Achter-) naam*", "Voorletters / -naam", "Tussenvoegsel", "Geslacht / type*", 
    "Straatnaam*", "Huisnummer*", "Postcode*", "Plaats*", 
    "Straatnaam postadres", "Huisnummer postadres", "Postcode postadres", "Plaats postadres", 
    "Categorie", "Telefoonnummer 1", "Type telefoonnummer 1", 
    "Telefoonnummer 2", "Type telefoonnummer 2", "Telefoonnummer 3", "Type telefoonnummer 3", 
    "E-mailadressen", "IBAN", "Is debiteur", "Is crediteur", "Incassomachtiging afgegeven", 
    "Factuur gewenst", "Factuurtoelichting gewenst", 
    "Achternaam contactpersoon 1", "Voorletters contactpersoon 1", "Tussenvoegsel contactpersoon 1", 
    "Geslacht contactpersoon 1", "Telefoonnummer 1 contactpersoon 1", "Telefoonnummer 2 contactpersoon 1", 
    "E-mailadres contactpersoon 1", 
    "Achternaam contactpersoon 2", "Voorletters contactpersoon 2", "Tussenvoegsel contactpersoon 2", 
    "Geslacht contactpersoon 2", "Telefoonnummer 1 contactpersoon 2", "Telefoonnummer 2 contactpersoon 2", 
    "E-mailadres contactpersoon 2"
]

output_data = []

# Process each grouped set
for index, group in grouped:
    row = dict.fromkeys(output_columns, "")

    owners = group.reset_index(drop=True)
    primary = owners.iloc[0]
    secondary = owners.iloc[1] if len(owners) > 1 else None
    tertiary = owners.iloc[2] if len(owners) > 2 else None

    # Fill in primary person
    voorletters, tussenvoegsel, achternaam = parse(primary["Eigenaar"])
    row["Voorletters / -naam"] = voorletters
    row["Tussenvoegsel"] = tussenvoegsel
    row["(Achter-) naam*"] = achternaam

    # Parse main address
    straat, huisnr, postcode, plaats = parse_address(primary["Adres"])
    row["Straatnaam*"] = straat
    row["Huisnummer*"] = huisnr
    row["Postcode*"] = postcode
    row["Plaats*"] = plaats

    # Parse post address
    postadres = primary["Postadres eigenenaar"]
    straat_post, huisnr_post, postcode_post, plaats_post = parse_address(postadres)
    row["Straatnaam postadres"] = straat_post
    row["Huisnummer postadres"] = huisnr_post
    row["Postcode postadres"] = postcode_post
    row["Plaats postadres"] = plaats_post

    # Other fields
    row["Categorie"] = primary["Unittype"]
    row["E-mailadressen"] = primary["Email eigenaar"]
    row["Telefoonnummer 1"] = str(primary["Telefoon eigenaar"]).split(',')[0].split(';')[0]

    # Contactpersoon 1
    if secondary is not None:
        voorletters, tussenvoegsel, achternaam = parse(secondary["Eigenaar"])
        row["Achternaam contactpersoon 1"] = achternaam
        row["Tussenvoegsel contactpersoon 1"] = tussenvoegsel
        row["Voorletters contactpersoon 1"] = voorletters
        row["E-mailadres contactpersoon 1"] = secondary["Email eigenaar"]
        row["Telefoonnummer 1 contactpersoon 1"] = str(secondary["Telefoon eigenaar"]).split(',')[0]

    # Contactpersoon 2
    if tertiary is not None:
        voorletters, tussenvoegsel, achternaam = parse(tertiary["Eigenaar"])
        row["Achternaam contactpersoon 1"] = achternaam
        row["Tussenvoegsel contactpersoon 1"] = tussenvoegsel
        row["Voorletters contactpersoon 1"] = voorletters
        row["E-mailadres contactpersoon 2"] = tertiary["Email eigenaar"]
        row["Telefoonnummer 1 contactpersoon 2"] = str(tertiary["Telefoon eigenaar"]).split(',')[0]

    output_data.append(row)

# Save to Excel
output_df = pd.DataFrame(output_data)
output_df.to_excel("converted_ledenlijst.xlsx", index=False) 

from openpyxl import load_workbook

# Load the workbook you just saved
wb = load_workbook("converted_ledenlijst.xlsx")
ws = wb.active

# Auto-adjust column widths
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get column name like 'A', 'B', etc.
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

# Save again with adjusted widths
wb.save("converted_ledenlijst.xlsx")

print("Data has been processed and saved to 'converted_ledenlijst.xlsx'.")