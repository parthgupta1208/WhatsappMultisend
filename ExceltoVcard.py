import openpyxl
import vobject

# Load the Excel file and select the sheet with the contacts
workbook = openpyxl.load_workbook('contact.xlsx')
sheet = workbook['Sheet1']

# Create a list to store the vCards
vcards = []

# Iterate over the rows in the sheet and create a vCard for each contact
for row in sheet.iter_rows(min_row=2, values_only=True):
    # Extract the contact information
    first_name, phone = row

    # Create a vCard for the contact
    vcard = vobject.vCard()
    vcard.add('fn').value = first_name
    vcard.add('tel').value = str(phone)

    # Add the vCard to the list
    vcards.append(vcard)

# Write the vCards to a file
with open('contacts.vcf', 'w') as f:
    for vcard in vcards:
        f.write(vcard.serialize())

print('vCards written to contacts.vcf')