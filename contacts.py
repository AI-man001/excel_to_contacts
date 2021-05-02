import glob, os, re
from pandas import *

contact = """\
BEGIN:VCARD
VERSION:2.1
FN:{0}
TEL;CELL:{1}
END:VCARD
"""


def remove_prefix(text, prefix):
  return text[text.startswith(prefix) and len(prefix):]


_filternames = ["12 FEB"]


def saveContacts():
  count = 0
  _path = os.path.join('excels', '*.xlsx')
  worksheets = [_file for _file in glob.glob(_path)]
  #try:
  for worksheet in worksheets:
    xls = ExcelFile(worksheet)
    for sheetname in xls.sheet_names:
      if sheetname in _filternames:
        data = xls.parse(sheetname)
        values = data.to_dict()
        names = values.get('NAME')
        surnames = values.get('SURNAME')
        countries = values.get('COUNTRY')
        phones = values.get('PHONE NUMBER')
        if names:
          contacts = {
              k: {
                  'name': f"{names[k]} {surnames[k]} {countries[k]}",
                  'number': f"+{re.sub(' ', '', str(phones[k]))}"
              }
              for k in phones
          }
          for _id in contacts:
            count += 1
            filename = "zoomcontacts.vcf"
            with open(filename, "a") as text_file:
              name = contacts[_id].get('name').title()
              text_file.write(contact.format(name, contacts[_id].get('number')))
    print(count)
  # except Exception as ex:
  #   print(ex)


saveContacts()