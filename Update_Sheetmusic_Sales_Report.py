from openpyxl import load_workbook
from settings import *
# -------------------------------------------------------------------------------------------------

# Open Master Excel workbook
xl_wb = load_workbook(master_xl_file, read_only=True)
xl_sheet = xl_wb.worksheets[0]

# Retrieve sheetmusic titles
master_list = []

for row in xl_sheet.iter_rows(min_row=4, min_col=1, max_col=1):
	if row[0].value is not None:
		master_list.append(row[0].value)
	else:
		break

# Open Musicnotes Excel workbook
xl_wb = load_workbook(musicnotes_xl_file, read_only=True)
xl_sheet = xl_wb.worksheets[0]

# Retrieve sheetmusic revenue data
print('Retrieving data..')

musicnotes_data = {}

for row in xl_sheet.iter_rows(min_row=5, min_col=2, max_col=6, values_only=True):
	if row[0] is not None:
		musicnotes_data[row[0].lower()] = {
			'Downloads': row[2],
			'Sales': row[3],
			'Revenue': row[4]
		}
	else:
		break

custom_printer.pprint(musicnotes_data)
print()



# Open Latest Revenue Excel workbook
xl_wb = load_workbook(xl_file)
xl_sheet = xl_wb.worksheets[0]

# Delete previous data
for row in xl_sheet.iter_rows(min_row=4, min_col=1, max_col=4):
	for cell in row:
		cell.value = None

# Write data
print('Writing data..')

for row, title in enumerate(master_list, 4):
	xl_sheet[f'A{row}'] = title
	title = title.lower()

	if title in musicnotes_data:
		title_data = musicnotes_data[title]
		xl_sheet[f'B{row}'] = title_data['Downloads']
		xl_sheet[f'C{row}'] = round(title_data['Sales'], 2)
		xl_sheet[f'D{row}'] = round(title_data['Revenue'], 2)

# Update sheet title
xl_sheet.title = reporting_period
print()

xl_wb.save(xl_file)
