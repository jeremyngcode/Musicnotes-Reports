from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
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

# Define styles
no_border = Side(border_style=None)
thin_border = Side(border_style='thin')

no_fill = PatternFill(patternType=None)
grey_fill = PatternFill(
	patternType='solid',
	fgColor='808080'
)
yellow_fill = PatternFill(
	patternType='solid',
	fgColor='FFFF00'
)

# Delete previous data
for row in xl_sheet.iter_rows(min_row=4, min_col=1, max_col=4):
	if row[0].value is None:
		last_row_entry = row[0].row
		break

	for cell in row:
		cell.value = None

for row in xl_sheet[f'A{last_row_entry}:D{last_row_entry+1}']:
	for cell in row:
		cell.value = None
		cell.fill = no_fill
		cell.border = Border(
			top=no_border,
			right=no_border,
			bottom=no_border,
			left=no_border
		)

xl_sheet['B1'] = None

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
else:
	last_row_entry = row

for cell in xl_sheet[last_row_entry+1]:
	cell.fill = grey_fill

# Write formula values for totals
cols = ('B', 'C', 'D')
totals_row = last_row_entry + 2

for cell, col in zip(xl_sheet[totals_row][1:4], cols):
	cell.value = f'=SUM({col}4:{col}{last_row_entry})'

# Write reporting period as header
header = xl_sheet['B1']

header.font = Font(
	bold=True,
	underline='single'
)

if quarter == 'Q1':
	header.value = f'{year} {quarter}'
	header.fill = yellow_fill
elif quarter in ('Q2', 'Q3', 'Q4'):
	header.value = quarter
	header.fill = no_fill
else:
	raise ValueError('Invalid quarter input in settings.')

# Add right outer border
col_D = xl_sheet.column_dimensions['D']
col_D.border = Border(right=thin_border)

for cell in xl_sheet['D']:
	cell.border = Border(right=thin_border)

# Style total revenue cell
total_revenue = xl_sheet[f'D{totals_row}']

total_revenue.border = Border(
	top=thin_border,
	right=thin_border,
	bottom=thin_border,
	left=thin_border
)
total_revenue.fill = yellow_fill

# Update sheet title
xl_sheet.title = reporting_period
print()

xl_wb.save(xl_file)
