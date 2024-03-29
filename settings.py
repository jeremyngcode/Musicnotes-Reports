from os import path
import pathlib
import pprint
# -------------------------------------------------------------------------------------------------

# EDIT ACCORDINGLY
year = 'YYYY'
quarter = 'QQ'
reporting_period = f'{year}-{quarter}'

master_xl_file = pathlib.Path(
	"C:/path/to/my/master_excel_file.xlsx"
)
musicnotes_xl_file = pathlib.Path(
	"C:/path/to/my/latest_musicnotes_revenue_report.xlsx"
)
template_xl_file = pathlib.Path(
	path.join(path.dirname(__file__), "xl-template.xlsx")
)

# Script outputs to this file path
output_file = pathlib.Path(
	"C:/path/to/my/script_output.xlsx"
)

# Settings for pprint PrettyPrinter
custom_printer = pprint.PrettyPrinter(
	depth=None,
	indent=1,
	width=100,
	sort_dicts=False,
	compact=False
)
