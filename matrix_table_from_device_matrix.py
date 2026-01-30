import streamlit as st
import pandas as pd

import streamlit as st
import pandas as pd

# Load the device matrix
matrix_df = pd.read_excel('PLC_Device_Matrix.xlsx', engine='openpyxl')

# Get unique PLC and Manual configurations from the matrix
plc_columns = matrix_df['PLC configuration'].dropna().unique()
manual_rows = matrix_df['Manual configuration'].dropna().unique()

# Create an empty DataFrame for the table
exported_df = pd.read_excel('exported_configurations.xlsx', engine='openpyxl')

result_df = pd.DataFrame(index=manual_rows, columns=plc_columns)

# Fill the table with priorities from exported_df

# Sprint assignment rules - dynamically loaded from Excel
# Build sprint_map from the Excel file
sprint_map = {}

# Normalize sprint values for matching
def normalize_sprint(s):
	return s.strip().lower()

for _, row in matrix_df.iterrows():
	plc_config = row['PLC configuration']
	sprint_rules = row['sprint aasignment rules']
	if pd.notna(plc_config) and pd.notna(sprint_rules):
		# Parse the sprint rules string (format: "sprint 130 , sprint 131")
		sprints = [normalize_sprint(s) for s in str(sprint_rules).split(',')]
		sprint_map[plc_config] = sprints

print("Loaded sprint assignment rules from Excel:")
for plc, sprints in sprint_map.items():
    print(f"{plc}: {sprints}")

# Load sprint color codes from Excel
# Map color names to hex codes
color_name_to_hex = {
    # Light colors
    'light blue': '#cce6ff',
    'light yellow': '#ffffcc',
    'light orange': '#ffd9b3',
    'light red': '#ffcccc',
    'light green': '#d9f2d9',
    'light purple': '#f2d9ff',
    'light pink': '#ffccf2',
    'light brown': '#e6d3b3',
    'light violet': '#e6ccff',
    'light indigo': '#ccddff',
    
    # Regular colors
    'blue': '#99ccff',
    'yellow': '#ffff99',
    'orange': '#ffcc99',
    'red': '#ff9999',
    'green': '#99ff99',
    'purple': '#cc99ff',
    'pink': '#ff99cc',
    'brown': '#cc9966',
    'violet': '#cc99ff',
    'indigo': '#9999ff',
    
    # Dark colors
    'dark blue': '#3366cc',
    'dark yellow': '#cccc00',
    'dark orange': '#cc6600',
    'dark red': '#cc3333',
    'dark green': '#339933',
    'dark purple': '#6633cc',
    'dark pink': '#cc3399',
    'dark brown': '#996633',
    'dark violet': '#663399',
    'dark indigo': '#333399'
}

sprint_color_map = {}

for _, row in matrix_df.iterrows():
	color_code = row['Sprint display color code']
	if pd.notna(color_code):
		# Parse format: "sprint 130- light blue"
		parts = str(color_code).split('-', 1)
		if len(parts) == 2:
			sprint_num = normalize_sprint(parts[0])  # normalized "sprint 130"
			color_name = parts[1].strip().lower()  # "light blue"
			hex_color = color_name_to_hex.get(color_name, '#ff0000')  # Default red if color name not recognized
			sprint_color_map[sprint_num] = hex_color

print("\nLoaded sprint color mappings from Excel:")
for sprint, color in sprint_color_map.items():
    print(f"{sprint}: {color}")

# Helper to get sprint group for a PLC config
def get_sprint_group(plc):
	# First try exact match
	if plc in sprint_map:
		return sprint_map[plc]
	# If no exact match, try partial match (for backward compatibility)
	for key in sprint_map:
		if key in plc:
			return sprint_map[key]
	return []

# Track sprint assignment index for each PLC config
sprint_indices = {plc: 0 for plc in plc_columns}


# Improved sprint assignment for equal distribution
from collections import Counter

# First, count how many cells need sprint assignment per PLC column
plc_priority_cells = {plc: [] for plc in plc_columns}
for _, row in exported_df.iterrows():
	plc = row['PLC configuration']
	manual = row['Manual configuration']
	priority = row['Priority']
	if plc in plc_columns and manual in manual_rows:
		plc_priority_cells[plc].append(manual)





# Helper to distribute sprints globally and balanced
def distribute_sprints(plc_patterns, sprint_values):
	# plc_patterns now contains the exact PLC names in the group
	plc_group = [plc for plc in plc_columns if plc in plc_patterns]
	cells_group = []
	for plc in plc_group:
		cells_group.extend([(manual, plc) for manual in plc_priority_cells[plc]])
	n = len(cells_group)
	sprint_count = len(sprint_values)
	base = n // sprint_count
	remainder = n % sprint_count
	sprint_list = []
	for i, sprint in enumerate(sprint_values):
		sprint_list.extend([sprint] * (base + (1 if i < remainder else 0)))
	for i, (manual, plc) in enumerate(cells_group):
		result_df.at[manual, plc] = sprint_list[i]
	return set(plc_group)

# Apply global distribution for each unique sprint group
# Group PLCs by their sprint assignment rules
from collections import defaultdict
sprint_groups = defaultdict(list)
for plc in plc_columns:
	sprints = get_sprint_group(plc)
	if sprints:
		# Use tuple of sprints as key for grouping
		sprints_key = tuple(sprints)
		sprint_groups[sprints_key].append(plc)

done_plc = set()
for sprints_key, plc_list in sprint_groups.items():
	# Get the PLC patterns (for matching)
	plc_patterns = plc_list
	sprint_values = list(sprints_key)
	done_plc |= distribute_sprints(plc_patterns, sprint_values)

# For other columns, distribute sprints in round-robin
for plc in plc_columns:
	if plc in done_plc:
		continue
	sprints = get_sprint_group(plc)
	cells = plc_priority_cells[plc]
	if sprints:
		for i, manual in enumerate(cells):
			result_df.at[manual, plc] = sprints[i % len(sprints)]
	else:
		for manual in cells:
			result_df.at[manual, plc] = ''

st.title('Device Matrix Table')
st.write('Table with PLC configuration as columns and Manual configuration as rows:')

# Define a function to color cells based on sprint - colors from Excel
def color_sprint(val):
	if pd.isna(val) or val == '':
		return ''  # No color for empty cells
	# Normalize sprint value for lookup
	norm_val = normalize_sprint(val)
	color = sprint_color_map.get(norm_val, '#ff0000')  # Red for missing colors
	return f'background-color: {color}'

# Replace all NaN or empty cells with empty string for display
result_df = result_df.fillna('')

styled_df = result_df.style.map(color_sprint)


st.write(styled_df.to_html(), unsafe_allow_html=True)

# Add General checks information below the matrix
st.header('General Checks')

# Load general checks from Excel
general_checks = matrix_df['General checks'].dropna().tolist()
if general_checks:
	for i, check in enumerate(general_checks, 1):
		st.write(f"**{i}.** {check}")
else:
	st.write("No general checks information available.")# Export test matrix button
if st.button('Export test matrix'):
	export_path = 'Test Matrix.xlsx'
	# Use XlsxWriter for better formatting
	with pd.ExcelWriter(export_path, engine='xlsxwriter') as writer:
		# Write data starting from row 2 and column 2 to leave space for master headings
		result_df.to_excel(writer, sheet_name='Matrix', index=True, startrow=2, startcol=1)
		workbook  = writer.book
		worksheet = writer.sheets['Matrix']

		# Set column widths for better readability
		worksheet.set_column(0, 0, 30)  # Device configurations column
		worksheet.set_column(1, 1, 30)  # Manual configuration column
		worksheet.set_column(2, len(result_df.columns) + 1, 18)

		# Add master heading formats
		master_header_format = workbook.add_format({
			'bold': True,
			'text_wrap': True,
			'valign': 'center',
			'align': 'center',
			'fg_color': '#92CDDC',
			'border': 1,
			'font_size': 12
		})
		
		# Add a header format for column headings
		header_format = workbook.add_format({
			'bold': True,
			'text_wrap': True,
			'valign': 'center',
			'align': 'center',
			'fg_color': '#D7E4BC',
			'border': 1
		})
		
		# Write master row heading "Device configurations" - merge vertically across all device rows
		worksheet.merge_range(3, 0, len(result_df.index) + 2, 0, 'Device configurations', master_header_format)
		
		# Write master column heading "PLC configurations" at (0, 2) and merge across all PLC columns
		worksheet.merge_range(0, 2, 0, len(result_df.columns) + 1, 'PLC configurations', master_header_format)
		
		# Rewrite the column headers at row 2 (since we started data at row 2)
		for col_num, value in enumerate(result_df.reset_index().columns.values):
			worksheet.write(2, col_num + 1, value, header_format)

		# Add cell coloring for sprints (from Excel color codes)
		# Adjust for new starting position (row 3, col 2 in Excel)
		for row in range(len(result_df.index)):
			for col in range(len(result_df.columns)):
				value = result_df.iloc[row, col]
				if value and value != '':
					# Use color from Excel, default to red if not found
					norm_val = normalize_sprint(value)
					color = sprint_color_map.get(norm_val, '#ff0000')
					cell_format = workbook.add_format({'bg_color': color, 'border': 1})
					worksheet.write(row + 3, col + 2, value, cell_format)
				else:
					cell_format = workbook.add_format({'border': 1})
					worksheet.write(row + 3, col + 2, value, cell_format)

		# Add legend below the table - dynamically generated from Excel colors
		legend_start_row = len(result_df.index) + 5
		
		# Legend title
		legend_title_format = workbook.add_format({'bold': True, 'font_size': 12})
		worksheet.write(legend_start_row, 0, 'Legend:', legend_title_format)
		
		# Group sprints by color for legend
		color_to_sprints = {}
		for sprint, color in sprint_color_map.items():
			if color not in color_to_sprints:
				color_to_sprints[color] = []
			color_to_sprints[color].append(sprint)
		
		legend_entries = []
		for color, sprints in sorted(color_to_sprints.items()):
			sprints_str = ', '.join(sorted(sprints, key=lambda x: int(x.split()[1])))
			legend_entries.append((sprints_str, color))
		
		for idx, (text, color) in enumerate(legend_entries):
			legend_format = workbook.add_format({'bg_color': color, 'border': 1})
			worksheet.write(legend_start_row + idx + 1, 0, text, legend_format)
		
		# Adjust column width for legend
		worksheet.set_column(0, 0, max(30, max((len(text) for text, _ in legend_entries), default=30) + 5))

		# Add General Checks section below the legend
		general_checks = matrix_df['General checks'].dropna().tolist()
		if general_checks:
			general_start_row = legend_start_row + len(legend_entries) + 3
			
			# General Checks title
			general_title_format = workbook.add_format({'bold': True, 'font_size': 14})
			worksheet.write(general_start_row, 0, 'General Checks:', general_title_format)
			
			# General checks entries
			general_format = workbook.add_format({'border': 1, 'text_wrap': True})
			for idx, check in enumerate(general_checks):
				worksheet.write(general_start_row + idx + 1, 0, f"{idx + 1}. {check}", general_format)
			
			# Adjust column width for general checks
			max_check_length = max(len(f"{i+1}. {check}") for i, check in enumerate(general_checks))
			current_width = max(30, max_check_length + 5)
			worksheet.set_column(0, 0, current_width)

	st.success(f'Test matrix exported to {export_path}')

# Add legend below the table - dynamically generated from Excel colors
legend_html = "<div style='margin-top:20px;'><b>Legend:</b><br>"
# Group sprints by color for legend
color_to_sprints = {}
for sprint, color in sprint_color_map.items():
	if color not in color_to_sprints:
		color_to_sprints[color] = []
	color_to_sprints[color].append(sprint)

for color, sprints in sorted(color_to_sprints.items()):
	sprints_str = ', '.join([s.capitalize() for s in sorted(sprints, key=lambda x: int(x.split()[1]))])
	legend_html += f"<span style='background-color:{color};padding:4px 12px;border-radius:4px;'>{sprints_str}</span><br>"
legend_html += "</div>"
st.markdown(legend_html, unsafe_allow_html=True)
