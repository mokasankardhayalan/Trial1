import streamlit as st
import pandas as pd

# Load the Excel file
df = pd.read_excel('PLC_Device_Matrix.xlsx', engine='openpyxl')

# Extract columns
plc_config_col = 'PLC configuration'
manual_config_col = 'Manual configuration'

plc_config_values = df[plc_config_col].dropna().unique()
manual_config_values = df[manual_config_col].dropna().unique()

st.title('PLC Device Matrix Viewer')

# Create two columns for side-by-side layout
col1, col2 = st.columns(2)

# PLC configuration in first column
with col1:
    st.header('PLC Configuration')
    plc_selected = st.radio('Select PLC configuration:', plc_config_values)
    st.write(f'Selected PLC configuration: {plc_selected}')

# Manual configuration in second column
with col2:
    st.header('Manual Configuration')
    manual_selected = st.radio('Select Manual configuration:', manual_config_values)
    st.write(f'Selected Manual configuration: {manual_selected}')

# Priority template
st.header('Priority')
priority_levels = ['very low', 'low', 'medium', 'high', 'very high']
priority_selected = st.radio('Select Priority:', priority_levels)
st.write(f'Selected Priority: {priority_selected}')

# Add configuration button and table display
if 'configurations' not in st.session_state:
	st.session_state['configurations'] = []

if st.button('Add configuration'):
	st.session_state['configurations'].append({
		'PLC configuration': plc_selected,
		'Manual configuration': manual_selected,
		'Priority': priority_selected
	})

if st.session_state['configurations']:
	st.header('Added Configurations')
	configs_df = pd.DataFrame(st.session_state['configurations'])
	configs_df = pd.DataFrame(st.session_state['configurations'])
	st.table(configs_df)
	selected_idx = st.radio('Select configuration to remove:', options=range(len(configs_df)), format_func=lambda i: f"Row {i+1}")
	if st.button('Remove configuration'):
		if len(st.session_state['configurations']) > 0:
			st.session_state['configurations'].pop(selected_idx)

	# Export configurations to Excel
	if st.button('Export the configuration'):
		export_path = 'exported_configurations.xlsx'
		configs_df.to_excel(export_path, index=False)
		st.success(f'Configurations exported to {export_path}')
