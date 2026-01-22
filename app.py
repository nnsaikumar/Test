import pandas as pd
import streamlit as st
from io import BytesIO, StringIO
import openpyxl
from openpyxl.styles import Alignment
import numpy as np

def parse_pasted_data(pasted_text):
    """Parse pasted Excel data (tab-separated) into DataFrame"""
    try:
        # Try to parse as tab-separated (Excel default when copying)
        df = pd.read_csv(StringIO(pasted_text), sep='\t', engine='python')
        return df
    except Exception as e:
        st.error(f"Error parsing pasted data: {e}")
        return None

def parse_uploaded_file(uploaded_file, sheet_name='Form Definitions', is_ptd=False):
    """Parse uploaded Excel file into DataFrame"""
    try:
        if is_ptd:
            # For PTD files: Skip first row (index 0) and use second row as headers
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl', header=1)
        else:
            # For SDS files: Use first row as headers (default)
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
        return df
    except ValueError as e:
        st.error(f"Error: Sheet '{sheet_name}' not found in the uploaded file. Available sheets: {pd.ExcelFile(uploaded_file).sheet_names}")
        return None
    except Exception as e:
        st.error(f"Error reading uploaded file: {e}")
        return None

def process_ptd_dataframe(df):
    """Process PTD dataframe: remove specific columns and filter by 'Used in trial'"""
    if df is None:
        return None
    
    # Columns to remove from PTD
    columns_to_remove = [
        'Modification comments + Highlight Cells where change made',
        'Library source'
    ]
    
    # Remove columns if they exist
    for col in columns_to_remove:
        if col in df.columns:
            df = df.drop(columns=[col])
    
    # Filter by "Used in trial (Y, N, Mod)" column
    trial_column_names = [
        'Used in trial (Y, N, Mod)',
        'Used in trial (Y, N, Mod) ',  # with trailing space
        'Used in trial',
    ]
    
    trial_column = None
    for col_name in trial_column_names:
        if col_name in df.columns:
            trial_column = col_name
            break
    
    if trial_column:
        # Filter to only keep rows where the value is 'Y'
        original_count = len(df)
        df = df[df[trial_column].astype(str).str.strip().str.upper() == 'Y'].copy()
        filtered_count = len(df)
        return df, original_count, filtered_count
    else:
        # If column not found, return dataframe as is
        st.warning("⚠️ Column 'Used in trial (Y, N, Mod)' not found in PTD data. Skipping filter.")
        return df, len(df), len(df)

def find_matching_rows(source_df, target_df, item_name):
    """
    Find matching rows based on Item Name, Item Group Label, or Form Label - Optimized version
    """
    # First try to match by Item Name (most common case)
    source_row = source_df[source_df['Item Name'] == item_name]
    target_row = target_df[target_df['Item Name'] == item_name]
    
    if not source_row.empty or not target_row.empty:
        return source_row, target_row, 'Item Name'
    
    # If Item Name not found, try Item Group Label
    if not source_row.empty:
        ig_label = source_row['Item Group Label'].values[0]
        target_row = target_df[target_df['Item Group Label'] == ig_label]
        if not target_row.empty:
            return source_row, target_row, 'Item Group Label'
    
    if not target_row.empty:
        ig_label = target_row['Item Group Label'].values[0]
        source_row = source_df[source_df['Item Group Label'] == ig_label]
        if not source_row.empty:
            return source_row, target_row, 'Item Group Label'
    
    # If still not found, try Form Label
    if not source_row.empty:
        form_label = source_row['Form Label'].values[0]
        target_row = target_df[target_df['Form Label'] == form_label]
        if not target_row.empty:
            return source_row, target_row, 'Form Label'
    
    if not target_row.empty:
        form_label = target_row['Form Label'].values[0]
        source_row = source_df[source_df['Form Label'] == form_label]
        if not source_row.empty:
            return source_row, target_row, 'Form Label'
    
    return source_row, target_row, None

def compare_values_vectorized(val1, val2):
    """Optimized compare two values and return status"""
    # Handle NaN values
    val1_nan = pd.isna(val1)
    val2_nan = pd.isna(val2)
    
    if val1_nan and val2_nan:
        return 'match', '✓', ''
    if val1_nan and not val2_nan:
        return 'missing_source', '⚠', 'Missing in Source'
    if not val1_nan and val2_nan:
        return 'missing_target', '⚠', 'Missing in Target'
    
    # Convert to string for comparison
    str1 = str(val1).strip()
    str2 = str(val2).strip()
    
    if str1 == str2:
        return 'match', '✓', ''
    else:
        return 'mismatch', '✗', 'Values differ'

def create_comparison_dataframe(source_row, target_row, source_df, source_name, target_name):
    """Create a comparison dataframe for display - only source columns - Optimized"""
    # Columns to ignore from comparison
    IGNORE_COLUMNS = ['Definition Last Modified', 'Relationship Last Modified']
    
    # Get only source columns
    source_columns = [col for col in source_df.columns if col not in IGNORE_COLUMNS]
    
    # Pre-allocate lists for better performance
    comparison_data = []
    
    # Get source and target values in bulk
    if not source_row.empty:
        source_values = source_row.iloc[0]
    else:
        source_values = pd.Series([None] * len(source_columns), index=source_columns)
    
    if not target_row.empty and len(target_row) > 0:
        target_values = target_row.iloc[0]
    else:
        target_values = pd.Series([None] * len(source_columns), index=source_columns)
    
    for col in source_columns:
        source_value = source_values.get(col)
        
        # Get target value if column exists in target
        if col in target_values.index:
            target_value = target_values.get(col)
        else:
            target_value = None
        
        status, symbol, note = compare_values_vectorized(source_value, target_value)
        
        # Replace generic terms with actual names
        if 'Source' in note:
            note = note.replace('Source', source_name)
        if 'Target' in note:
            note = note.replace('Target', target_name)
        
        comparison_data.append({
            'Column Name': col,
            f'{source_name} Value': source_value if not pd.isna(source_value) else '',
            f'{target_name} Value': target_value if not pd.isna(target_value) else f'Column not in {target_name}',
            'Status': symbol,
            'Match': status,
            'Note': note
        })
    
    return pd.DataFrame(comparison_data)

def highlight_differences(row):
    """Apply styling to highlight differences"""
    if row['Match'] == 'match':
        return ['background-color: #90EE90'] * len(row)  # Light green
    elif row['Match'] == 'mismatch':
        return ['background-color: #FFB6C1'] * len(row)  # Light red
    elif 'missing' in str(row['Match']):
        return ['background-color: #FFE4B5'] * len(row)  # Light orange
    else:
        return [''] * len(row)

def create_comprehensive_report(all_comparisons, source_name, target_name, source_df):
    """Create a comprehensive Excel report with summary and issue-only details - Optimized"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Comparison Summary - Build all data at once
        summary_data = []
        for item_name, data in all_comparisons.items():
            comp_df = data['comparison_df']
            
            # Vectorized counting
            match_mask = comp_df['Match'] == 'match'
            mismatch_mask = comp_df['Match'] == 'mismatch'
            missing_target_mask = comp_df['Match'] == 'missing_target'
            missing_source_mask = comp_df['Match'] == 'missing_source'
            
            total_cols = len(comp_df)
            matches = match_mask.sum()
            mismatches = mismatch_mask.sum()
            missing_target = missing_target_mask.sum()
            missing_source = missing_source_mask.sum()
            match_percentage = (matches / total_cols * 100) if total_cols > 0 else 0
            
            summary_data.append({
                'Item Name': item_name,
                'Match Type': data['match_type'],
                f'In {source_name}': '✓' if data['source_exists'] else '✗',
                f'In {target_name}': '✓' if data['target_exists'] else '✗',
                'Total Columns': total_cols,
                'Matches': matches,
                'Mismatches': mismatches,
                f'Missing in {target_name}': missing_target,
                f'Missing in {source_name}': missing_source,
                'Match %': f"{match_percentage:.1f}%"
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Comparison Summary', index=False)
        
        # Sheet 2: Issues Only (simplified format) - Optimized
        issues_data = []
        
        # Pre-fetch source rows for all items to avoid repeated lookups
        source_info_cache = {}
        for item_name in all_comparisons.keys():
            source_row = source_df[source_df['Item Name'] == item_name]
            if not source_row.empty:
                source_info_cache[item_name] = {
                    'form_name': source_row['Form Name'].values[0] if 'Form Name' in source_row.columns else '',
                    'form_label': source_row['Form Label'].values[0] if 'Form Label' in source_row.columns else '',
                    'form_short_label': source_row['Form Short Label'].values[0] if 'Form Short Label' in source_row.columns else '',
                    'item_group_name': source_row['Item Group Name'].values[0] if 'Item Group Name' in source_row.columns else '',
                    'item_group_label': source_row['Item Group Label'].values[0] if 'Item Group Label' in source_row.columns else ''
                }
        
        for item_name, data in all_comparisons.items():
            comp_df = data['comparison_df']
            matches = (comp_df['Match'] == 'match').sum()
            match_rate = (matches / len(comp_df) * 100) if len(comp_df) > 0 else 0
            
            # Only include items with less than 100% match
            if match_rate < 100.0:
                # Get cached source info
                if item_name in source_info_cache:
                    info = source_info_cache[item_name]
                    form_name = info['form_name']
                    form_label = info['form_label']
                    form_short_label = info['form_short_label']
                    item_group_name = info['item_group_name']
                    item_group_label = info['item_group_label']
                else:
                    form_name = ''
                    form_label = ''
                    form_short_label = ''
                    item_group_name = ''
                    item_group_label = ''
                
                # Find all rows with issues (not 100% match) - vectorized
                issue_rows = comp_df[comp_df['Match'] != 'match']
                
                for _, row in issue_rows.iterrows():
                    column_name = row['Column Name']
                    source_value = row[f'{source_name} Value']
                    target_value = row[f'{target_name} Value']
                    
                    # Determine issue type
                    if row['Match'] == 'mismatch':
                        issue_type = f"{column_name}: Value mismatch ({source_name}: '{source_value}' vs {target_name}: '{target_value}')"
                    elif row['Match'] == 'missing_target':
                        issue_type = f"{column_name}: Missing in {target_name} ({source_name} has: '{source_value}')"
                    elif row['Match'] == 'missing_source':
                        issue_type = f"{column_name}: Missing in {source_name} ({target_name} has: '{target_value}')"
                    else:
                        issue_type = f"{column_name}: {row['Note']}"
                    
                    issues_data.append({
                        'Item Name': item_name,
                        'Form Name': form_name,
                        'Form Label': form_label,
                        'Form Short Label': form_short_label,
                        'Item Group Name': item_group_name,
                        'Item Group Label': item_group_label,
                        'Issue Type': issue_type
                    })
        
        if issues_data:
            issues_df = pd.DataFrame(issues_data)
            issues_df.to_excel(writer, sheet_name='Issues Only', index=False)
            
            # Access the workbook and worksheet to format
            workbook = writer.book
            worksheet = workbook['Issues Only']
            
            # Set text wrapping to False and alignment for all cells
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, 
                                          min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical='top', horizontal='left')
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                    except:
                        pass
                adjusted_width = min(max_length + 2, 80)  # Cap at 80 characters for Issue Type
                worksheet.column_dimensions[column_letter].width = adjusted_width
        else:
            # If all items have 100% match, create an empty sheet with a message
            empty_df = pd.DataFrame({
                'Message': ['All items have 100% match rate. No discrepancies to report.']
            })
            empty_df.to_excel(writer, sheet_name='Issues Only', index=False)
    
    output.seek(0)
    return output

@st.cache_data
def get_unique_items(df, column_name='Item Name'):
    """Cached function to get unique items"""
    return set(df[column_name].dropna().unique())

def main():
    st.set_page_config(page_title="PTD vs SDS Comparison Tool", layout="wide")
    
    st.title("📊 PTD vs SDS Comparison Tool")
    
    # Initialize session state
    if 'selected_items' not in st.session_state:
        st.session_state.selected_items = []
    if 'ptd_df' not in st.session_state:
        st.session_state.ptd_df = None
    if 'sds_df' not in st.session_state:
        st.session_state.sds_df = None
    if 'comparison_direction' not in st.session_state:
        st.session_state.comparison_direction = "PTD → SDS (Compare PTD columns against SDS)"
    if 'all_comparisons' not in st.session_state:
        st.session_state.all_comparisons = None
    if 'comparison_complete' not in st.session_state:
        st.session_state.comparison_complete = False
    if 'input_method' not in st.session_state:
        st.session_state.input_method = "📋 Copy-Paste from Excel"
    
    # Input method selection
    st.markdown("---")
    st.subheader("📥 Select Input Method")
    input_method = st.radio(
        "How would you like to provide the data?",
        ["📋 Copy-Paste from Excel", "📁 Upload Excel Files"],
        horizontal=True,
        help="Choose whether to copy-paste data directly or upload Excel files"
    )
    st.session_state.input_method = input_method
    
    # Comparison direction selection
    st.markdown("---")
    comparison_direction = st.radio(
        "🔄 Select Comparison Direction:",
        [
            "PTD → SDS (Compare PTD columns against SDS)",
            "SDS → PTD (Compare SDS columns against PTD)",
            "SDS → SDS (Compare SDS columns against SDS)"
        ],
        index=["PTD → SDS (Compare PTD columns against SDS)",
               "SDS → PTD (Compare SDS columns against PTD)",
               "SDS → SDS (Compare SDS columns against SDS)"].index(st.session_state.comparison_direction),
        help="Choose which file's columns to use as the basis for comparison"
    )
    
    # Update session state
    st.session_state.comparison_direction = comparison_direction
    
    # Determine labels based on comparison direction
    if "SDS → SDS" in comparison_direction:
        left_label = "SDS1"
        right_label = "SDS2"
        left_key = "sds1"
        right_key = "sds2"
    else:
        left_label = "PTD"
        right_label = "SDS"
        left_key = "ptd"
        right_key = "sds"
    
    st.markdown("---")
    
    # Data input section based on selected method
    if input_method == "📋 Copy-Paste from Excel":
        st.markdown("#### 📋 How to Copy from Excel:")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**1️⃣ Select Data**")
            st.caption("Select all rows including headers in Excel")
        with col2:
            st.markdown("**2️⃣ Copy**")
            st.caption("Press Ctrl+C (Windows)")
        with col3:
            st.markdown("**3️⃣ Paste**")
            st.caption("Click in text area and press Ctrl+V")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader(f"📄 {left_label} Data")
            st.caption("Copy data from Excel and paste here")
            left_text = st.text_area(
                f"Paste {left_label} data here:",
                height=300,
                placeholder=f"Select and copy your {left_label} data from Excel, then paste here...",
                key=f"{left_key}_paste",
                help=f"Select all {left_label} data in Excel (including headers), copy (Ctrl+C), and paste here (Ctrl+V)"
            )
            
            if left_text:
                with st.spinner(f"Parsing {left_label} data..."):
                    left_df = parse_pasted_data(left_text)
                    if left_df is not None:
                        # Process PTD dataframe if it's PTD (but not SDS1)
                        if left_label == "PTD":
                            processed_result = process_ptd_dataframe(left_df)
                            if isinstance(processed_result, tuple):
                                left_df, original_count, filtered_count = processed_result
                                st.session_state.ptd_df = left_df
                                st.success(f"✅ {left_label} data loaded: {len(left_df)} rows, {len(left_df.columns)} columns")
                                if original_count != filtered_count:
                                    st.info(f"ℹ️ Filtered from {original_count} to {filtered_count} rows (keeping only 'Y' in 'Used in trial' column)")
                            else:
                                st.session_state.ptd_df = processed_result
                        else:
                            st.session_state.ptd_df = left_df
                            st.success(f"✅ {left_label} data loaded: {len(left_df)} rows, {len(left_df.columns)} columns")
            elif st.session_state.ptd_df is not None:
                st.success(f"✅ {left_label} data loaded: {len(st.session_state.ptd_df)} rows, {len(st.session_state.ptd_df.columns)} columns")
        
        with col2:
            st.subheader(f"📄 {right_label} Data")
            st.caption("Copy data from Excel and paste here")
            right_text = st.text_area(
                f"Paste {right_label} data here:",
                height=300,
                placeholder=f"Select and copy your {right_label} data from Excel, then paste here...",
                key=f"{right_key}_paste",
                help=f"Select all {right_label} data in Excel (including headers), copy (Ctrl+C), and paste here (Ctrl+V)"
            )
            
            if right_text:
                with st.spinner(f"Parsing {right_label} data..."):
                    right_df = parse_pasted_data(right_text)
                    if right_df is not None:
                        # Process PTD dataframe only if it's actual PTD (in SDS → PTD comparison)
                        # For SDS2, don't process as PTD
                        if right_label == "PTD":
                            processed_result = process_ptd_dataframe(right_df)
                            if isinstance(processed_result, tuple):
                                right_df, original_count, filtered_count = processed_result
                                st.session_state.sds_df = right_df
                                st.success(f"✅ {right_label} data loaded: {len(right_df)} rows, {len(right_df.columns)} columns")
                                if original_count != filtered_count:
                                    st.info(f"ℹ️ Filtered from {original_count} to {filtered_count} rows (keeping only 'Y' in 'Used in trial' column)")
                            else:
                                st.session_state.sds_df = processed_result
                        else:
                            st.session_state.sds_df = right_df
                            st.success(f"✅ {right_label} data loaded: {len(right_df)} rows, {len(right_df.columns)} columns")
            elif st.session_state.sds_df is not None:
                st.success(f"✅ {right_label} data loaded: {len(st.session_state.sds_df)} rows, {len(st.session_state.sds_df.columns)} columns")
    
    else:  # Upload Excel Files
        st.markdown("#### 📁 Upload Excel Files:")
        st.info("ℹ️ **Note:** The tool will read data from the '**Form Definitions**' sheet in each uploaded file.")
        if left_label == "PTD":
            st.warning("⚠️ **For PTD files:** First row will be skipped and second row will be used as column headers.")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader(f"📄 {left_label} File")
            st.caption("Upload Excel file (.xlsx or .xls)")
            left_file = st.file_uploader(
                f"Upload {left_label} Excel file:",
                type=['xlsx', 'xls'],
                key=f"{left_key}_upload",
                help=f"Select your {left_label} Excel file to upload. Data will be read from 'Form Definitions' sheet."
            )
            
            if left_file is not None:
                # Determine if this is a PTD file (only actual PTD, not SDS1)
                is_ptd_file = (left_label == "PTD")
                
                with st.spinner(f"Reading {left_label} file from 'Form Definitions' sheet..."):
                    left_df = parse_uploaded_file(left_file, sheet_name='Form Definitions', is_ptd=is_ptd_file)
                    if left_df is not None:
                        # Process PTD dataframe only if it's actual PTD
                        if is_ptd_file:
                            processed_result = process_ptd_dataframe(left_df)
                            if isinstance(processed_result, tuple):
                                left_df, original_count, filtered_count = processed_result
                                st.session_state.ptd_df = left_df
                                st.success(f"✅ {left_label} data loaded from 'Form Definitions' sheet (first row skipped): {len(left_df)} rows, {len(left_df.columns)} columns")
                                if original_count != filtered_count:
                                    st.info(f"ℹ️ Filtered from {original_count} to {filtered_count} rows (keeping only 'Y' in 'Used in trial' column)")
                            else:
                                st.session_state.ptd_df = processed_result
                                st.success(f"✅ {left_label} data loaded from 'Form Definitions' sheet (first row skipped): {len(processed_result)} rows, {len(processed_result.columns)} columns")
                        else:
                            # For SDS1, don't skip row and don't process as PTD
                            st.session_state.ptd_df = left_df
                            st.success(f"✅ {left_label} data loaded from 'Form Definitions' sheet: {len(left_df)} rows, {len(left_df.columns)} columns")
                    else:
                        st.session_state.ptd_df = None
            elif st.session_state.ptd_df is not None:
                st.success(f"✅ {left_label} data loaded: {len(st.session_state.ptd_df)} rows, {len(st.session_state.ptd_df.columns)} columns")
        
        with col2:
            st.subheader(f"📄 {right_label} File")
            st.caption("Upload Excel file (.xlsx or .xls)")
            right_file = st.file_uploader(
                f"Upload {right_label} Excel file:",
                type=['xlsx', 'xls'],
                key=f"{right_key}_upload",
                help=f"Select your {right_label} Excel file to upload. Data will be read from 'Form Definitions' sheet."
            )
            
            if right_file is not None:
                # Only PTD (in SDS→PTD) should skip first row, not SDS2
                is_ptd_file = (right_label == "PTD")
                
                with st.spinner(f"Reading {right_label} file from 'Form Definitions' sheet..."):
                    right_df = parse_uploaded_file(right_file, sheet_name='Form Definitions', is_ptd=is_ptd_file)
                    if right_df is not None:
                        # Process PTD dataframe only for actual PTD, not SDS2
                        if is_ptd_file:
                            processed_result = process_ptd_dataframe(right_df)
                            if isinstance(processed_result, tuple):
                                right_df, original_count, filtered_count = processed_result
                                st.session_state.sds_df = right_df
                                st.success(f"✅ {right_label} data loaded from 'Form Definitions' sheet (first row skipped): {len(right_df)} rows, {len(right_df.columns)} columns")
                                if original_count != filtered_count:
                                    st.info(f"ℹ️ Filtered from {original_count} to {filtered_count} rows (keeping only 'Y' in 'Used in trial' column)")
                            else:
                                st.session_state.sds_df = processed_result
                                st.success(f"✅ {right_label} data loaded from 'Form Definitions' sheet (first row skipped): {len(processed_result)} rows, {len(processed_result.columns)} columns")
                        else:
                            # For regular SDS (including SDS2), don't skip row
                            st.session_state.sds_df = right_df
                            st.success(f"✅ {right_label} data loaded from 'Form Definitions' sheet: {len(right_df)} rows, {len(right_df.columns)} columns")
                    else:
                        st.session_state.sds_df = None
            elif st.session_state.sds_df is not None:
                st.success(f"✅ {right_label} data loaded: {len(st.session_state.sds_df)} rows, {len(st.session_state.sds_df.columns)} columns")
    
    # Use data from session state
    ptd_df = st.session_state.ptd_df
    sds_df = st.session_state.sds_df
    
    # Check which data is available
    has_left = ptd_df is not None
    has_right = sds_df is not None
    
    if has_left and has_right:
        st.markdown("---")
        st.success("✅ Both datasets loaded successfully! Ready to compare.")
        
        # Set source and target based on selection
        if "PTD → SDS" in comparison_direction:
            source_df = ptd_df
            target_df = sds_df
            source_name = "PTD"
            target_name = "SDS"
        elif "SDS → PTD" in comparison_direction:
            source_df = sds_df
            target_df = ptd_df
            source_name = "SDS"
            target_name = "PTD"
        elif "SDS → SDS" in comparison_direction:
            source_df = ptd_df  # SDS1
            target_df = sds_df  # SDS2
            source_name = "SDS1"
            target_name = "SDS2"
        
        # Display basic info
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"📊 {source_name} Records: {len(source_df)}")
        with col2:
            st.info(f"📊 {target_name} Records: {len(target_df)}")
        with col3:
            # Count columns excluding ignored ones
            IGNORE_COLUMNS = ['Definition Last Modified', 'Relationship Last Modified']
            source_cols = [col for col in source_df.columns if col not in IGNORE_COLUMNS]
            st.info(f"📋 {source_name} Columns: {len(source_cols)}")
        
        st.markdown("---")
        
        # Get all unique Item Names from both dataframes - using cached function
        source_items = get_unique_items(source_df, 'Item Name')
        target_items = get_unique_items(target_df, 'Item Name')
        all_items = sorted(list(source_items.union(target_items)))
        
        # Show items only in source or only in target
        only_in_source = source_items - target_items
        only_in_target = target_items - source_items
        
        if only_in_source or only_in_target:
            with st.expander("⚠️ View Items Not in Both Files"):
                col1, col2 = st.columns(2)
                with col1:
                    if only_in_source:
                        st.warning(f"**Only in {source_name} ({len(only_in_source)} items):**")
                        st.write(list(only_in_source))
                with col2:
                    if only_in_target:
                        st.warning(f"**Only in {target_name} ({len(only_in_target)} items):**")
                        st.write(list(only_in_target))
        
        st.markdown("---")
        
        # Multi-select item comparison
        st.subheader("Select Items to Compare")
        
        # Quick selection options
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("✅ Select All", use_container_width=True):
                st.session_state.selected_items = all_items
        
        with col2:
            if st.button("❌ Clear Selection", use_container_width=True):
                st.session_state.selected_items = []
        
        with col3:
            if st.button(f"🔍 Select from {source_name} only", use_container_width=True):
                st.session_state.selected_items = sorted(list(source_items))
        
        # Multi-select widget
        selected_items = st.multiselect(
            "Select Item Names to Compare:",
            options=all_items,
            default=st.session_state.selected_items,
            help="You can select multiple items to compare at once"
        )
        
        # Update session state with current selection
        st.session_state.selected_items = selected_items
        
        # Show selection count
        if selected_items:
            st.info(f"📋 Selected {len(selected_items)} item(s) for comparison")
        else:
            st.warning("⚠️ No items selected. Please select at least one item to compare.")
        
        if selected_items and st.button("🔍 Compare Selected Items", type="primary", use_container_width=True):
            all_comparisons = {}
            
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Batch processing with progress updates every 50 items
            batch_size = 50
            total_items = len(selected_items)
            
            for idx, item_name in enumerate(selected_items):
                # Update progress less frequently for better performance
                if idx % batch_size == 0 or idx == total_items - 1:
                    status_text.text(f"Comparing {idx+1}/{total_items}: {item_name}")
                    progress_bar.progress((idx + 1) / total_items)
                
                source_row, target_row, match_type = find_matching_rows(source_df, target_df, item_name)
                
                if source_row.empty and target_row.empty:
                    # Skip warning for performance - can log to list if needed
                    pass
                else:
                    comparison_df = create_comparison_dataframe(source_row, target_row, source_df, source_name, target_name)
                    all_comparisons[item_name] = {
                        'comparison_df': comparison_df,
                        'match_type': match_type,
                        'source_exists': not source_row.empty,
                        'target_exists': not target_row.empty
                    }
            
            status_text.text("✅ Comparison complete!")
            progress_bar.progress(1.0)
            
            # Store in session state
            st.session_state.all_comparisons = all_comparisons
            st.session_state.comparison_complete = True
            st.session_state.source_name = source_name
            st.session_state.target_name = target_name
            st.session_state.source_df = source_df  # Store source dataframe for report generation
            
            st.success(f"✅ Completed comparison for {len(all_comparisons)} items!")
            
            st.markdown("---")
            
            # Display summary for selected items
            st.subheader("📊 Comparison Summary")
            
            summary_data = []
            for item_name, data in all_comparisons.items():
                comp_df = data['comparison_df']
                
                # Vectorized counting
                match_mask = comp_df['Match'] == 'match'
                total_cols = len(comp_df)
                matches = match_mask.sum()
                mismatches = (comp_df['Match'] == 'mismatch').sum()
                missing_target = (comp_df['Match'] == 'missing_target').sum()
                missing_source = (comp_df['Match'] == 'missing_source').sum()
                match_percentage = (matches / total_cols * 100) if total_cols > 0 else 0
                
                summary_data.append({
                    'Item Name': item_name,
                    'Match Type': data['match_type'],
                    f'In {source_name}': '✓' if data['source_exists'] else '✗',
                    f'In {target_name}': '✓' if data['target_exists'] else '✗',
                    'Total Columns': total_cols,
                    'Matches ✅': matches,
                    'Mismatches ❌': mismatches,
                    f'Missing in {target_name} ⚠️': missing_target,
                    f'Missing in {source_name} ⚠️': missing_source,
                    'Match %': f"{match_percentage:.1f}%"
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, height=300)
            
            # Filter out items with 100% match rate for detailed results
            items_to_display = {}
            items_with_100_match = []
            
            for item_name, data in all_comparisons.items():
                comp_df = data['comparison_df']
                matches = (comp_df['Match'] == 'match').sum()
                match_rate = (matches / len(comp_df) * 100) if len(comp_df) > 0 else 0
                
                if match_rate == 100.0:
                    items_with_100_match.append(item_name)
                else:
                    items_to_display[item_name] = data
            
            # Display individual comparisons (excluding 100% matches)
            st.subheader("📋 Detailed Comparison Results")
            
            # Show info about hidden items
            if items_with_100_match:
                st.success(f"✅ {len(items_with_100_match)} item(s) with 100% match are hidden from detailed view")
                with st.expander(f"View {len(items_with_100_match)} items with 100% match"):
                    st.write(items_with_100_match)
            
            if len(items_to_display) == 0:
                st.info("🎉 All items have 100% match rate! No discrepancies to display.")
            else:
                st.info(f"Showing {len(items_to_display)} item(s) with discrepancies")
                
                # Limit tabs to first 20 items for performance
                if len(items_to_display) > 20:
                    st.warning(f"⚠️ Displaying first 20 items in tabs. Download the report to see all {len(items_to_display)} items with discrepancies.")
                    items_to_show = dict(list(items_to_display.items())[:20])
                else:
                    items_to_show = items_to_display
                
                # Tab selection for each item (excluding 100% matches)
                tabs = st.tabs([item_name for item_name in items_to_show.keys()])
                
                for tab, (item_name, data) in zip(tabs, items_to_show.items()):
                    with tab:
                        comparison_df = data['comparison_df']
                        
                        # Show match information
                        col1, col2 = st.columns(2)
                        with col1:
                            if data['match_type']:
                                st.success(f"✅ Matched by: **{data['match_type']}**")
                            st.info(f"In {source_name}: {'✓' if data['source_exists'] else '✗'}")
                        with col2:
                            matches = (comparison_df['Match'] == 'match').sum()
                            match_rate = (matches/len(comparison_df)*100)
                            st.metric("Match Rate", f"{match_rate:.1f}%")
                            st.info(f"In {target_name}: {'✓' if data['target_exists'] else '✗'}")
                        
                        # Summary counts for this item
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            total_rows = len(comparison_df)
                            st.metric("Total Rows", total_rows)
                        with col2:
                            match_count = (comparison_df['Match'] == 'match').sum()
                            st.metric("✅ Matches", match_count)
                        with col3:
                            mismatch_count = (comparison_df['Match'] == 'mismatch').sum()
                            st.metric("❌ Mismatches", mismatch_count)
                        with col4:
                            missing_count = ((comparison_df['Match'] == 'missing_source') | 
                                           (comparison_df['Match'] == 'missing_target')).sum()
                            st.metric("⚠️ Missing", missing_count)
                        
                        st.markdown("---")
                        
                        # Display full comparison table
                        st.markdown("##### Complete Comparison Table")
                        styled_df = comparison_df.style.apply(highlight_differences, axis=1)
                        st.dataframe(styled_df, use_container_width=True, height=500)
            
            # Legend
            st.markdown("---")
            st.markdown("**Legend:**")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("🟢 **Green**: Values match")
            with col2:
                st.markdown("🔴 **Red**: Values differ")
            with col3:
                st.markdown("🟡 **Orange**: Value missing in one file")
        
        # Download Report Section (always visible if comparison is complete)
        if st.session_state.comparison_complete and st.session_state.all_comparisons:
            st.markdown("---")
            st.markdown("---")
            st.subheader("📥 Download Comprehensive Report")
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.info("📊 The report includes:\n"
                       "- **Sheet 1**: Comparison Summary (all items)\n"
                       "- **Sheet 2**: Issues Only (simplified format with Form Name, Form Label, Form Short Label, Item Group Name, Item Group Label, and Issue Type)")
            
            with col2:
                # Generate report with progress indicator
                with st.spinner("Generating Excel report..."):
                    report_output = create_comprehensive_report(
                        st.session_state.all_comparisons,
                        st.session_state.source_name,
                        st.session_state.target_name,
                        st.session_state.source_df
                    )
                
                st.download_button(
                    label="📥 Download Report",
                    data=report_output,
                    file_name=f"comparison_report_{st.session_state.source_name}_vs_{st.session_state.target_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
    
    elif not has_left and not has_right:
        st.warning(f"⚠️ Please {'paste' if input_method == '📋 Copy-Paste from Excel' else 'upload'} both {left_label} and {right_label} data to start comparing.")
    elif not has_left:
        st.warning(f"⚠️ Please {'paste' if input_method == '📋 Copy-Paste from Excel' else 'upload'} {left_label} data.")
    else:
        st.warning(f"⚠️ Please {'paste' if input_method == '📋 Copy-Paste from Excel' else 'upload'} {right_label} data.")

if __name__ == "__main__":
    main()
