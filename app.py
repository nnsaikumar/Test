import pandas as pd
import streamlit as st
from io import BytesIO, StringIO
import openpyxl
from openpyxl.styles import Alignment
import numpy as np
from functools import lru_cache
import time

# Set page config at the very top
st.set_page_config(page_title="PTD vs SDS Comparison Tool", layout="wide")

def parse_pasted_data(pasted_text):
    """Parse pasted Excel data (tab-separated) into DataFrame"""
    try:
        df = pd.read_csv(StringIO(pasted_text), sep='\t', engine='python')
        return df
    except Exception as e:
        st.error(f"Error parsing pasted data: {e}")
        return None

def parse_uploaded_file(uploaded_file, sheet_name='Form Definitions', is_ptd=False):
    """Parse uploaded Excel file into DataFrame"""
    try:
        if is_ptd and sheet_name == 'Form Definitions':
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl', header=1)
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
        return df
    except ValueError as e:
        st.error(f"Error: Sheet '{sheet_name}' not found. Available sheets: {pd.ExcelFile(uploaded_file).sheet_names}")
        return None
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def convert_decimal_column(df):
    """
    Convert 'Decimal' column values to proper decimal format.
    Examples: "1" -> "1.0", "2" -> "2.0", "1.0" -> "1.0"
    """
    if df is None:
        return df
    
    # Find the Decimal column (handle various spacing)
    decimal_column = None
    for col in df.columns:
        if col.strip().lower() == 'decimal':
            decimal_column = col
            break
    
    if decimal_column:
        def format_decimal(value):
            """Format a single decimal value"""
            if pd.isna(value) or value == '' or value is None:
                return value
            
            try:
                # Convert to string first
                str_value = str(value).strip()
                
                # If empty after strip, return as is
                if not str_value:
                    return value
                
                # Try to convert to float
                float_value = float(str_value)
                
                # Check if it's a whole number
                if float_value == int(float_value):
                    # Return with .0
                    return f"{int(float_value)}.0"
                else:
                    # Return as is (already has decimal)
                    return str(float_value)
            except (ValueError, TypeError):
                # If conversion fails, return original value
                return value
        
        # Apply formatting
        df[decimal_column] = df[decimal_column].apply(format_decimal)
    
    return df

def process_codelists(df):
    """
    Process Codelists sheet: filter out rows with blank Choice Code
    """
    if df is None:
        return None
    
    # Find Choice Code column (handle various spacing)
    choice_code_col = None
    for col in df.columns:
        if col.strip().lower() == 'choice code':
            choice_code_col = col
            break
    
    if choice_code_col:
        original_count = len(df)
        # Filter out blank Choice Code values
        df = df[df[choice_code_col].notna()].copy()
        df = df[df[choice_code_col].astype(str).str.strip() != ''].copy()
        filtered_count = len(df)
        return df, original_count, filtered_count
    else:
        return df, len(df), len(df)

def process_ptd_dataframe(df):
    """
    Process PTD dataframe: 
    1. Filter by 'Used in trial' (Y, Yes, Mod)
    2. Remove specific columns
    3. Convert Decimal column format
    """
    if df is None:
        return None
    
    original_count = len(df)
    
    # Step 1: Filter by trial column (Y, Yes, Mod variations)
    trial_column_names = [
        'Used in trial (Y, N, Mod)',
        'Used in trial (Y, N, Mod) ',
        ' Used in trial (Y, N, Mod)',
        ' Used in trial (Y, N, Mod) ',
        'Used in trial',
    ]
    
    trial_column = next((col for col in trial_column_names if col in df.columns), None)
    
    if trial_column:
        # Transform once, then filter for Y, Yes, or Mod
        normalized_column = df[trial_column].astype(str).str.strip().str.upper()
        df = df[normalized_column.isin(['Y', 'YES', 'MOD'])].copy()
        filtered_count = len(df)
    else:
        st.warning("⚠️ Column 'Used in trial (Y, N, Mod)' not found. Skipping filter.")
        filtered_count = original_count
    
    # Step 2: Remove specific columns (handles various spacing)
    columns_to_remove_patterns = [
        'Modification comments + Highlight Cells where change made',
        'Library source',
        'Used in trial (Y, N, Mod)'
    ]
    
    columns_to_remove = []
    for col in df.columns:
        col_stripped = col.strip()
        if col_stripped in columns_to_remove_patterns:
            columns_to_remove.append(col)
    
    df = df.drop(columns=columns_to_remove, errors='ignore')
    
    # Step 3: Convert Decimal column format
    df = convert_decimal_column(df)
    
    return df, original_count, filtered_count

def process_sds_dataframe(df):
    """
    Process SDS dataframe:
    Convert Decimal column format only
    """
    if df is None:
        return df
    
    # Convert Decimal column format
    df = convert_decimal_column(df)
    
    return df

def compare_codelists(codelist_name, source_codelists_df, target_codelists_df, source_name, target_name):
    """
    Compare codelists between source and target for a given codelist name.
    Returns comparison result dictionary.
    """
    if source_codelists_df is None or target_codelists_df is None:
        return None
    
    if pd.isna(codelist_name) or str(codelist_name).strip() == '':
        return None
    
    # Find Name column (handle various spacing)
    source_name_col = None
    target_name_col = None
    
    for col in source_codelists_df.columns:
        if col.strip().lower() == 'name':
            source_name_col = col
            break
    
    for col in target_codelists_df.columns:
        if col.strip().lower() == 'name':
            target_name_col = col
            break
    
    if not source_name_col or not target_name_col:
        return None
    
    # Filter codelists by name
    source_codelist = source_codelists_df[
        source_codelists_df[source_name_col].astype(str).str.strip() == str(codelist_name).strip()
    ].copy()
    
    target_codelist = target_codelists_df[
        target_codelists_df[target_name_col].astype(str).str.strip() == str(codelist_name).strip()
    ].copy()
    
    if source_codelist.empty and target_codelist.empty:
        return {
            'status': 'not_found',
            'message': f"Codelist '{codelist_name}' not found in either file",
            'matches': 0,
            'mismatches': 0,
            'details': []
        }
    
    if source_codelist.empty:
        return {
            'status': 'missing_source',
            'message': f"Codelist '{codelist_name}' not found in {source_name}",
            'matches': 0,
            'mismatches': len(target_codelist),
            'details': []
        }
    
    if target_codelist.empty:
        return {
            'status': 'missing_target',
            'message': f"Codelist '{codelist_name}' not found in {target_name}",
            'matches': 0,
            'mismatches': len(source_codelist),
            'details': []
        }
    
    # Find columns for comparison
    cols_to_find = ['choice code', 'choice label']
    source_cols = {}
    target_cols = {}
    
    for col in source_codelist.columns:
        col_lower = col.strip().lower()
        if col_lower in cols_to_find:
            source_cols[col_lower] = col
    
    for col in target_codelist.columns:
        col_lower = col.strip().lower()
        if col_lower in cols_to_find:
            target_cols[col_lower] = col
    
    # Compare row by row
    matches = 0
    mismatches = 0
    details = []
    
    # Create comparison key using Choice Code
    if 'choice code' in source_cols and 'choice code' in target_cols:
        source_codes = set(source_codelist[source_cols['choice code']].astype(str).str.strip())
        target_codes = set(target_codelist[target_cols['choice code']].astype(str).str.strip())
        
        all_codes = source_codes.union(target_codes)
        
        for code in all_codes:
            source_rows = source_codelist[
                source_codelist[source_cols['choice code']].astype(str).str.strip() == code
            ]
            target_rows = target_codelist[
                target_codelist[target_cols['choice code']].astype(str).str.strip() == code
            ]
            
            if source_rows.empty and not target_rows.empty:
                mismatches += 1
                target_label = target_rows.iloc[0][target_cols['choice label']] if 'choice label' in target_cols else ''
                details.append({
                    'name': codelist_name,
                    'choice_code': code,
                    'source_label': 'Missing',
                    'target_label': target_label,
                    'status': 'missing_source'
                })
            elif not source_rows.empty and target_rows.empty:
                mismatches += 1
                source_label = source_rows.iloc[0][source_cols['choice label']] if 'choice label' in source_cols else ''
                details.append({
                    'name': codelist_name,
                    'choice_code': code,
                    'source_label': source_label,
                    'target_label': 'Missing',
                    'status': 'missing_target'
                })
            elif not source_rows.empty and not target_rows.empty:
                source_label = source_rows.iloc[0][source_cols['choice label']] if 'choice label' in source_cols else ''
                target_label = target_rows.iloc[0][target_cols['choice label']] if 'choice label' in target_cols else ''
                
                if str(source_label).strip() == str(target_label).strip():
                    matches += 1
                    details.append({
                        'name': codelist_name,
                        'choice_code': code,
                        'source_label': source_label,
                        'target_label': target_label,
                        'status': 'match'
                    })
                else:
                    mismatches += 1
                    details.append({
                        'name': codelist_name,
                        'choice_code': code,
                        'source_label': source_label,
                        'target_label': target_label,
                        'status': 'mismatch'
                    })
    
    status = 'match' if mismatches == 0 else 'mismatch'
    
    return {
        'status': status,
        'message': f"Codelist '{codelist_name}': {matches} matches, {mismatches} mismatches",
        'matches': matches,
        'mismatches': mismatches,
        'details': details
    }

def find_matching_rows_optimized(source_df, target_df, item_name, source_dict, target_dict):
    """
    Optimized matching using pre-built dictionaries
    """
    # Try Item Name first
    source_row = source_dict.get(item_name)
    target_row = target_dict.get(item_name)
    
    if source_row is not None or target_row is not None:
        return source_row, target_row, 'Item Name'
    
    return None, None, None

def build_lookup_dictionaries(df):
    """Build lookup dictionaries for faster access"""
    item_dict = {}
    for idx, row in df.iterrows():
        item_name = row.get('Item Name')
        if pd.notna(item_name):
            item_dict[item_name] = row
    return item_dict

@st.cache_data
def compare_values_cached(val1, val2):
    """Cached comparison function"""
    val1_nan = pd.isna(val1)
    val2_nan = pd.isna(val2)
    
    if val1_nan and val2_nan:
        return 'match', '✓', ''
    if val1_nan and not val2_nan:
        return 'missing_source', '⚠', 'Missing in Source'
    if not val1_nan and val2_nan:
        return 'missing_target', '⚠', 'Missing in Target'
    
    str1 = str(val1).strip()
    str2 = str(val2).strip()
    
    if str1 == str2:
        return 'match', '✓', ''
    else:
        return 'mismatch', '✗', 'Values differ'

def create_comparison_dataframe_fast(source_row, target_row, source_columns, source_name, target_name, 
                                     codelist_comparison=None):
    """Vectorized comparison creation with codelist comparison"""
    IGNORE_COLUMNS = {'Definition Last Modified', 'Relationship Last Modified'}
    
    comparison_data = []
    
    if source_row is None:
        source_values = pd.Series([None] * len(source_columns), index=source_columns)
    else:
        source_values = source_row
    
    if target_row is None:
        target_values = pd.Series([None] * len(source_columns), index=source_columns)
    else:
        target_values = target_row
    
    for col in source_columns:
        if col in IGNORE_COLUMNS:
            continue
            
        source_value = source_values.get(col)
        target_value = target_values.get(col) if col in target_values.index else None
        
        status, symbol, note = compare_values_cached(source_value, target_value)
        
        note = note.replace('Source', source_name).replace('Target', target_name)
        
        comparison_data.append({
            'Column Name': col,
            f'{source_name} Value': source_value if not pd.isna(source_value) else '',
            f'{target_name} Value': target_value if not pd.isna(target_value) else f'Column not in {target_name}',
            'Status': symbol,
            'Match': status,
            'Note': note
        })
    
    # Add codelist comparison if available
    if codelist_comparison and codelist_comparison['details']:
        for detail in codelist_comparison['details']:
            if detail['status'] == 'match':
                symbol = '✓'
                status = 'match'
                note = ''
            elif detail['status'] == 'mismatch':
                symbol = '✗'
                status = 'mismatch'
                note = 'Codelist labels differ'
            elif detail['status'] == 'missing_source':
                symbol = '⚠'
                status = 'missing_source'
                note = f'Choice Code missing in {source_name}'
            else:  # missing_target
                symbol = '⚠'
                status = 'missing_target'
                note = f'Choice Code missing in {target_name}'
            
            comparison_data.append({
                'Column Name': f"Codelist: {detail['name']} | Code: {detail['choice_code']}",
                f'{source_name} Value': detail['source_label'],
                f'{target_name} Value': detail['target_label'],
                'Status': symbol,
                'Match': status,
                'Note': note
            })
    
    return pd.DataFrame(comparison_data)

def highlight_differences(row):
    """Apply styling to highlight differences"""
    match_type = row['Match']
    if match_type == 'match':
        return ['background-color: #90EE90'] * len(row)
    elif match_type == 'mismatch':
        return ['background-color: #FFB6C1'] * len(row)
    elif 'missing' in match_type:
        return ['background-color: #FFE4B5'] * len(row)
    else:
        return [''] * len(row)

def create_comprehensive_report_fast(all_comparisons, source_name, target_name, source_df):
    """Optimized report generation"""
    output = BytesIO()
    
    # Pre-build summary data
    summary_data = []
    issues_data = []
    
    # Pre-fetch source info for all items
    source_info_cache = {}
    for item_name in all_comparisons.keys():
        source_row = source_df[source_df['Item Name'] == item_name]
        if not source_row.empty:
            row = source_row.iloc[0]
            source_info_cache[item_name] = {
                'form_name': row.get('Form Name', ''),
                'form_label': row.get('Form Label', ''),
                'form_short_label': row.get('Form Short Label', ''),
                'item_group_name': row.get('Item Group Name', ''),
                'item_group_label': row.get('Item Group Label', '')
            }
    
    # Build summary and issues in one pass
    for item_name, data in all_comparisons.items():
        comp_df = data['comparison_df']
        
        match_counts = comp_df['Match'].value_counts()
        total_cols = len(comp_df)
        matches = match_counts.get('match', 0)
        mismatches = match_counts.get('mismatch', 0)
        missing_target = match_counts.get('missing_target', 0)
        missing_source = match_counts.get('missing_source', 0)
        match_percentage = (matches / total_cols * 100) if total_cols > 0 else 0
        
        # Add codelist info
        codelist_status = ''
        if data.get('codelist_comparison'):
            cl_comp = data['codelist_comparison']
            if cl_comp['status'] == 'match':
                codelist_status = f"✓ CL: {cl_comp['matches']} matches"
            else:
                codelist_status = f"✗ CL: {cl_comp['mismatches']} issues"
        
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
            'Codelist Status': codelist_status,
            'Match %': f"{match_percentage:.1f}%"
        })
        
        # Only process issues if not 100% match
        if match_percentage < 100.0:
            info = source_info_cache.get(item_name, {})
            
            # Filter non-matching rows
            issue_rows = comp_df[comp_df['Match'] != 'match']
            
            for _, row in issue_rows.iterrows():
                column_name = row['Column Name']
                source_value = row[f'{source_name} Value']
                target_value = row[f'{target_name} Value']
                
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
                    'Form Name': info.get('form_name', ''),
                    'Form Label': info.get('form_label', ''),
                    'Form Short Label': info.get('form_short_label', ''),
                    'Item Group Name': info.get('item_group_name', ''),
                    'Item Group Label': info.get('item_group_label', ''),
                    'Issue Type': issue_type
                })
    
    # Write to Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Comparison Summary', index=False)
        
        # Issues sheet
        if issues_data:
            issues_df = pd.DataFrame(issues_data)
            issues_df.to_excel(writer, sheet_name='Issues Only', index=False)
            
            workbook = writer.book
            worksheet = workbook['Issues Only']
            
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical='top', horizontal='left')
            
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 80)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        else:
            empty_df = pd.DataFrame({'Message': ['All items have 100% match rate.']})
            empty_df.to_excel(writer, sheet_name='Issues Only', index=False)
    
    output.seek(0)
    return output

@st.cache_data
def get_unique_items(df, column_name='Item Name'):
    """Cached function to get unique items"""
    return set(df[column_name].dropna().unique())

def main():
    st.title("📊 PTD vs SDS Comparison Tool")
    
    # Initialize session state
    if 'selected_items' not in st.session_state:
        st.session_state.selected_items = []
    if 'ptd_df' not in st.session_state:
        st.session_state.ptd_df = None
    if 'sds_df' not in st.session_state:
        st.session_state.sds_df = None
    if 'ptd_codelists_df' not in st.session_state:
        st.session_state.ptd_codelists_df = None
    if 'sds_codelists_df' not in st.session_state:
        st.session_state.sds_codelists_df = None
    if 'comparison_direction' not in st.session_state:
        st.session_state.comparison_direction = "PTD → SDS (Compare PTD columns against SDS)"
    if 'all_comparisons' not in st.session_state:
        st.session_state.all_comparisons = None
    if 'comparison_complete' not in st.session_state:
        st.session_state.comparison_complete = False
    if 'input_method' not in st.session_state:
        st.session_state.input_method = "📁 Upload Excel Files"
    if 'swap_triggered' not in st.session_state:
        st.session_state.swap_triggered = False
    
    # Input method selection
    st.markdown("---")
    st.subheader("📥 Select Input Method")
    input_method = st.radio(
        "How would you like to provide the data?",
        ["📁 Upload Excel Files"],
        horizontal=True
    )
    st.session_state.input_method = input_method
    
    # Comparison direction selection
    st.markdown("---")
    comparison_direction = st.radio(
        "🔄 Select Comparison Direction:",
        [
            "PTD → SDS (Compare PTD columns against SDS)",
            "SDS → PTD (Compare SDS columns against PTD)",
            "Parental SDS → Child SDS (Compare Parental SDS columns against Child SDS)",
            "Parental PTD → Child PTD (Compare Parental PTD columns against Child PTD)"
        ],
        index=["PTD → SDS (Compare PTD columns against SDS)",
               "SDS → PTD (Compare SDS columns against PTD)",
               "Parental SDS → Child SDS (Compare Parental SDS columns against Child SDS)",
               "Parental PTD → Child PTD (Compare Parental PTD columns against Child PTD)"].index(st.session_state.comparison_direction)
    )
    
    st.session_state.comparison_direction = comparison_direction
    
    # Determine labels
    if "Parental SDS → Child SDS" in comparison_direction:
        left_label, right_label = "Parental SDS", "Child SDS"
        left_key, right_key = "parental_sds", "child_sds"
        is_left_ptd, is_right_ptd = False, False
    elif "Parental PTD → Child PTD" in comparison_direction:
        left_label, right_label = "Parental PTD", "Child PTD"
        left_key, right_key = "parental_ptd", "child_ptd"
        is_left_ptd, is_right_ptd = True, True
    else:
        left_label, right_label = "PTD", "SDS"
        left_key, right_key = "ptd", "sds"
        is_left_ptd = (left_label == "PTD")
        is_right_ptd = (right_label == "PTD")
    
    st.markdown("---")
    
    # Upload files
    st.markdown("#### 📁 Upload Excel Files:")
    st.info("ℹ️ Data will be read from '**Form Definitions**' and '**Codelists**' sheets.")
    
    # Show warning based on file types
    if is_left_ptd and is_right_ptd:
        st.warning("⚠️ **For both PTD files:** First row skipped in Form Definitions, second row as headers. Data filtered for Y/Yes/Mod only.")
    elif is_left_ptd:
        st.warning(f"⚠️ **For {left_label} file:** First row skipped in Form Definitions, second row as headers. Data filtered for Y/Yes/Mod only.")
    elif is_right_ptd:
        st.warning(f"⚠️ **For {right_label} file:** First row skipped in Form Definitions, second row as headers. Data filtered for Y/Yes/Mod only.")
    
    st.info("ℹ️ **Automatic Processing:**\n"
            "- PTD data filtered for Y/Yes/Mod in 'Used in trial' column\n"
            "- 'Modification comments' and 'Used in trial' columns removed\n"
            "- 'Decimal' column formatted (e.g., 1 → 1.0)\n"
            "- Codelists filtered to remove blank Choice Code rows\n"
            "- Codelist comparison for matching 'Name', 'Choice Code', and 'Choice Label'")
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader(f"📄 {left_label} File")
        left_file = st.file_uploader(
            f"Upload {left_label} Excel file:",
            type=['xlsx', 'xls'],
            key=f"{left_key}_upload"
        )
        
        if left_file is not None:
            with st.spinner(f"Reading {left_label} file..."):
                # Read Form Definitions
                left_df = parse_uploaded_file(left_file, sheet_name='Form Definitions', is_ptd=is_left_ptd)
                if left_df is not None:
                    if is_left_ptd:
                        processed_result = process_ptd_dataframe(left_df)
                        if isinstance(processed_result, tuple):
                            left_df, original_count, filtered_count = processed_result
                            st.session_state.ptd_df = left_df
                            st.success(f"✅ {left_label} Form Definitions loaded: {len(left_df)} rows")
                            if original_count != filtered_count:
                                st.info(f"ℹ️ Filtered from {original_count} to {filtered_count} rows (Y/Yes/Mod only)")
                    else:
                        # Process SDS for decimal conversion
                        left_df = process_sds_dataframe(left_df)
                        st.session_state.ptd_df = left_df
                        st.success(f"✅ {left_label} Form Definitions loaded: {len(left_df)} rows")
                
                # Read Codelists
                left_codelists = parse_uploaded_file(left_file, sheet_name='Codelists', is_ptd=False)
                if left_codelists is not None:
                    processed_cl = process_codelists(left_codelists)
                    if isinstance(processed_cl, tuple):
                        left_codelists, cl_original, cl_filtered = processed_cl
                        st.session_state.ptd_codelists_df = left_codelists
                        st.success(f"✅ {left_label} Codelists loaded: {len(left_codelists)} rows")
                        if cl_original != cl_filtered:
                            st.info(f"ℹ️ Codelists filtered from {cl_original} to {cl_filtered} rows")
        elif st.session_state.ptd_df is not None:
            st.success(f"✅ {left_label} data loaded")
    
    with col2:
        st.subheader(f"📄 {right_label} File")
        right_file = st.file_uploader(
            f"Upload {right_label} Excel file:",
            type=['xlsx', 'xls'],
            key=f"{right_key}_upload"
        )
        
        if right_file is not None:
            with st.spinner(f"Reading {right_label} file..."):
                # Read Form Definitions
                right_df = parse_uploaded_file(right_file, sheet_name='Form Definitions', is_ptd=is_right_ptd)
                if right_df is not None:
                    if is_right_ptd:
                        processed_result = process_ptd_dataframe(right_df)
                        if isinstance(processed_result, tuple):
                            right_df, original_count, filtered_count = processed_result
                            st.session_state.sds_df = right_df
                            st.success(f"✅ {right_label} Form Definitions loaded: {len(right_df)} rows")
                            if original_count != filtered_count:
                                st.info(f"ℹ️ Filtered from {original_count} to {filtered_count} rows (Y/Yes/Mod only)")
                    else:
                        # Process SDS for decimal conversion
                        right_df = process_sds_dataframe(right_df)
                        st.session_state.sds_df = right_df
                        st.success(f"✅ {right_label} Form Definitions loaded: {len(right_df)} rows")
                
                # Read Codelists
                right_codelists = parse_uploaded_file(right_file, sheet_name='Codelists', is_ptd=False)
                if right_codelists is not None:
                    processed_cl = process_codelists(right_codelists)
                    if isinstance(processed_cl, tuple):
                        right_codelists, cl_original, cl_filtered = processed_cl
                        st.session_state.sds_codelists_df = right_codelists
                        st.success(f"✅ {right_label} Codelists loaded: {len(right_codelists)} rows")
                        if cl_original != cl_filtered:
                            st.info(f"ℹ️ Codelists filtered from {cl_original} to {cl_filtered} rows")
        elif st.session_state.sds_df is not None:
            st.success(f"✅ {right_label} data loaded")
    
    # Use data from session state
    ptd_df = st.session_state.ptd_df
    sds_df = st.session_state.sds_df
    ptd_codelists_df = st.session_state.ptd_codelists_df
    sds_codelists_df = st.session_state.sds_codelists_df
    
    has_left = ptd_df is not None
    has_right = sds_df is not None
    has_left_cl = ptd_codelists_df is not None
    has_right_cl = sds_codelists_df is not None
    
    if has_left and has_right:
        st.markdown("---")
        st.success("✅ Both datasets loaded! Ready to compare.")
        
        if has_left_cl and has_right_cl:
            st.success("✅ Codelists loaded for both files!")
        elif not has_left_cl and not has_right_cl:
            st.warning("⚠️ Codelists not loaded. Codelist comparison will be skipped.")
        else:
            st.warning("⚠️ Codelists loaded for only one file. Codelist comparison will show missing data.")
        
        # Set source and target
        if "PTD → SDS" in comparison_direction:
            source_df, target_df = ptd_df, sds_df
            source_name, target_name = "PTD", "SDS"
            source_codelists, target_codelists = ptd_codelists_df, sds_codelists_df
        elif "SDS → PTD" in comparison_direction:
            source_df, target_df = sds_df, ptd_df
            source_name, target_name = "SDS", "PTD"
            source_codelists, target_codelists = sds_codelists_df, ptd_codelists_df
        elif "Parental SDS → Child SDS" in comparison_direction:
            source_df, target_df = ptd_df, sds_df
            source_name, target_name = "Parental SDS", "Child SDS"
            source_codelists, target_codelists = ptd_codelists_df, sds_codelists_df
        else:  # Parental PTD → Child PTD
            source_df, target_df = ptd_df, sds_df
            source_name, target_name = "Parental PTD", "Child PTD"
            source_codelists, target_codelists = ptd_codelists_df, sds_codelists_df
        
        # Display info
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.info(f"📊 {source_name} Records: {len(source_df)}")
        with col2:
            st.info(f"📊 {target_name} Records: {len(target_df)}")
        with col3:
            IGNORE_COLUMNS = {'Definition Last Modified', 'Relationship Last Modified'}
            source_cols = [col for col in source_df.columns if col not in IGNORE_COLUMNS]
            st.info(f"📋 {source_name} Columns: {len(source_cols)}")
        with col4:
            if source_codelists is not None:
                st.info(f"📋 {source_name} Codelists: {len(source_codelists)} rows")
            else:
                st.warning(f"⚠️ No Codelists")
        
        st.markdown("---")
        
        # Get unique items
        source_items = get_unique_items(source_df)
        target_items = get_unique_items(target_df)
        all_items = sorted(list(source_items.union(target_items)))
        
        only_in_source = source_items - target_items
        only_in_target = target_items - source_items
        
        if only_in_source or only_in_target:
            with st.expander("⚠️ View Items Not in Both Files"):
                col1, col2 = st.columns(2)
                with col1:
                    if only_in_source:
                        st.warning(f"**Only in {source_name} ({len(only_in_source)}):**")
                        st.write(list(only_in_source))
                with col2:
                    if only_in_target:
                        st.warning(f"**Only in {target_name} ({len(only_in_target)}):**")
                        st.write(list(only_in_target))
        
        st.markdown("---")
        
        # Item selection
        st.subheader("Select Items to Compare")
        
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
        
        selected_items = st.multiselect(
            "Select Item Names:",
            options=all_items,
            default=st.session_state.selected_items
        )
        
        st.session_state.selected_items = selected_items
        
        if selected_items:
            st.info(f"📋 Selected {len(selected_items)} item(s)")
        else:
            st.warning("⚠️ No items selected")
        
        if selected_items and st.button("🔍 Compare Selected Items", type="primary", use_container_width=True):
            all_comparisons = {}
            
            # Build lookup dictionaries for faster access
            with st.spinner("Building lookup indices..."):
                source_dict = build_lookup_dictionaries(source_df)
                target_dict = build_lookup_dictionaries(target_df)
                source_columns = [col for col in source_df.columns if col not in {'Definition Last Modified', 'Relationship Last Modified'}]
            
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total_items = len(selected_items)
            start_time = time.time()
            
            for idx, item_name in enumerate(selected_items):
                if idx % 20 == 0 or idx == total_items - 1:
                    status_text.text(f"Comparing {idx+1}/{total_items}: {item_name}")
                    progress_bar.progress((idx + 1) / total_items)
                
                source_row, target_row, match_type = find_matching_rows_optimized(
                    source_df, target_df, item_name, source_dict, target_dict
                )
                
                if source_row is not None or target_row is not None:
                    # Check for Codelist column
                    codelist_comparison = None
                    codelist_value = None
                    
                    # Find Codelist column
                    codelist_col = None
                    for col in source_df.columns:
                        if col.strip().lower() == 'codelist':
                            codelist_col = col
                            break
                    
                    if codelist_col and source_row is not None:
                        codelist_value = source_row.get(codelist_col)
                        if pd.notna(codelist_value) and str(codelist_value).strip() != '':
                            # Compare codelists
                            codelist_comparison = compare_codelists(
                                codelist_value, 
                                source_codelists, 
                                target_codelists,
                                source_name,
                                target_name
                            )
                    
                    comparison_df = create_comparison_dataframe_fast(
                        source_row, target_row, source_columns, source_name, target_name,
                        codelist_comparison
                    )
                    all_comparisons[item_name] = {
                        'comparison_df': comparison_df,
                        'match_type': match_type,
                        'source_exists': source_row is not None,
                        'target_exists': target_row is not None,
                        'codelist_comparison': codelist_comparison
                    }
            
            elapsed_time = time.time() - start_time
            status_text.text(f"✅ Comparison complete! ({elapsed_time:.2f}s)")
            progress_bar.progress(1.0)
            
            st.session_state.all_comparisons = all_comparisons
            st.session_state.comparison_complete = True
            st.session_state.source_name = source_name
            st.session_state.target_name = target_name
            st.session_state.source_df = source_df
            
            st.success(f"✅ Completed comparison for {len(all_comparisons)} items in {elapsed_time:.2f}s!")
            
            st.markdown("---")
            
            # Summary
            st.subheader("📊 Comparison Summary")
            
            summary_data = []
            for item_name, data in all_comparisons.items():
                comp_df = data['comparison_df']
                match_counts = comp_df['Match'].value_counts()
                
                total_cols = len(comp_df)
                matches = match_counts.get('match', 0)
                match_percentage = (matches / total_cols * 100) if total_cols > 0 else 0
                
                # Codelist info
                cl_info = "N/A"
                if data.get('codelist_comparison'):
                    cl_comp = data['codelist_comparison']
                    if cl_comp['status'] == 'match':
                        cl_info = f"✓ {cl_comp['matches']} matches"
                    elif cl_comp['status'] in ['missing_source', 'missing_target', 'not_found']:
                        cl_info = cl_comp['message']
                    else:
                        cl_info = f"✗ {cl_comp['mismatches']} issues"
                
                summary_data.append({
                    'Item Name': item_name,
                    'Match Type': data['match_type'],
                    f'In {source_name}': '✓' if data['source_exists'] else '✗',
                    f'In {target_name}': '✓' if data['target_exists'] else '✗',
                    'Total Columns': total_cols,
                    'Matches ✅': matches,
                    'Mismatches ❌': match_counts.get('mismatch', 0),
                    f'Missing in {target_name} ⚠️': match_counts.get('missing_target', 0),
                    f'Missing in {source_name} ⚠️': match_counts.get('missing_source', 0),
                    'Codelist': cl_info,
                    'Match %': f"{match_percentage:.1f}%"
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, height=300)
            
            # Filter for display
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
            
            st.subheader("📋 Detailed Results")
            
            if items_with_100_match:
                st.success(f"✅ {len(items_with_100_match)} item(s) with 100% match hidden")
                with st.expander(f"View {len(items_with_100_match)} items with 100% match"):
                    st.write(items_with_100_match)
            
            if not items_to_display:
                st.info("🎉 All items have 100% match!")
            else:
                st.info(f"Showing {len(items_to_display)} item(s) with discrepancies")
                
                if len(items_to_display) > 20:
                    st.warning(f"⚠️ Displaying first 20 items. Download report for all {len(items_to_display)} items.")
                    items_to_show = dict(list(items_to_display.items())[:20])
                else:
                    items_to_show = items_to_display
                
                tabs = st.tabs([item_name for item_name in items_to_show.keys()])
                
                for tab, (item_name, data) in zip(tabs, items_to_show.items()):
                    with tab:
                        comparison_df = data['comparison_df']
                        match_counts = comparison_df['Match'].value_counts()
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            if data['match_type']:
                                st.success(f"✅ Matched by: **{data['match_type']}**")
                            st.info(f"In {source_name}: {'✓' if data['source_exists'] else '✗'}")
                            
                            # Codelist info
                            if data.get('codelist_comparison'):
                                cl_comp = data['codelist_comparison']
                                if cl_comp['status'] == 'match':
                                    st.success(f"✓ Codelist: {cl_comp['matches']} matches")
                                else:
                                    st.warning(f"⚠️ Codelist: {cl_comp['message']}")
                        
                        with col2:
                            matches = match_counts.get('match', 0)
                            match_rate = (matches/len(comparison_df)*100)
                            st.metric("Match Rate", f"{match_rate:.1f}%")
                            st.info(f"In {target_name}: {'✓' if data['target_exists'] else '✗'}")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Rows", len(comparison_df))
                        with col2:
                            st.metric("✅ Matches", match_counts.get('match', 0))
                        with col3:
                            st.metric("❌ Mismatches", match_counts.get('mismatch', 0))
                        with col4:
                            missing = match_counts.get('missing_source', 0) + match_counts.get('missing_target', 0)
                            st.metric("⚠️ Missing", missing)
                        
                        st.markdown("---")
                        st.markdown("##### Complete Comparison Table")
                        styled_df = comparison_df.style.apply(highlight_differences, axis=1)
                        st.dataframe(styled_df, use_container_width=True, height=500)
            
            st.markdown("---")
            st.markdown("**Legend:**")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown("🟢 **Green**: Values match")
            with col2:
                st.markdown("🔴 **Red**: Values differ")
            with col3:
                st.markdown("🟡 **Orange**: Value missing")
            with col4:
                st.markdown("📋 **Codelist rows**: Show Choice Code comparisons")
        
        # Download report
        if st.session_state.comparison_complete and st.session_state.all_comparisons:
            st.markdown("---")
            st.markdown("---")
            st.subheader("📥 Download Report")
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.info("📊 Report includes:\n"
                       "- **Sheet 1**: Comparison Summary (with Codelist status)\n"
                       "- **Sheet 2**: Issues Only (detailed)")
            
            with col2:
                with st.spinner("Generating report..."):
                    report_output = create_comprehensive_report_fast(
                        st.session_state.all_comparisons,
                        st.session_state.source_name,
                        st.session_state.target_name,
                        st.session_state.source_df
                    )
                
                st.download_button(
                    label="📥 Download Report",
                    data=report_output,
                    file_name=f"comparison_{st.session_state.source_name}_vs_{st.session_state.target_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
    
    elif not has_left and not has_right:
        st.warning(f"⚠️ Please provide both {left_label} and {right_label} data.")
    elif not has_left:
        st.warning(f"⚠️ Please provide {left_label} data.")
    else:
        st.warning(f"⚠️ Please provide {right_label} data.")

if __name__ == "__main__":
    main()
