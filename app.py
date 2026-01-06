import pandas as pd
import streamlit as st
from io import BytesIO, StringIO
import openpyxl

def parse_pasted_data(pasted_text):
    """Parse pasted Excel data (tab-separated) into DataFrame"""
    try:
        # Try to parse as tab-separated (Excel default when copying)
        df = pd.read_csv(StringIO(pasted_text), sep='\t', engine='python')
        return df
    except Exception as e:
        st.error(f"Error parsing pasted data: {e}")
        return None

def find_matching_rows(source_df, target_df, item_name):
    """
    Find matching rows based on Item Name, Item Group Label, or Form Label
    """
    # First try to match by Item Name
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

def compare_values(val1, val2):
    """Compare two values and return status"""
    # Handle NaN values
    if pd.isna(val1) and pd.isna(val2):
        return 'match', '✓', ''
    if pd.isna(val1) and not pd.isna(val2):
        return 'missing_source', '⚠', 'Missing in Source'
    if not pd.isna(val1) and pd.isna(val2):
        return 'missing_target', '⚠', 'Missing in Target'
    
    # Convert to string for comparison
    str1, str2 = str(val1).strip(), str(val2).strip()
    
    if str1 == str2:
        return 'match', '✓', ''
    else:
        return 'mismatch', '✗', 'Values differ'

def create_comparison_dataframe(source_row, target_row, source_df, source_name, target_name):
    """Create a comparison dataframe for display - only source columns"""
    comparison_data = []
    
    # Columns to ignore from comparison (specifically for SDS)
    IGNORE_COLUMNS = ['Definition Last Modified', 'Relationship Last Modified']
    
    # Get only source columns
    source_columns = list(source_df.columns)
    
    for col in source_columns:
        # Skip ignored columns if they're in the source
        if col in IGNORE_COLUMNS:
            continue
            
        # Get source value
        source_value = source_row[col].values[0] if not source_row.empty else None
        
        # Get target value if column exists in target and is not in ignore list
        target_value = None
        if not target_row.empty and col in target_row.columns:
            target_value = target_row[col].values[0]
        
        status, symbol, note = compare_values(source_value, target_value)
        
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

def export_comparison_to_excel(all_comparisons, source_name, target_name):
    """Export all comparisons to Excel with multiple sheets"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Create summary sheet
        summary_data = []
        for item_name, comparison_df in all_comparisons.items():
            summary_data.append({
                'Item Name': item_name,
                'Total Columns': len(comparison_df),
                'Matches': len(comparison_df[comparison_df['Match'] == 'match']),
                'Mismatches': len(comparison_df[comparison_df['Match'] == 'mismatch']),
                f'Missing in {source_name}': len(comparison_df[comparison_df['Match'] == 'missing_source']),
                f'Missing in {target_name}': len(comparison_df[comparison_df['Match'] == 'missing_target'])
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Create individual comparison sheets
        for item_name, comparison_df in all_comparisons.items():
            # Clean sheet name (Excel has limitations)
            sheet_name = item_name[:31] if len(item_name) > 31 else item_name
            comparison_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="PTD vs SDS Comparison Tool", layout="wide")
    
    st.title("📊 PTD vs SDS Comparison Tool")
    #st.markdown("### Copy-Paste Data Comparison")
    #st.markdown("---")
    
    # Initialize session state for selected items
    if 'selected_items' not in st.session_state:
        st.session_state.selected_items = []
    if 'ptd_df' not in st.session_state:
        st.session_state.ptd_df = None
    if 'sds_df' not in st.session_state:
        st.session_state.sds_df = None
    
    # Instructions
    #st.info("💡 **Instructions:** Copy your data from Excel (including headers) and paste it in the text areas below. Make sure to copy the entire table with column headers.")
    
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
    
    ptd_df = None
    sds_df = None
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 PTD Data")
        st.caption("Copy data from Excel and paste here")
        ptd_text = st.text_area(
            "Paste PTD data here:",
            height=300,
            placeholder="Select and copy your PTD data from Excel, then paste here...",
            key="ptd_paste",
            help="Select all PTD data in Excel (including headers), copy (Ctrl+C), and paste here (Ctrl+V)"
        )
        
        if ptd_text:
            with st.spinner("Parsing PTD data..."):
                ptd_df = parse_pasted_data(ptd_text)
                if ptd_df is not None:
                    st.session_state.ptd_df = ptd_df
                    st.success(f"✅ PTD data loaded: {len(ptd_df)} rows, {len(ptd_df.columns)} columns")
        elif st.session_state.ptd_df is not None:
            st.success(f"✅ PTD data loaded: {len(st.session_state.ptd_df)} rows, {len(st.session_state.ptd_df.columns)} columns")
    
    with col2:
        st.subheader("📄 SDS Data")
        st.caption("Copy data from Excel and paste here")
        sds_text = st.text_area(
            "Paste SDS data here:",
            height=300,
            placeholder="Select and copy your SDS data from Excel, then paste here...",
            key="sds_paste",
            help="Select all SDS data in Excel (including headers), copy (Ctrl+C), and paste here (Ctrl+V)"
        )
        
        if sds_text:
            with st.spinner("Parsing SDS data..."):
                sds_df = parse_pasted_data(sds_text)
                if sds_df is not None:
                    st.session_state.sds_df = sds_df
                    st.success(f"✅ SDS data loaded: {len(sds_df)} rows, {len(sds_df.columns)} columns")
        elif st.session_state.sds_df is not None:
            st.success(f"✅ SDS data loaded: {len(st.session_state.sds_df)} rows, {len(st.session_state.sds_df.columns)} columns")
    
    # Use data from session state
    ptd_df = st.session_state.ptd_df
    sds_df = st.session_state.sds_df
    
    if ptd_df is not None and sds_df is not None:
        st.markdown("---")
        st.success("✅ Both datasets loaded successfully! Ready to compare.")
        
        # Show ignored columns info
        #st.info("ℹ️ The following columns are excluded from comparison: **Definition Last Modified**, **Relationship Last Modified**")
        
        # Comparison direction selection
        st.markdown("---")
        comparison_direction = st.radio(
            "🔄 Select Comparison Direction:",
            ["PTD → SDS (Compare PTD columns against SDS)", 
             "SDS → PTD (Compare SDS columns against PTD)"],
            help="Choose which file's columns to use as the basis for comparison"
        )
        
        # Set source and target based on selection
        if "PTD → SDS" in comparison_direction:
            source_df = ptd_df
            target_df = sds_df
            source_name = "PTD"
            target_name = "SDS"
        else:
            source_df = sds_df
            target_df = ptd_df
            source_name = "SDS"
            target_name = "PTD"
        
        # Display basic info
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
        
        # Get all unique Item Names from both dataframes
        source_items = set(source_df['Item Name'].dropna().unique())
        target_items = set(target_df['Item Name'].dropna().unique())
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
        
        # Comparison mode selection
        comparison_mode = st.radio(
            "Select Comparison Mode:",
            ["🔍 Multi-Select Item Comparison", "📊 Bulk Comparison (All Items)"],
            horizontal=True
        )
        
        st.markdown("---")
        
        if comparison_mode == "🔍 Multi-Select Item Comparison":
            # Multi-select item comparison
            st.subheader("Select Items to Compare")
            
            # Quick selection options
            col1, col2, col3 = st.columns(3)
            with col1:
                select_all = st.button("✅ Select All", use_container_width=True)
                if select_all:
                    st.session_state.selected_items = all_items
                    st.rerun()
            
            with col2:
                clear_selection = st.button("❌ Clear Selection", use_container_width=True)
                if clear_selection:
                    st.session_state.selected_items = []
                    st.rerun()
            
            with col3:
                select_source = st.button(f"🔍 Select from {source_name} only", use_container_width=True)
                if select_source:
                    st.session_state.selected_items = sorted(list(source_items))
                    st.rerun()
            
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
                
                for idx, item_name in enumerate(selected_items):
                    status_text.text(f"Comparing {idx+1}/{len(selected_items)}: {item_name}")
                    
                    source_row, target_row, match_type = find_matching_rows(source_df, target_df, item_name)
                    
                    if source_row.empty and target_row.empty:
                        st.warning(f"⚠️ No matching records found for {item_name}")
                    else:
                        comparison_df = create_comparison_dataframe(source_row, target_row, source_df, source_name, target_name)
                        all_comparisons[item_name] = {
                            'comparison_df': comparison_df,
                            'match_type': match_type,
                            'source_exists': not source_row.empty,
                            'target_exists': not target_row.empty
                        }
                    
                    progress_bar.progress((idx + 1) / len(selected_items))
                
                status_text.text("✅ Comparison complete!")
                
                st.success(f"✅ Completed comparison for {len(all_comparisons)} items!")
                
                st.markdown("---")
                
                # Display summary for selected items
                st.subheader("📊 Comparison Summary")
                
                summary_data = []
                for item_name, data in all_comparisons.items():
                    comp_df = data['comparison_df']
                    total_cols = len(comp_df)
                    matches = len(comp_df[comp_df['Match'] == 'match'])
                    mismatches = len(comp_df[comp_df['Match'] == 'mismatch'])
                    missing_target = len(comp_df[comp_df['Match'] == 'missing_target'])
                    missing_source = len(comp_df[comp_df['Match'] == 'missing_source'])
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
                
                #st.markdown("---")
                
                # Display individual comparisons
                st.subheader("📋 Detailed Comparison Results")
                
                # Tab selection for each item
                if len(all_comparisons) > 0:
                    tabs = st.tabs([item_name for item_name in all_comparisons.keys()])
                    
                    for tab, (item_name, data) in zip(tabs, all_comparisons.items()):
                        with tab:
                            comparison_df = data['comparison_df']
                            
                            # Show match information
                            col1, col2 = st.columns(2)
                            with col1:
                                if data['match_type']:
                                    st.success(f"✅ Matched by: **{data['match_type']}**")
                                st.info(f"In {source_name}: {'✓' if data['source_exists'] else '✗'}")
                            with col2:
                                matches = len(comparison_df[comparison_df['Match'] == 'match'])
                                st.metric("Match Rate", f"{(matches/len(comparison_df)*100):.1f}%")
                                st.info(f"In {target_name}: {'✓' if data['target_exists'] else '✗'}")
                            
                            # Summary counts for this item
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                total_rows = len(comparison_df)
                                st.metric("Total Rows", total_rows)
                            with col2:
                                match_count = len(comparison_df[comparison_df['Match'] == 'match'])
                                st.metric("✅ Matches", match_count)
                            with col3:
                                mismatch_count = len(comparison_df[comparison_df['Match'] == 'mismatch'])
                                st.metric("❌ Mismatches", mismatch_count)
                            with col4:
                                missing_count = len(comparison_df[
                                    (comparison_df['Match'] == 'missing_source') | 
                                    (comparison_df['Match'] == 'missing_target')
                                ])
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
                
                # Export selected comparisons
                st.markdown("---")
                export_comparisons = {item_name: data['comparison_df'] 
                                     for item_name, data in all_comparisons.items()}
                excel_output = export_comparison_to_excel(export_comparisons, source_name, target_name)
                
                st.download_button(
                    label=f"📥 Download Selected Comparisons ({len(all_comparisons)} items)",
                    data=excel_output,
                    file_name=f"selected_comparison_{source_name}_vs_{target_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        else:
            # Bulk comparison (All items)
            st.info(f"📊 Total items to compare: **{len(all_items)}**")
            
            col1, col2 = st.columns([3, 1])
            with col1:
                start_comparison = st.button("🚀 Start Bulk Comparison", 
                                             type="primary", 
                                             use_container_width=True)
            
            if start_comparison:
                all_comparisons = {}
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, item_name in enumerate(all_items):
                    status_text.text(f"Comparing {idx+1}/{len(all_items)}: {item_name}")
                    
                    source_row, target_row, match_type = find_matching_rows(source_df, target_df, item_name)
                    
                    if not source_row.empty:  # Only compare if item exists in source
                        comparison_df = create_comparison_dataframe(source_row, target_row, source_df, source_name, target_name)
                        all_comparisons[item_name] = comparison_df
                    
                    progress_bar.progress((idx + 1) / len(all_items))
                
                status_text.text("✅ Comparison complete!")
                
                st.success(f"✅ Completed comparison for **{len(all_comparisons)}** items!")
                
                st.markdown("---")
                
                # Display summary
                st.subheader("📊 Bulk Comparison Summary")
                
                summary_data = []
                for item_name, comp_df in all_comparisons.items():
                    total_cols = len(comp_df)
                    matches = len(comp_df[comp_df['Match'] == 'match'])
                    mismatches = len(comp_df[comp_df['Match'] == 'mismatch'])
                    missing_target = len(comp_df[comp_df['Match'] == 'missing_target'])
                    missing_source = len(comp_df[comp_df['Match'] == 'missing_source'])
                    
                    match_percentage = (matches / total_cols * 100) if total_cols > 0 else 0
                    
                    summary_data.append({
                        'Item Name': item_name,
                        'Total Columns': total_cols,
                        'Matches ✅': matches,
                        'Mismatches ❌': mismatches,
                        f'Missing in {target_name} ⚠️': missing_target,
                        f'Missing in {source_name} ⚠️': missing_source,
                        'Match %': f"{match_percentage:.1f}%"
                    })
                
                summary_df = pd.DataFrame(summary_data)
                
                # Sort and filter options
                col1, col2 = st.columns(2)
                with col1:
                    sort_option = st.selectbox(
                        "Sort by:",
                        ["Item Name", "Total Columns", "Matches ✅", "Mismatches ❌", "Match %"],
                        key="sort_bulk"
                    )
                with col2:
                    filter_bulk = st.selectbox(
                        "Filter by:",
                        ["Show All", "Only Items with Mismatches", "Only Items with Missing Values", "100% Match Only"],
                        key="filter_bulk"
                    )
                
                # Apply bulk filter
                filtered_summary = summary_df.copy()
                if filter_bulk == "Only Items with Mismatches":
                    filtered_summary = summary_df[summary_df['Mismatches ❌'] > 0].copy()
                elif filter_bulk == "Only Items with Missing Values":
                    filtered_summary = summary_df[
                        (summary_df[f'Missing in {target_name} ⚠️'] > 0) | 
                        (summary_df[f'Missing in {source_name} ⚠️'] > 0)
                    ].copy()
                elif filter_bulk == "100% Match Only":
                    filtered_summary = summary_df[summary_df['Match %'] == '100.0%'].copy()
                
                # Apply sort
                if sort_option == "Match %":
                    filtered_summary['Match_Numeric'] = filtered_summary['Match %'].str.rstrip('%').astype(float)
                    filtered_summary = filtered_summary.sort_values('Match_Numeric', ascending=False)
                    filtered_summary = filtered_summary.drop('Match_Numeric', axis=1)
                else:
                    filtered_summary = filtered_summary.sort_values(sort_option, ascending=False)
                
                st.info(f"Showing {len(filtered_summary)} of {len(summary_df)} items")
                st.dataframe(filtered_summary, use_container_width=True, height=400)
                
                # Export all comparisons
                st.markdown("---")
                excel_output = export_comparison_to_excel(all_comparisons, source_name, target_name)
                
                st.download_button(
                    label="📥 Download All Comparisons (Excel with Summary)",
                    data=excel_output,
                    file_name=f"bulk_comparison_{source_name}_vs_{target_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success("💡 The Excel file contains a Summary sheet and individual sheets for each item comparison.")
    
    elif ptd_df is None and sds_df is None:
        st.warning("⚠️ Please paste both PTD and SDS data to start comparing.")
    elif ptd_df is None:
        st.warning("⚠️ Please paste PTD data.")
    else:
        st.warning("⚠️ Please paste SDS data.")

if __name__ == "__main__":
    main()