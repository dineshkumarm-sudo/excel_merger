%%writefile consolidator_app.py
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Universal Excel Consolidator", layout="wide")
st.title("🗂️ Universal Multi-Excel Schema Mapper & Consolidator")
st.write("Upload multiple Excel sheets, declare your custom target columns, and compile them into a unified master file.")

# 1. Configuration Sidebar
st.sidebar.header("⚙️ Consolidation Settings")

# Completely generic, generic default configuration fields
default_headers = "First Name, Last Name, Email, Phone Number, Status"
headers_input = st.sidebar.text_area(
    "Target Headers (Comma Separated):", 
    value=default_headers,
    help="Type the exact column headers you want to extract and combine into your final sheet."
)

# Parse inputs into a clean configuration list
target_headers = [h.strip() for h in headers_input.split(",") if h.strip()]

# 2. File Upload Zone
uploaded_files = st.file_uploader(
    "Upload Your Excel Files (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📚 Loaded {len(uploaded_files)} files into memory buffer. Ready to process.")
    
    if st.button("🚀 Run Consolidation & Schema Audit"):
        if not target_headers:
            st.error("Please provide at least one target header configuration.")
            st.stop()
            
        master_rows = []
        missing_headers_log = {} 
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, file_obj in enumerate(uploaded_files):
            file_name = file_obj.name
            status_text.text(f"Scanning schema for: {file_name}...")
            
            try:
                # Read only header metadata first to optimize processing speeds
                df_header = pd.read_excel(file_obj, nrows=0)
                file_columns = df_header.columns.tolist()
                
                # Case-insensitive header matching map
                column_map = {str(col).lower().strip(): col for col in file_columns}
                
                missing_in_this_file = []
                matched_file_columns = {}
                
                # Check file headers against user configuration inputs
                for target in target_headers:
                    target_lower = target.lower().strip()
                    if target_lower in column_map:
                        matched_file_columns[target] = column_map[target_lower]
                    else:
                        missing_in_this_file.append(target)
                        
                # Log files that have structural mismatch errors
                if missing_in_this_file:
                    missing_headers_log[file_name] = missing_in_this_file
                    
                # Load content only if matching header elements exist
                columns_to_load = list(matched_file_columns.values())
                
                if columns_to_load:
                    file_obj.seek(0)
                    df_full = pd.read_excel(file_obj, usecols=columns_to_load)
                    
                    # Dynamically convert rows to dictionaries mapping back to target configs
                    for _, row in df_full.iterrows():
                        compiled_row = {"Source File Name": file_name}
                        
                        for target, actual_col in matched_file_columns.items():
                            compiled_row[target] = row[actual_col]
                            
                        # Missing columns in specific files safely populate as blank (None)
                        for target in missing_in_this_file:
                            compiled_row[target] = None
                            
                        master_rows.append(compiled_row)
                        
            except Exception as e:
                st.error(f"❌ Critical parse error on file '{file_name}': {str(e)}")
                
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        status_text.text("🎉 System processing and consolidation complete!")
        
        # 3. Missing Schema Report Panel
        st.subheader("⚠️ Missing Schema Audit Report")
        if missing_headers_log:
            st.warning("The following files were missing one or more of your target headers:")
            
            audit_records = []
            for f_name, missing_list in missing_headers_log.items():
                audit_records.append({
                    "File Name": f_name,
                    "Missing Headers Count": len(missing_list),
                    "Specific Columns Missing": ", ".join(missing_list)
                })
            st.dataframe(pd.DataFrame(audit_records), use_container_width=True)
        else:
            st.success("✅ Clean Sweep! All uploaded files contain 100% of your target headers.")
            
        # 4. Master Output Table Synthesis
        if master_rows:
            master_df = pd.DataFrame(master_rows)
            
            # Reorder columns with the source log file name tracking column up front
            ordered_cols = ["Source File Name"] + target_headers
            master_df = master_df[ordered_cols]
            
            st.subheader(f"📊 Consolidated Master Preview ({len(master_df)} Total Rows Combined)")
            st.dataframe(master_df.head(100), use_container_width=True)
            
            # Excel export script wrapper
            out_buffer = io.BytesIO()
            with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
                master_df.to_excel(writer, index=False, sheet_name="Master Consolidated")
            out_buffer.seek(0)
            
            st.download_button(
                label="📥 Download Consolidated Master File",
                data=out_buffer.getvalue(),
                file_name="Master_Consolidated_Data.xlsx",
                mime="application/vnd.ms-excel"
            )
        else:
            st.error("No row records matched your target configurations across any file.")