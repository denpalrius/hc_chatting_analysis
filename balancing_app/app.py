import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

# Import helper modules
from schedule_balancer import ScheduleBalancer
from data_validator import DataValidator
from excel_formatter import ExcelFormatter

def main():
    st.set_page_config(
        page_title="HumaneCare Schedule Balancer",
        page_icon="🏥",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    
    st.title("🏥 HumaneCare Schedule Balancer")
    st.markdown("---")
    
    # Instructions
    with st.expander("📋 Instructions"):
        st.markdown("""
        **How to use this application:**
        
        1. **Upload your Excel file** containing healthcare provider schedules
        2. **Click Process** to automatically balance the schedules
        3. **Download** the processed file with color-coded changes
        
        **The system will automatically:**
        - Balance each provider to maximum 16 hours per day (minimum 2 hours)
        - Ensure each individual totals exactly 24 hours per day
        - Add supplemental providers when needed
        - Apply color coding to show all changes made
        - Handle exception cases when standard balancing fails
        
        **Color Coding:**
        - 🔴 **Red highlight**: Days that cannot be fully balanced
        - 🟢 **Green highlight**: New additions or entries changed from 0 to positive
        - 🟠 **Orange highlight**: Reduced or modified non-zero entries
        - 🟢 **Green font**: Names of newly added providers
        """)
    
    # File upload section
    st.subheader("📁 Upload Schedule File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload your healthcare provider schedule Excel file"
    )
    
    # Additional providers configuration
    st.subheader("👥 Additional Provider Configuration")
    with st.expander("Configure Additional Providers (Optional)"):
        st.write("**Default Supplemental Providers:**")
        st.write("- Charles Sagini, RN/House Manager")
        st.write("- Josephine Sagini, RN/Program Manager")
        st.write("- Faith Murerwa, RN/House Supervisor")
        
        st.write("**Additional Emergency Providers:**")
        additional_providers_text = st.text_area(
            "Enter additional provider names (one per line):",
            value="Carolyn Porter, LPN",
            help="These providers will be used as a last resort for balancing"
        )
        
        additional_providers = []
        if additional_providers_text.strip():
            additional_providers = [line.strip() for line in additional_providers_text.strip().split('\n') if line.strip()]
    
    if uploaded_file is not None:
        try:
            # Display file info
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            
            # Process button
            if st.button("🔄 Process Schedule"):
                with st.spinner("Processing schedule... This may take a few moments."):
                    # Process the file
                    processed_file, summary = process_schedule_file(uploaded_file)
                    
                    if processed_file is not None:
                        st.success("✅ Schedule processing completed!")
                        
                        # Display summary (optional feature)
                        if summary:
                            with st.expander("📊 Processing Summary"):
                                display_summary(summary)
                        
                        # Download button
                        st.subheader("💾 Download Processed File")
                        st.download_button(
                            label="📥 Download Balanced Schedule",
                            data=processed_file,
                            file_name=f"balanced_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    else:
                        st.error("❌ Failed to process the file. Please check the file format and try again.")
            
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.info("Please ensure your Excel file follows the expected format with provider schedules.")

def process_schedule_file(uploaded_file):
    """
    Main processing function that coordinates all balancing operations
    """
    try:
        # Save uploaded file temporarily for openpyxl processing
        temp_file_path = f"/tmp/{uploaded_file.name}"
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Validate data structure using pandas first
        df = pd.read_excel(temp_file_path, sheet_name=None, engine='openpyxl')
        validator = DataValidator()
        
        if not validator.validate_file_structure(df):
            st.error("❌ Invalid file structure. Please check your Excel file format.")
            return None, None
        
        # Extract day blocks with formatting using openpyxl
        day_blocks, workbook = validator.extract_day_blocks_with_formatting(temp_file_path)
        
        if not day_blocks:
            st.error("❌ No valid day blocks found in the file.")
            return None, None
        
        # Validate day blocks
        issues = validator.validate_day_blocks(day_blocks)
        if issues:
            st.warning(f"⚠️ Found {len(issues)} validation issues. Processing will continue.")
            for issue in issues[:5]:  # Show first 5 issues
                st.warning(f"- {issue}")
        
        # Initialize the schedule balancer with dynamic individuals and additional providers
        individuals = validator.expected_individuals
        additional_providers = []
        # Get additional providers from session state if available
        if 'additional_providers' in locals():
            additional_providers = locals()['additional_providers']
        
        balancer = ScheduleBalancer(individuals=individuals, additional_providers=additional_providers)
        
        # Balance the schedule
        processed_workbook, summary = balancer.balance_schedule(day_blocks, workbook)
        
        if not processed_workbook:
            st.error("❌ Failed to process the schedule.")
            return None, None
        
        # Create formatted Excel file
        formatter = ExcelFormatter()
        processed_file = formatter.create_formatted_excel(processed_workbook, balancer.get_changes_log())
        
        # Clean up temp file
        import os
        try:
            os.remove(temp_file_path)
        except:
            pass
        
        return processed_file, summary
        
    except Exception as e:
        st.error(f"Error in process_schedule_file: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None, None

def display_summary(summary):
    """
    Display processing summary statistics
    """
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Days Processed", summary['total_days_processed'])
    
    with col2:
        st.metric("Days Balanced", summary['days_balanced'])
    
    with col3:
        st.metric("Days Unbalanced", summary['days_unbalanced'])
    
    with col4:
        st.metric("Entries Modified", summary['entries_modified'])

if __name__ == "__main__":
    main()
