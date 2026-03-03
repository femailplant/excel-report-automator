import streamlit as st
import pandas as pd
import time
from excel_mapper import process_mapping_execution, create_zip_archive, generate_mock_mapping_file

# Must be the first Streamlit command
st.set_page_config(page_title="Excel Report Automator", layout="wide", initial_sidebar_state="collapsed")

# Inject Custom CSS for the advanced UI look
st.markdown("""
<style>
    /* Global Background and Text */
    .stApp {
        background-color: #1a1b26;
        color: #c0caf5;
    }
    
    /* Headers */
    h1, h2, h3, h4, h5, h6, .st-emotion-cache-10trblm {
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    
    /* Step Indicator Container */
    .step-indicator {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 2rem 0;
        gap: 15px;
    }
    .step-line {
        height: 2px;
        width: 150px;
        background: linear-gradient(90deg, #ff7b54, #ffb26b);
        border-radius: 2px;
    }
    .step-line.inactive {
        background: #3b4252;
    }
    .step-circle {
        width: 24px;
        height: 24px;
        border-radius: 50%;
        background-color: #ff7b54;
        display: flex;
        justify-content: center;
        align-items: center;
        box-shadow: 0 0 10px rgba(255, 123, 84, 0.5);
    }
    .step-circle.inactive {
        background-color: #3b4252;
        box-shadow: none;
    }
    
    /* Custom Columns/Cards Styling */
    [data-testid="column"] {
        background-color: rgba(30, 32, 48, 0.7);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(4px);
    }
    
    /* File Uploader override */
    [data-testid="stFileUploadDropzone"] {
        background-color: rgba(255, 255, 255, 0.03);
        border: 1px dashed rgba(255, 255, 255, 0.2);
        border-radius: 8px;
    }

    /* Primary Generate Button Styling */
    button[kind="primary"] {
        background: linear-gradient(90deg, #ff7b54, #ff5722);
        border: none;
        box-shadow: 0 0 20px rgba(255, 87, 34, 0.6);
        color: white;
        font-weight: bold;
        padding: 10px 30px;
        border-radius: 8px;
        transition: all 0.3s ease;
        display: block;
        margin: 0 auto;
        width: 50% !important;
    }
    button[kind="primary"]:hover {
        background: linear-gradient(90deg, #ff5722, #e64a19);
        box-shadow: 0 0 30px rgba(255, 87, 34, 0.8);
    }
    
    /* Log Window Expander styling */
    [data-testid="stExpander"] {
        background-color: #16161e;
        border: 1px solid rgba(255,255,255,0.05);
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Main Title Area
st.title("📊 Excel Report Automator")
st.markdown("Extract values from multiple files and merge them into your designed templates easily in one click.")

# Sidebar Helper
st.sidebar.header("🛠️ Tools & Settings")
st.sidebar.download_button(
    "📥 Download Mapping Config Example",
    data=generate_mock_mapping_file(),
    file_name="Mapping_Template_Example.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.write("") # Spacer

# Three Main Columns (Cards)
col1, col2, col3 = st.columns(3, gap="large")

with col1:
    st.markdown("### 1. Mapping Configuration")
    st.markdown("<span style='font-size:0.85rem; color:#8c92ac;'>Upload Mapping File (Excel/CSV)</span>", unsafe_allow_html=True)
    mapping_file = st.file_uploader("Mapping Config limits 200MB - XLSX, CSV", type=["xlsx", "csv"], label_visibility="collapsed")

with col2:
    st.markdown("### 2. Source Data Files")
    st.markdown("<span style='font-size:0.85rem; color:#8c92ac;'>Upload Source Excel files</span>", unsafe_allow_html=True)
    source_files = st.file_uploader("Source Data limits 700MB - XLSX", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")

with col3:
    st.markdown("### 3. Template Files")
    st.markdown("<span style='font-size:0.85rem; color:#8c92ac;'>Upload Template Excel files</span>", unsafe_allow_html=True)
    template_files = st.file_uploader("Templates limits 200MB - XLSX", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")

# UI Step Indicator
st.markdown("""
<div class="step-indicator">
    <div style="text-align:center;">
        <div class="step-circle" style="margin:0 auto; width:12px; height:12px;"></div>
        <span style="font-size:0.8rem; color:#ff7b54; margin-top:5px; display:block;">Mapping</span>
    </div>
    <div class="step-line"></div>
    <div style="text-align:center;">
        <div class="step-circle" style="margin:0 auto; width:12px; height:12px;"></div>
        <span style="font-size:0.8rem; color:#ff7b54; margin-top:5px; display:block;">Complete</span>
    </div>
    <div class="step-line inactive"></div>
    <div style="text-align:center;">
        <div class="step-circle inactive" style="margin:0 auto; width:12px; height:12px;"></div>
        <span style="font-size:0.8rem; color:#8c92ac; margin-top:5px; display:block;">Template</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.write("")

# Action Area
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    generate_clicked = st.button("🚀 Generate Reports", type="primary", use_container_width=True)

if generate_clicked:
    if not mapping_file or not source_files or not template_files:
        st.error("⚠️ Please upload all required files in the 3 columns before processing.")
    else:
        # Use st.status to simulate the terminal logic window at the bottom
        with st.status("Initializing process...", expanded=True) as status:
            try:
                st.write("✅ Validation Successful. Mapping, Sources, and Templates verified.")
                time.sleep(0.5)
                
                # Load mapping
                st.write("🔄 Parsing configuration mapping file...")
                if mapping_file.name.endswith(".csv"):
                    mapping_df = pd.read_csv(mapping_file)
                else:
                    mapping_df = pd.read_excel(mapping_file)

                # Process files into dict format required by backend
                st.write("🔄 Loading source data into memory...")
                source_dict = {f.name: f.read() for f in source_files}
                
                st.write("🔄 Loading templates into memory...")
                template_dict = {f.name: f.read() for f in template_files}

                # Core execution
                st.write("🚀 Processing Data Injection...")
                generated_reports = process_mapping_execution(mapping_df, source_dict, template_dict)
                
                if not generated_reports:
                    status.update(label="Process finished with no results", state="error", expanded=True)
                    st.warning("Processed successfully, but no output was generated. Please check if your mapping file definitions perfectly match the uploaded file names.")
                else:
                    st.write(f"✅ Reports Created. ({len(generated_reports)} files matching signatures)")
                    status.update(label="Complete!", state="complete", expanded=False)
                    
                    st.success(f"🎉 Successfully generated {len(generated_reports)} report(s)!")
                    
                    col_dl1, col_dl2, col_dl3 = st.columns([1,2,1])
                    with col_dl2:
                        if len(generated_reports) == 1:
                            # Single file download
                            filename, file_bytes = list(generated_reports.items())[0]
                            st.download_button(
                                label=f"💾 Download {filename}",
                                data=file_bytes,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            # Zip download for multiple reports
                            zip_bytes = create_zip_archive(generated_reports)
                            st.download_button(
                                label="🗂️ Download All as ZIP archive",
                                data=zip_bytes,
                                file_name="Generated_Reports.zip",
                                mime="application/zip",
                                use_container_width=True
                            )

            except Exception as e:
                status.update(label="Failed to process", state="error", expanded=True)
                st.error(f"❌ An error occurred during processing: {str(e)}")
                st.info("Check your mapping configuration to ensure strict matching with file/sheet/cell names.")
