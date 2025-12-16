import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
from io import BytesIO

# --- CONFIGURATION & SETUP ---
st.set_page_config(page_title="Universal Excel Filler", layout="wide")

st.title("🤖 AI Auto-Fill for Excel")
st.markdown("""
**Workflow:**
1. Upload your **Template Excel Sheet** (so I know what columns you need).
2. Upload **Images** or Paste **Text**.
3. Review the data and **Download**.
""")

# --- SIDEBAR: API KEY ---
with st.sidebar:
    st.header("🔑 Settings")
    # Tries to get key from secrets first (for hosted version), else asks user
    if 'GOOGLE_API_KEY' in st.secrets:
        api_key = st.secrets['GOOGLE_API_KEY']
        st.success("API Key loaded securely!")
    else:
        api_key = st.text_input("Enter Google Gemini API Key", type="password")
        st.markdown("[Get Key Here](https://aistudio.google.com/)")

# --- SESSION STATE (To remember data between clicks) ---
if 'final_data' not in st.session_state:
    st.session_state.final_data = pd.DataFrame()
if 'template_columns' not in st.session_state:
    st.session_state.template_columns = []

# --- AI FUNCTION ---
def extract_data_with_gemini(content, mime_type, columns):
    """
    Sends image/text + column names to AI. 
    Returns a dictionary matching the columns.
    """
    if not api_key:
        st.error("Missing API Key!")
        return None

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')

    # Dynamic Prompt: We tell AI EXACTLY what columns to look for
    column_list_str = ", ".join(columns)
    
    prompt = f"""
    You are a strict data extraction bot.
    
    Task: Extract information from the provided input to fill a database row.
    
    TARGET COLUMNS: [{column_list_str}]
    
    RULES:
    1. Only extract data that matches the meaning of the Target Columns.
    2. If a column is "COMPANY", look for store names or organizations (e.g., "Moksh", "Ninjacart").
    3. If a column is "PRICE" or "AMOUNT", look for the total value.
    4. If data for a column is NOT found in the input, return null (empty).
    5. Return ONLY a valid JSON object. Keys must match the TARGET COLUMNS exactly.
    6. If the input contains a table with multiple items, return a LIST of JSON objects (one for each row).
    
    Example Output format:
    [
        {{"{columns[0]}": "Value1", "{columns[1]}": "Value2", ...}},
        {{"{columns[0]}": "Value3", "{columns[1]}": "Value4", ...}}
    ]
    """

    try:
        # Generate content
        response = model.generate_content([prompt, content])
        
        # Clean response (remove markdown if present)
        text_res = response.text.strip()
        if text_res.startswith("```json"):
            text_res = text_res[7:-3]
        
        return json.loads(text_res)
    except Exception as e:
        st.error(f"AI Error: {e}")
        return []

# --- STEP 1: UPLOAD EXCEL TEMPLATE ---
st.subheader("Step 1: Upload Template Excel")
template_file = st.file_uploader("Upload the Excel file you want to fill", type=['xlsx', 'xls'])

if template_file:
    # Read headers
    try:
        df_template = pd.read_excel(template_file)
        st.session_state.template_columns = df_template.columns.tolist()
        st.success(f"✅ Columns Detected: {st.session_state.template_columns}")
        
        # Initialize session dataframe with these columns if empty
        if st.session_state.final_data.empty:
            st.session_state.final_data = pd.DataFrame(columns=st.session_state.template_columns)
            
    except Exception as e:
        st.error("Error reading Excel file. Make sure it's valid.")

# --- STEP 2: DATA INPUT ---
if st.session_state.template_columns:
    st.markdown("---")
    st.subheader("Step 2: Upload Data Sources")
    
    col_input, col_preview = st.columns([1, 2])
    
    with col_input:
        tab1, tab2 = st.tabs(["📸 Images", "📝 Text"])
        
        # --- IMAGE INPUT ---
        with tab1:
            img_files = st.file_uploader("Upload Images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
            if st.button("Process Images"):
                if not img_files:
                    st.warning("No images selected.")
                else:
                    progress = st.progress(0)
                    new_rows = []
                    
                    for idx, img_file in enumerate(img_files):
                        # Convert to format for Gemini
                        from PIL import Image
                        image_data = Image.open(img_file)
                        
                        extracted = extract_data_with_gemini(image_data, "image/jpeg", st.session_state.template_columns)
                        
                        if extracted:
                            # If list (multiple rows), extend. If dict (single row), append.
                            if isinstance(extracted, list):
                                new_rows.extend(extracted)
                            else:
                                new_rows.append(extracted)
                        
                        progress.progress((idx + 1) / len(img_files))
                    
                    # Add to master data
                    if new_rows:
                        new_df = pd.DataFrame(new_rows)
                        st.session_state.final_data = pd.concat([st.session_state.final_data, new_df], ignore_index=True)
                        st.success("Images Processed!")

        # --- TEXT INPUT ---
        with tab2:
            txt_input = st.text_area("Paste Text (WhatsApp/SMS)")
            if st.button("Process Text"):
                if not txt_input:
                    st.warning("No text entered.")
                else:
                    extracted = extract_data_with_gemini(txt_input, "text/plain", st.session_state.template_columns)
                    if extracted:
                        if isinstance(extracted, list):
                            new_df = pd.DataFrame(extracted)
                        else:
                            new_df = pd.DataFrame([extracted])
                            
                        st.session_state.final_data = pd.concat([st.session_state.final_data, new_df], ignore_index=True)
                        st.success("Text Processed!")

    # --- STEP 3: PREVIEW & EDIT ---
    with col_preview:
        st.subheader("Step 3: Preview & Edit")
        st.info("Double-click any cell to fix mistakes before downloading.")
        
        # EDITABLE DATA FRAME
        edited_df = st.data_editor(
            st.session_state.final_data,
            num_rows="dynamic",
            use_container_width=True,
            key="editor"
        )
        
        # Sync changes back to session state
        st.session_state.final_data = edited_df

        st.markdown("---")
        
        # --- STEP 4: EXPORT ---
        # Logic to preserve user's template headers
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
        st.download_button(
            label="⬇️ Download Final Excel",
            data=output.getvalue(),
            file_name="Filled_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        if st.button("🗑️ Clear All Rows"):
            st.session_state.final_data = pd.DataFrame(columns=st.session_state.template_columns)
            st.rerun()

else:
    st.info("👆 Please upload your Excel Template in Step 1 to begin.")

    
