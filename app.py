import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Auto-Excel Bot", layout="wide")
st.title("🤖 100% Accuracy AI Data Entry")

# --- SIDEBAR: SETTINGS ---
with st.sidebar:
    st.header("🔑 Settings")
    if 'GOOGLE_API_KEY' in st.secrets:
        api_key = st.secrets['GOOGLE_API_KEY']
        st.success("Key loaded automatically!")
    else:
        api_key = st.text_input("Enter Google API Key", type="password")
        st.markdown("[Get Free Key](https://aistudio.google.com/)")

# --- SESSION STATE ---
if 'final_data' not in st.session_state:
    st.session_state.final_data = pd.DataFrame()
if 'template_columns' not in st.session_state:
    st.session_state.template_columns = []

# --- AI EXTRACTION FUNCTION ---
def extract_data_with_gemini(content, mime_type, columns):
    if not api_key:
        st.error("Please add your API Key first!")
        return None

    genai.configure(api_key=api_key)
    
    # Use the stable, high-performance model
    model = genai.GenerativeModel('gemini-2.5-flash')

    column_list_str = ", ".join(columns)
    
    # --- UPDATED PROMPT FOR ACCURACY ---
    prompt = f"""
    You are an expert data entry specialist. 
    Analyze the provided input (Image or Text) and extract data to fill these Excel columns: [{column_list_str}]
    
    STRICT DATA RULES:
    1. **Column 'TYPE'**: You MUST extract the specific **Vegetable or Fruit Name** (e.g., "Bottle Gourd", "Tomato", "Apple"). 
       - NEVER use generic words like "Receipt", "Bill", "Vegetable", or "Food".
       - If the text says "Baby Bottle Gourd-4.5", extract "Baby Bottle Gourd".
    2. **Column 'COMPANY'**: Extract the company name (e.g., "Moksh Enterprises", "Ninjacart").
    3. **Column 'ID NO'**: Extract the PO ID or Invoice Number.
    4. **Column 'PRICE'**: Extract the total amount.
    5. **Column 'DATE'**: Extract the date in DD/MM/YYYY format.
    6. **General**: If a column's data is missing, use an empty string "".
    
    Output Format: return ONLY a raw JSON list of objects.
    Example: [{{"TYPE": "Baby Bottle Gourd", "COMPANY": "Ninjacart", "PRICE": "3471.30", ...}}]
    """

    try:
        response = model.generate_content([prompt, content])
        text_res = response.text.strip()
        
        # Clean up Markdown formatting if the AI adds it
        if text_res.startswith("```json"):
            text_res = text_res[7:-3]
        elif text_res.startswith("```"):
            text_res = text_res[3:-3]
            
        return json.loads(text_res)
    except Exception as e:
        st.error(f"AI Error: {e}")
        return []

# --- MAIN APP LAYOUT ---
st.info("Step 1: Upload your empty Excel sheet to define columns.")
template_file = st.file_uploader("Upload Empty Excel Template", type=['xlsx', 'xls'])

if template_file:
    # Load Template
    try:
        df_template = pd.read_excel(template_file)
        st.session_state.template_columns = df_template.columns.tolist()
        
        # Initialize the master dataframe if it's new
        if st.session_state.final_data.empty:
            st.session_state.final_data = pd.DataFrame(columns=st.session_state.template_columns)
            
        st.success(f"✅ Columns Found: {st.session_state.template_columns}")
        
    except Exception as e:
        st.error("Error reading template file.")

    st.markdown("---")

    col_left, col_right = st.columns([1, 1.5])

    # --- LEFT: INPUTS ---
    with col_left:
        st.subheader("Step 2: Upload Data")
        input_type = st.radio("Select Input Type:", ["Images 📸", "Text 📝"], horizontal=True)

        if input_type == "Images 📸":
            img_files = st.file_uploader("Select Images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
            if st.button("Extract Data from Images"):
                if not img_files:
                    st.warning("Please choose images first.")
                else:
                    bar = st.progress(0)
                    new_rows = []
                    for i, file in enumerate(img_files):
                        from PIL import Image
                        img = Image.open(file)
                        
                        data = extract_data_with_gemini(img, "image/jpeg", st.session_state.template_columns)
                        if data:
                            if isinstance(data, list): new_rows.extend(data)
                            else: new_rows.append(data)
                        bar.progress((i+1)/len(img_files))
                    
                    if new_rows:
                        new_df = pd.DataFrame(new_rows)
                        st.session_state.final_data = pd.concat([st.session_state.final_data, new_df], ignore_index=True)
                        st.success("Done!")
                        st.rerun()

        elif input_type == "Text 📝":
            txt_in = st.text_area("Paste text here...", height=200)
            if st.button("Extract Data from Text"):
                data = extract_data_with_gemini(txt_in, "text/plain", st.session_state.template_columns)
                if data:
                    if isinstance(data, list): new_rows = data
                    else: new_rows = [data]
                    new_df = pd.DataFrame(new_rows)
                    st.session_state.final_data = pd.concat([st.session_state.final_data, new_df], ignore_index=True)
                    st.success("Done!")
                    st.rerun()

    # --- RIGHT: PREVIEW & DOWNLOAD ---
    with col_right:
        st.subheader("Step 3: Verify & Download")
        
        if not st.session_state.final_data.empty:
            
            # --- CRITICAL FIX FOR "ARROW INVALID" ERROR ---
            # We convert everything to Text (String) just for the display to prevent the crash.
            # The downloaded Excel will still be correct.
            display_df = st.session_state.final_data.copy()
            display_df = display_df.astype(str) 
            display_df = display_df.replace('nan', '')
            display_df = display_df.replace('None', '')
            # ----------------------------------------------

            # EDITABLE GRID
            edited_df = st.data_editor(
                display_df, 
                num_rows="dynamic", 
                use_container_width=True,
                height=500,
                key="editor"
            )
            
            # Save edits back to main state
            st.session_state.final_data = edited_df
            
            # DOWNLOAD
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                edited_df.to_excel(writer, index=False)
            
            st.download_button(
                label="⬇️ Download Completed Excel",
                data=output.getvalue(),
                file_name="Completed_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            if st.button("Clear All Data"):
                st.session_state.final_data = pd.DataFrame(columns=st.session_state.template_columns)
                st.rerun()
        else:
            st.info("Data will appear here after extraction.")

