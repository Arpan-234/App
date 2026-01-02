import streamlit as st
import pandas as pd
import io
import os
from pathlib import Path
import anthropic
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import json
from datetime import datetime

# Configure page
st.set_page_config(
    page_title="AI Excel & PowerPoint Agent",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-title {
        font-size: 2.5rem;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize Claude client
@st.cache_resource
def get_claude_client():
    api_key = st.secrets.get("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        st.error("ANTHROPIC_API_KEY not found")
        st.stop()
    return anthropic.Anthropic(api_key=api_key)

try:
    client = get_claude_client()
except:
    st.error("Failed to initialize Claude client")
    st.stop()

# Sidebar
st.sidebar.title("⚙️ Configuration")
st.sidebar.markdown("---")

mode = st.sidebar.radio(
    "Select Mode:",
    ["📈 Excel Automation", "📊 PowerPoint Creation", "🔄 Excel → PowerPoint"]
)

api_model = st.sidebar.selectbox(
    "AI Model:",
    ["claude-3-5-sonnet-20241022", "claude-3-opus-20250219"]
)

# Helper functions
def call_claude(prompt: str, system_prompt: str = None) -> str:
    try:
        message = client.messages.create(
            model=api_model,
            max_tokens=4096,
            system=system_prompt or "You are a helpful assistant",
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text
    except Exception as e:
        return f"Error: {str(e)}"

def analyze_excel_with_ai(df: pd.DataFrame, task: str) -> dict:
    df_summary = f"Rows: {df.shape[0]}, Cols: {df.shape[1]}\n{df.head().to_string()}"
    prompt = f"Analyze this data for {task}:\n{df_summary}\nReturn JSON format."
    response = call_claude(prompt)
    try:
        return json.loads(response)
    except:
        return {"result": response}

def create_ppt_from_data(data_dict: dict, title: str, theme_color: tuple) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Title slide
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme_color)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    # Content slides
    for section_title, content in data_dict.items():
        if section_title == "title":
            continue
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = section_title
        title.text_frame.paragraphs[0].font.size = Pt(44)
        
        body = slide.placeholders[1].text_frame
        body.clear()
        
        if isinstance(content, list):
            for item in content:
                p = body.add_paragraph()
                p.text = str(item)[:100]
                p.font.size = Pt(18)
        elif isinstance(content, dict):
            for k, v in content.items():
                p = body.add_paragraph()
                p.text = f"{k}: {v}"
                p.font.size = Pt(16)
    
    return prs

# Main UI
st.markdown("<h1 class='main-title'>📊 AI Excel & PowerPoint Agent</h1>", unsafe_allow_html=True)
st.markdown("Powered by Claude & Streamlit")
st.markdown("---")

if mode == "📈 Excel Automation":
    st.header("Excel Automation")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx", "xls", "csv"])
    task = st.selectbox("Task", ["clean", "analyze", "summarize", "visualize"])
    
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"Loaded: {df.shape[0]} rows, {df.shape[1]} cols")
            st.dataframe(df.head())
            
            if st.button("Analyze"):
                with st.spinner("Analyzing..."):
                    results = analyze_excel_with_ai(df, task)
                st.json(results)
        except Exception as e:
            st.error(f"Error: {str(e)}")

elif mode == "📊 PowerPoint Creation":
    st.header("Create PowerPoint")
    uploaded_file = st.file_uploader("Upload Excel for PPT", type=["xlsx", "xls", "csv"])
    title = st.text_input("Title", "AI Report")
    color = st.selectbox("Theme", ["Blue", "Green", "Red"])
    colors = {"Blue": (31, 119, 180), "Green": (44, 160, 44), "Red": (214, 39, 40)}
    
    if uploaded_file and st.button("Generate PPT"):
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            
            ppt_content = {
                "Overview": ["Generated presentation", "Data analysis"],
                "Summary": {"Rows": df.shape[0], "Columns": df.shape[1]}
            }
            
            prs = create_ppt_from_data(ppt_content, title, colors[color])
            ppt_buffer = io.BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)
            
            st.download_button(
                "Download PPT",
                ppt_buffer,
                f"{title}.pptx",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"Error: {str(e)}")

else:
    st.header("Excel → PowerPoint Workflow")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx", "xls", "csv"])
    title = st.text_input("Presentation Title", "AI Report")
    
    if uploaded_file and st.button("Create Full Workflow"):
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"Processing: {df.shape[0]} rows")
            
            ppt_content = {
                "Data Summary": [f"Total Records: {df.shape[0]}", f"Fields: {df.shape[1]}"],
                "Columns": list(df.columns)[:5]
            }
            
            prs = create_ppt_from_data(ppt_content, title, (31, 119, 180))
            ppt_buffer = io.BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)
            
            st.download_button(
                "Download PPT",
                ppt_buffer,
                f"{title}.pptx",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"Error: {str(e)}")

st.sidebar.markdown("---")
st.sidebar.info("Built with Streamlit & Claude AI")
