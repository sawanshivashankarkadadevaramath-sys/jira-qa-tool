import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from io import BytesIO

st.set_page_config(page_title="Jira to QA Test Matrix Generator", page_icon="🧪", layout="wide")

st.title("🧪 Jira to QA Test Matrix Generator")
st.markdown("Upload one or multiple Jira `.doc` / HTML export files to instantly generate a standardized MVA QA Test Case matrix.")

uploaded_files = st.file_uploader("Upload Jira Export Files (.doc / .html)", type=["doc", "html"], accept_multiple_files=True)

col1, col2 = st.columns(2)
with col1:
    assigned_to = st.text_input("Assigned To (QA Lead/Tester Name)", value="Sawan")
with col2:
    environment = st.text_input("Environment", value="QA1")

def parse_jira_html(html_content, filename):
    key_match = re.search(r'\[([A-Z0-9]+-\d+)\]', html_content)
    req_id = key_match.group(1) if key_match else filename.split('.')[0]
    
    status_match = re.search(r'<b>Status:</b></td>\s*<td[^>]*>(.*?)</td>', html_content, re.DOTALL)
    status = re.sub(r'<[^>]+>', '', status_match.group(1)).strip() if status_match else "New"
    
    prio_match = re.search(r'<b>Priority:</b>\s*</td>\s*<td[^>]*>\s*(.*?)\s*</td>', html_content, re.DOTALL)
    prio = re.sub(r'<[^>]+>', '', prio_match.group(1)).strip() if prio_match else "Medium"
    
    desc_match = re.search(r'id="descriptionArea">(.*?)</td>', html_content, re.DOTALL)
    if desc_match:
        lines = [re.sub(r'<[^>]+>', '', l).strip() for l in desc_match.group(1).split('<br') if re.sub(r'<[^>]+>', '', l).strip()]
        description = lines[0] if lines else ""
    else:
        description = ""
        
    ac_match = re.search(r'<b>Acceptance Criteria:</b></td>\s*<td[^>]*>(.*?)</td>', html_content, re.DOTALL)
    criteria = []
    if ac_match:
        items = re.findall(r'<li>(.*?)</li>', ac_match.group(1), re.DOTALL)
        for item in items:
            clean_item = re.sub(r'<[^>]+>', '', item).strip()
            if clean_item and "google" not in clean_item.lower():
                criteria.append(clean_item)
                
    if not criteria:
        criteria.append(f"Verify functionality and acceptance criteria for {req_id}")
        
    return {
        "req_id": req_id,
        "description": description,
        "status": status,
        "priority": prio,
        "criteria": criteria
    }

if uploaded_files and st.button("🚀 Generate Test Matrix", type="primary"):
    all_data = []
    headers = [
        'REQ ID \n(User Story)', 'NAME', 'DESCRIPTION', 'Status', 
        'STEP NUMBER', 'STEP ACTION', 'STEP EXPECTED RESULTS', 
        'APPLICATION', 'TEST CASE TYPE', 'PRIORITY', 'Assigned To', 'Environment', 'User Story '
    ]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QA_Test_Matrix"
    ws.append(headers)
    
    header_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True)
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC')
    )
    
    for col_idx in range(1, 14):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        
    row_idx = 2
    for f in uploaded_files:
        html = f.read().decode('utf-8', errors='ignore')
        parsed = parse_jira_html(html, f.name)
        
        req_id = parsed["req_id"]
        desc = parsed["description"]
        status = parsed["status"]
        prio = parsed["priority"]
        
        for idx, item in enumerate(parsed["criteria"], start=1):
            tc_seq = f"{idx:03d}"
            tc_name = f"TC{tc_seq}_{req_id}_Functional Verification"
            action = f"Verify {item}"
            expected = item
            
            row = [
                req_id, tc_name, desc, status, idx, action, expected,
                "MVA", "Manual", prio, assigned_to, environment, req_id
            ]
            ws.append(row)
            all_data.append(row)
            
            for col_idx in range(1, 14):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Calibri", size=11)
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = thin_border
            row_idx += 1
            
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    st.success(f"Successfully generated {len(all_data)} test cases across {len(uploaded_files)} stories!")
    st.download_button(
        label="📥 Download Formatted Excel Matrix (.xlsx)",
        data=output,
        file_name="Jira_QA_Test_Matrix_MVA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
