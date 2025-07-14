import streamlit as st
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import os
from PIL import Image
import base64
from streamlit_drawable_canvas import st_canvas
import numpy as np
from utils import (add_header_with_logo, add_footer, style_heading, 
                  create_info_table, add_logo_to_doc, set_cell_margins)
from equipment_inspection import equipment_inspector

# Page configuration
st.set_page_config(
    page_title="Halton KSA Service Reports",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f4788;
        font-weight: bold;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #2c5aa0;
        font-weight: bold;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
    .stButton > button {
        background-color: #1f4788;
        color: white;
        font-weight: bold;
        padding: 0.5rem 2rem;
        border-radius: 5px;
        border: none;
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #2c5aa0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'report_data' not in st.session_state:
    st.session_state.report_data = {}
if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False
if 'technician_signature' not in st.session_state:
    st.session_state.technician_signature = None


def create_technical_report(data):
    """Generate a Professional Technical Report Word document"""
    doc = Document()
    
    # Set document margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.header_distance = Inches(0.5)
        section.footer_distance = Inches(0.5)
    
    # Check for logo and add professional header
    logo_path = add_logo_to_doc(doc)
    add_header_with_logo(doc, logo_path)
    
    # Add footer
    add_footer(doc)
    
    # Add some space after header
    doc.add_paragraph()
    
    # Add title with styling
    title = doc.add_heading('TECHNICAL REPORT', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_heading(title, level=0)
    
    # Add report reference and date
    ref_para = doc.add_paragraph()
    ref_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ref_run = ref_para.add_run(f"Report Date: {data.get('date', datetime.now().strftime('%B %d, %Y'))}")
    ref_run.font.size = Pt(11)
    ref_run.font.color.rgb = RGBColor(100, 100, 100)
    
    # Add a subtle line separator
    doc.add_paragraph('_' * 85)
    
    # GENERAL INFORMATION SECTION
    general_heading = doc.add_heading('1. GENERAL INFORMATION', level=1)
    style_heading(general_heading, level=1)
    
    # Create professional info table
    general_info = [
        ("Customer Name", data.get('customer_name', '')),
        ("Project Name", data.get('project_name', '')),
        ("Contact Person", data.get('contact_person', '')),
        ("Location", data.get('outlet_location', '')),
        ("Contact Number", data.get('contact_number', '')),
        ("Visit Type", data.get('visit_type', '')),
        ("Visit Classification", data.get('visit_class', ''))
    ]
    
    create_info_table(doc, general_info)
    doc.add_paragraph()  # Add spacing
    
    # EQUIPMENT INSPECTION SECTION
    equipment_heading = doc.add_heading('2. EQUIPMENT INSPECTION DETAILS', level=1)
    style_heading(equipment_heading, level=1)
    
    equipment_summary = data.get('equipment_inspection', [])
    
    if equipment_summary:
        for idx, equip in enumerate(equipment_summary):
            # Equipment header
            equip_title = doc.add_heading(f"{equip['type_name']} - {equip['serial_number']}", level=2)
            style_heading(equip_title, level=2)
            
            # Equipment info
            equip_info_para = doc.add_paragraph()
            equip_info_para.add_run(f"Location: {equip['location']}\n").font.size = Pt(11)
            equip_info_para.add_run(f"Photos Taken: {equip['photos_count']}\n").font.size = Pt(11)
            
            # Issues found
            if equip['issues_found']:
                issues_para = doc.add_paragraph()
                issues_para.add_run("Issues Identified:\n").bold = True
                for issue in equip['issues_found']:
                    issue_text = f"• {issue['item'].replace('_', ' ').title()}: {issue['answer']}"
                    if issue['comment']:
                        issue_text += f" - {issue['comment']}"
                    issues_para.add_run(issue_text + "\n").font.size = Pt(10)
            else:
                doc.add_paragraph("No issues identified during inspection.").font.size = Pt(11)
            
            # Add spacing between equipment
            if idx < len(equipment_summary) - 1:
                doc.add_paragraph()
    else:
        doc.add_paragraph("No equipment inspection data available.")
    
    # WORK PERFORMED SECTION
    work_heading = doc.add_heading('3. WORK PERFORMED', level=1)
    style_heading(work_heading, level=1)
    
    work_para = doc.add_paragraph()
    work_text = work_para.add_run(data.get('work_performed', ''))
    work_text.font.size = Pt(11)
    work_para.paragraph_format.line_spacing = 1.5
    work_para.paragraph_format.space_after = Pt(12)
    
    # FINDINGS AND OBSERVATIONS SECTION
    findings_heading = doc.add_heading('4. FINDINGS AND OBSERVATIONS', level=1)
    style_heading(findings_heading, level=1)
    
    findings_para = doc.add_paragraph()
    findings_text = findings_para.add_run(data.get('findings', ''))
    findings_text.font.size = Pt(11)
    findings_para.paragraph_format.line_spacing = 1.5
    findings_para.paragraph_format.space_after = Pt(12)
    
    # RECOMMENDATIONS SECTION
    if data.get('recommendations'):
        rec_heading = doc.add_heading('5. RECOMMENDATIONS', level=1)
        style_heading(rec_heading, level=1)
        
        rec_para = doc.add_paragraph()
        rec_text = rec_para.add_run(data.get('recommendations', ''))
        rec_text.font.size = Pt(11)
        rec_para.paragraph_format.line_spacing = 1.5
        rec_para.paragraph_format.space_after = Pt(12)
    
    # SERVICE INFORMATION SECTION
    service_heading = doc.add_heading('6. SERVICE INFORMATION', level=1)
    style_heading(service_heading, level=1)
    
    service_info = [
        ("Technician Name", data.get('technician_name', '')),
        ("Technician ID", data.get('technician_id', '')),
        ("Service Date", data.get('service_date', datetime.now().strftime('%B %d, %Y')))
    ]
    
    create_info_table(doc, service_info, col_widths=[2.5, 4])
    
    # Add page break before signatures
    doc.add_page_break()
    
    # SIGNATURE SECTION
    sig_heading = doc.add_heading('ACKNOWLEDGMENT AND SIGNATURES', level=1)
    style_heading(sig_heading, level=1)
    
    # Add acknowledgment text
    ack_para = doc.add_paragraph()
    ack_text = ack_para.add_run(
        "The undersigned acknowledge that the service described in this report has been "
        "completed satisfactorily and in accordance with the agreed specifications."
    )
    ack_text.font.size = Pt(10)
    ack_text.font.italic = True
    ack_para.paragraph_format.space_after = Pt(24)
    
    # Create signature table
    sig_table = doc.add_table(rows=4, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Configure signature columns
    for col in sig_table.columns:
        for cell in col.cells:
            cell.width = Inches(3)
    
    # Technician signature
    sig_table.cell(0, 0).text = "Service Technician:"
    
    # Add signature image if available
    if data.get('technician_signature'):
        sig_cell = sig_table.cell(1, 0)
        sig_para = sig_cell.paragraphs[0]
        sig_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = sig_para.add_run()
        run.add_picture(data.get('technician_signature'), width=Inches(1.5))
    else:
        sig_table.cell(1, 0).text = "_" * 35
    
    sig_table.cell(2, 0).text = data.get('technician_name', '')
    sig_table.cell(3, 0).text = f"Date: {datetime.now().strftime('%B %d, %Y')}"
    
    # Customer signature
    sig_table.cell(0, 1).text = "Customer Representative:"
    sig_table.cell(1, 1).text = "_" * 35
    sig_table.cell(2, 1).text = "Name: _________________________"
    sig_table.cell(3, 1).text = "Date: _________________________"
    
    # Style signature table
    for row in sig_table.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(11)
            set_cell_margins(cell, top=0.1, bottom=0.1)
    
    # Make labels bold
    sig_table.cell(0, 0).paragraphs[0].runs[0].font.bold = True
    sig_table.cell(0, 1).paragraphs[0].runs[0].font.bold = True
    
    # Add final note
    doc.add_paragraph()
    note_para = doc.add_paragraph()
    note_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note_text = note_para.add_run(
        "This report is confidential and proprietary to Halton Foodservice KSA.\n"
        "For service inquiries, please contact our Service Department."
    )
    note_text.font.size = Pt(9)
    note_text.font.color.rgb = RGBColor(128, 128, 128)
    
    # Save to bytes
    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    return doc_bytes

def main():
    # Header
    st.markdown('<h1 class="main-header">Halton KSA Service Reports</h1>', unsafe_allow_html=True)
    
    # Sidebar for report type selection
    with st.sidebar:
        st.markdown("### Report Type Selection")
        report_type = st.selectbox(
            "Select Report Type",
            ["Technical Report", "Testing and Commissioning Report (Coming Soon)", "General Service Report (Coming Soon)"],
            index=0
        )
        
        if report_type != "Technical Report":
            st.info("This report type will be available in future versions.")
            st.stop()
    
    # Main form for Technical Report
    st.markdown('<h2 class="section-header">Technical Report Form</h2>', unsafe_allow_html=True)
    
    with st.form("technical_report_form"):
        # General Information Section
        st.markdown("### General Information")
        col1, col2 = st.columns(2)
        
        with col1:
            customer_name = st.text_input("Customer's Name*", placeholder="e.g., SELA Company")
            project_name = st.text_input("Project Name*", placeholder="e.g., Stella kitchen hoods")
            contact_person = st.text_input("Contact Person*", placeholder="e.g., Sultan Alofi")
            outlet_location = st.text_input("Outlet/Location*", placeholder="e.g., Via - Riyadh")
        
        with col2:
            contact_number = st.text_input("Contact #*", placeholder="e.g., +966 55 558 5449")
            visit_type = st.selectbox(
                "Visit Type*",
                ["", "Servicing/Preventive Maintenance", "AMC (Contract)", "Emergency Service", "Installation", "Commissioning"]
            )
            # Store visit type in session state for equipment inspection
            st.session_state.visit_type = visit_type
            
            visit_class = st.selectbox(
                "Visit Class*",
                ["To be invoiced (Chargeable)", "Warranty", "Complaint", "Scheduled Maintenance", "Emergency Service"]
            )
            date = st.date_input("Report Date", value=datetime.now())
        
        # Equipment Inspection Section
        equipment_inspector.render_equipment_section()
        
        # Work Performed Section
        st.markdown("### Work Performed")
        work_performed = st.text_area(
            "Describe Work Performed*",
            placeholder="Detail all maintenance, repairs, or services completed...",
            height=150
        )
        
        # Findings Section
        st.markdown("### Findings and Observations")
        findings = st.text_area(
            "Findings and Observations*",
            placeholder="Document any issues found, equipment condition, etc...",
            height=150
        )
        
        # Recommendations Section
        st.markdown("### Recommendations")
        recommendations = st.text_area(
            "Recommendations",
            placeholder="Suggest any follow-up actions, parts needed, or future maintenance...",
            height=100
        )
        
        # Technician Information Section
        st.markdown("### Technician Information")
        col3, col4 = st.columns(2)
        
        with col3:
            technician_name = st.text_input("Technician Name*", placeholder="Enter your full name")
            technician_id = st.text_input("Technician ID*", placeholder="Enter your employee ID")
        
        with col4:
            service_date = st.date_input("Service Date", value=datetime.now())
        
        # Signature Section
        st.markdown("### Technician Signature")
        st.markdown("Please draw your signature below using your mouse or touchscreen")
        
        # Create signature canvas
        col1, col2 = st.columns([4, 1])
        
        with col1:
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)",  # Transparent fill
                stroke_width=3,
                stroke_color="#000000",
                background_color="#FFFFFF",
                background_image=None,
                update_streamlit=True,
                height=150,
                width=500,
                drawing_mode="freedraw",
                point_display_radius=0,
                display_toolbar=True,
                key="signature_canvas",
            )
        
        with col2:
            st.markdown("### ")  # Add spacing
            st.info("Use the trash icon in the canvas toolbar to clear")
        
        # Check if signature is drawn
        if canvas_result.image_data is not None:
            # Check if canvas has any drawing (non-transparent pixels)
            if np.any(canvas_result.image_data[:,:,3] > 0):
                st.success("✅ Signature captured")
        
        # Submit button
        submitted = st.form_submit_button("Generate Report", type="primary")
        
        if submitted:
            # Validation
            required_fields = {
                "Customer's Name": customer_name,
                "Project Name": project_name,
                "Contact Person": contact_person,
                "Outlet/Location": outlet_location,
                "Contact Number": contact_number,
                "Visit Type": visit_type,
                "Work Performed": work_performed,
                "Findings": findings,
                "Technician Name": technician_name,
                "Technician ID": technician_id
            }
            
            missing_fields = [field for field, value in required_fields.items() if not value]
            
            # Validate equipment data
            equipment_errors = equipment_inspector.validate_equipment_data()
            
            if missing_fields:
                st.error(f"Please fill in the following required fields: {', '.join(missing_fields)}")
            elif not st.session_state.equipment_list:
                st.error("Please add at least one equipment for inspection")
            elif equipment_errors:
                st.error("Equipment inspection errors:\n" + "\n".join(f"• {error}" for error in equipment_errors))
            else:
                # Process signature from canvas
                signature_img = None
                if canvas_result.image_data is not None and np.any(canvas_result.image_data[:,:,3] > 0):
                    # Convert canvas to image
                    sig_image = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                    
                    # Create white background
                    white_bg = Image.new('RGB', sig_image.size, 'white')
                    white_bg.paste(sig_image, mask=sig_image.split()[3])
                    
                    # Find the bounding box of the signature
                    bbox = white_bg.getbbox()
                    if bbox:
                        # Crop to signature
                        cropped = white_bg.crop(bbox)
                        
                        # Resize if too large
                        max_width, max_height = 200, 80
                        cropped.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
                        
                        # Convert to bytes
                        img_bytes = io.BytesIO()
                        cropped.save(img_bytes, format='PNG')
                        img_bytes.seek(0)
                        signature_img = img_bytes
                
                # Collect all data
                report_data = {
                    'customer_name': customer_name,
                    'project_name': project_name,
                    'contact_person': contact_person,
                    'outlet_location': outlet_location,
                    'contact_number': contact_number,
                    'visit_type': visit_type,
                    'visit_class': visit_class,
                    'date': date.strftime('%Y-%m-%d'),
                    'equipment_inspection': equipment_inspector.get_inspection_summary(),
                    'equipment_list': st.session_state.equipment_list,
                    'work_performed': work_performed,
                    'findings': findings,
                    'recommendations': recommendations,
                    'technician_name': technician_name,
                    'technician_id': technician_id,
                    'service_date': service_date.strftime('%Y-%m-%d'),
                    'technician_signature': signature_img
                }
                
                # Store data in session state for download outside form
                st.session_state.report_data = report_data
                st.session_state.report_generated = True
                st.session_state.customer_name = customer_name
                st.session_state.report_date = date
    
    # Handle report download outside of form
    if st.session_state.get('report_generated', False):
        try:
            # Generate the report
            doc_bytes = create_technical_report(st.session_state.report_data)
            
            # Create filename
            customer_name = st.session_state.customer_name
            date = st.session_state.report_date
            filename = f"Technical_Report_{customer_name.replace(' ', '_')}_{date.strftime('%Y%m%d')}.docx"
            
            # Success message
            st.success("✅ Report generated successfully!")
            
            # Download button (outside form)
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label="📥 Download Report",
                    data=doc_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            
            # Option to generate another report
            if st.button("Generate Another Report", type="secondary"):
                st.session_state.report_generated = False
                st.rerun()
                
        except Exception as e:
            st.error(f"❌ Error generating report: {str(e)}")
            if st.button("Try Again"):
                st.session_state.report_generated = False
                st.rerun()

if __name__ == "__main__":
    main()