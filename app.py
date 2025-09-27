import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import io
import base64
from datetime import datetime
import os

def create_certificate(data):
    """Create a Certificate of Conformity DOCX file"""
    try:
        # Create a new document
        doc = Document()
        
        # Add title
        title = doc.add_heading('CERTIFICATE OF CONFORMITY', 0)
        title.alignment = 1  # Center alignment
        
        # Add subtitle
        subtitle = doc.add_heading('Vehicle Handover Certificate', 1)
        subtitle.alignment = 1
        
        doc.add_paragraph()  # Add space
        
        # Certificate details
        p1 = doc.add_paragraph()
        p1.add_run('This is to certify that the following vehicle has been inspected and conforms to the agreed specifications:')
        
        doc.add_paragraph()  # Add space
        
        # Vehicle details table
        table = doc.add_table(rows=10, cols=2)
        table.style = 'Light Grid Accent 1'
        
        # Table data
        details = [
            ('Certificate Number:', data.get('certificate_number', '')),
            ('Date of Issue:', data.get('issue_date', '')),
            ('Seller Name:', data.get('seller_name', '')),
            ('Buyer Name:', data.get('buyer_name', '')),
            ('Vehicle Make:', data.get('vehicle_make', '')),
            ('Vehicle Model:', data.get('vehicle_model', '')),
            ('Vehicle Year:', data.get('vehicle_year', '')),
            ('VIN Number:', data.get('vin_number', '')),
            ('License Plate:', data.get('license_plate', '')),
            ('Odometer Reading:', data.get('odometer_reading', ''))
        ]
        
        for i, (label, value) in enumerate(details):
            table.cell(i, 0).text = label
            table.cell(i, 1).text = str(value)
        
        doc.add_paragraph()  # Add space
        
        # Conditions section
        conditions = doc.add_paragraph()
        conditions.add_run('Conditions of Conformity:').bold = True
        
        condition_list = [
            "Vehicle is in good mechanical condition",
            "All documents are complete and valid",
            "No outstanding finance or liens",
            "Vehicle is free from major accidents",
            "All accessories are functioning properly"
        ]
        
        for condition in condition_list:
            doc.add_paragraph(condition, style='List Bullet')
        
        doc.add_paragraph()  # Add space
        
        # Signatures section
        sig_table = doc.add_table(rows=1, cols=2)
        sig_table.style = 'Light Grid Accent 1'
        
        # Seller signature
        sig_table.cell(0, 0).text = "Seller's Signature: __________________"
        sig_table.cell(0, 0).add_paragraph(f"Name: {data.get('seller_name', '')}")
        sig_table.cell(0, 0).add_paragraph(f"Date: {data.get('issue_date', '')}")
        
        # Buyer signature
        sig_table.cell(0, 1).text = "Buyer's Signature: __________________"
        sig_table.cell(0, 1).add_paragraph(f"Name: {data.get('buyer_name', '')}")
        sig_table.cell(0, 1).add_paragraph(f"Date: {data.get('issue_date', '')}")
        
        # Save to bytes buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        return buffer
    
    except Exception as e:
        st.error(f"Error creating certificate: {str(e)}")
        return None

def main():
    st.set_page_config(page_title="Certificate of Conformity Generator", layout="wide")
    
    st.title("🚗 Certificate of Conformity Generator")
    st.write("Generate certificates for vehicle handover to buyers")
    
    # Initialize session state
    if 'certificate_data' not in st.session_state:
        st.session_state.certificate_data = None
    if 'certificate_file' not in st.session_state:
        st.session_state.certificate_file = None
    
    # Sidebar for CSV upload
    st.sidebar.header("📁 Upload CSV File")
    uploaded_file = st.sidebar.file_uploader("Choose a CSV file", type=['csv'])
    
    if uploaded_file is not None:
        try:
            # Read CSV file
            df = pd.read_csv(uploaded_file)
            st.sidebar.success("CSV file uploaded successfully!")
            
            # Display data preview
            st.subheader("📊 CSV Data Preview")
            st.dataframe(df)
            
            # Select record to generate certificate for
            if len(df) > 0:
                record_index = st.selectbox("Select record to generate certificate:", range(len(df)))
                
                if st.button("Generate Certificate"):
                    # Convert selected row to dictionary
                    selected_data = df.iloc[record_index].to_dict()
                    
                    # Add missing fields with default values
                    if 'issue_date' not in selected_data:
                        selected_data['issue_date'] = datetime.now().strftime("%Y-%m-%d")
                    if 'certificate_number' not in selected_data:
                        selected_data['certificate_number'] = f"COC-{datetime.now().strftime('%Y%m%d')}-{record_index+1:03d}"
                    
                    # Store in session state
                    st.session_state.certificate_data = selected_data
                    
                    # Generate certificate
                    with st.spinner("Generating certificate..."):
                        certificate_buffer = create_certificate(selected_data)
                        
                        if certificate_buffer:
                            st.session_state.certificate_file = certificate_buffer
                            st.success("Certificate generated successfully!")
            
        except Exception as e:
            st.error(f"Error reading CSV file: {str(e)}")
    
    # Manual data entry option
    st.sidebar.header("✍️ Manual Entry")
    with st.sidebar.form("manual_entry"):
        st.write("Or enter details manually:")
        seller_name = st.text_input("Seller Name")
        buyer_name = st.text_input("Buyer Name")
        vehicle_make = st.text_input("Vehicle Make")
        vehicle_model = st.text_input("Vehicle Model")
        vehicle_year = st.text_input("Vehicle Year")
        vin_number = st.text_input("VIN Number")
        license_plate = st.text_input("License Plate")
        odometer_reading = st.text_input("Odometer Reading")
        
        if st.form_submit_button("Generate from Manual Entry"):
            manual_data = {
                'seller_name': seller_name,
                'buyer_name': buyer_name,
                'vehicle_make': vehicle_make,
                'vehicle_model': vehicle_model,
                'vehicle_year': vehicle_year,
                'vin_number': vin_number,
                'license_plate': license_plate,
                'odometer_reading': odometer_reading,
                'issue_date': datetime.now().strftime("%Y-%m-%d"),
                'certificate_number': f"COC-{datetime.now().strftime('%Y%m%d')}-MAN"
            }
            
            st.session_state.certificate_data = manual_data
            
            with st.spinner("Generating certificate..."):
                certificate_buffer = create_certificate(manual_data)
                
                if certificate_buffer:
                    st.session_state.certificate_file = certificate_buffer
                    st.success("Certificate generated successfully!")
    
    # Display and download certificate
    if st.session_state.certificate_file and st.session_state.certificate_data:
        st.subheader("📄 Generated Certificate")
        
        # Display certificate data
        st.write("**Certificate Details:**")
        cert_data = st.session_state.certificate_data
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"**Certificate No:** {cert_data.get('certificate_number', '')}")
            st.write(f"**Issue Date:** {cert_data.get('issue_date', '')}")
            st.write(f"**Seller:** {cert_data.get('seller_name', '')}")
            st.write(f"**Buyer:** {cert_data.get('buyer_name', '')}")
        
        with col2:
            st.write(f"**Vehicle:** {cert_data.get('vehicle_make', '')} {cert_data.get('vehicle_model', '')}")
            st.write(f"**Year:** {cert_data.get('vehicle_year', '')}")
            st.write(f"**VIN:** {cert_data.get('vin_number', '')}")
            st.write(f"**License Plate:** {cert_data.get('license_plate', '')}")
        
        # Download button
        st.download_button(
            label="📥 Download Certificate",
            data=st.session_state.certificate_file,
            file_name=f"certificate_of_conformity_{cert_data.get('certificate_number', '')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        # Print button (opens in new tab for printing)
        certificate_b64 = base64.b64encode(st.session_state.certificate_file.getvalue()).decode()
        href = f'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{certificate_b64}'
        
        st.markdown(
            f'<a href="{href}" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; border-radius: 4px; margin-right: 10px;">🖨️ Open for Printing</a>',
            unsafe_allow_html=True
        )
        
        # Clear button
        if st.button("🗑️ Clear Certificate"):
            st.session_state.certificate_data = None
            st.session_state.certificate_file = None
            st.rerun()

if __name__ == "__main__":
    main()
