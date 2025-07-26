import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# Extend FPDF to handle header and footer
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, "LAST PAY CERTIFICATE (Obverse)", ln=True, align='C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 10)
        self.cell(0, 10, "This is a system-generated certificate.", align="C")

# Function to generate Last Pay Certificate PDF
def generate_last_pay_certificate(row):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Employee Details
    pdf.set_font("Arial", "", 12)
    pdf.multi_cell(0, 8,
        f"Last Pay Certificate of Shri/Smt. {row.get('Name', 'N/A')}\n"
        f"the office of {row.get('Office', 'N/A')} of\n"
        f"on {row.get('From_Date', 'N/A')} to {row.get('To_Date', 'N/A')} proceeding"
    )
    pdf.ln(3)

    # Paid Up To Section
    pdf.multi_cell(0, 8, f"2. He/She has been paid upto {row.get('Paid_Upto', 'N/A')} at the following rates :")
    pdf.ln(2)

    # Salary Table Header
    pdf.set_font("Arial", "B", 12)
    pdf.cell(90, 8, "Particulars", border=1, align="C")
    pdf.cell(40, 8, "Rate", border=1, align="C")
    pdf.cell(0, 8, "Rs.", border=1, align="C", ln=True)

    # Salary Table Rows
    pdf.set_font("Arial", "", 12)
    particulars = [
        "Substantive Pay", "Officiating Pay", "Special Pay", "Personal Pay",
        "Leave Salary", "Allowance", "D. A", "A. D. A", "Compensatory",
        "Local Allowance", "House Rent Allowance", "Rate of Deductions",
        "General Provident Fund", "Income Tax",
        "State Government Employees Insurance Scheme",
        "    (a) Composite rate", "    (b) Insurance rate only"
    ]
    for part in particulars:
        align = "R" if part.startswith("    ") else "L"
        pdf.cell(90, 8, part.strip(), border=1, align=align)
        pdf.cell(40, 8, "", border=1)
        pdf.cell(0, 8, "", border=1, ln=True)

    pdf.ln(3)

    # Sections 3 to 7
    pdf.multi_cell(0, 8, f"3. His/Her General Provident Fund Account No. {row.get('GPF_Account', 'N/A')} is maintained by the Accountant General/D.A.T.")
    pdf.ln(2)
    pdf.multi_cell(0, 8, f"4. He/She made over charge of the office of {row.get('Office_Charge', 'N/A')} on the {row.get('Charge_Date', 'N/A')} noon of {row.get('Charge_Month', 'N/A')}.")
    pdf.ln(2)
    pdf.multi_cell(0, 8, "5. Recoveries are to be made from the emoluments etc. of the Government Servant as detailed on the reverse.")
    pdf.ln(2)
    pdf.multi_cell(0, 8, "6. He/She entitled to draw the following :")
    pdf.ln(2)
    pdf.multi_cell(0, 8, f"7. He/She is also entitled to joining time for {row.get('Joining_Time', 'N/A')} days.")
    pdf.ln(2)

    # Insurance Table
    pdf.multi_cell(0, 8, "8. He/She finances the Insurance policies below from the Provident Fund.")
    pdf.ln(2)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(80, 8, "Name of Insurance Company", border=1, align="C")
    pdf.cell(40, 8, "No. of Policy", border=1, align="C")
    pdf.cell(40, 8, "Amount of Premium", border=1, align="C")
    pdf.cell(0, 8, "Due Date", border=1, align="C", ln=True)

    pdf.set_font("Arial", "", 12)
    for _ in range(3):
        pdf.cell(80, 8, "", border=1)
        pdf.cell(40, 8, "", border=1)
        pdf.cell(40, 8, "", border=1)
        pdf.cell(0, 8, "", border=1, ln=True)

    pdf.ln(3)

    # Footer section
    pdf.multi_cell(0, 8, "9. The details of the Income Tax recovered from him/her up to date from the beginning of the current financial year are noted on the reverse.")
    pdf.ln(6)
    pdf.cell(80, 8, "Date ..........................................", border=0)
    pdf.cell(0, 8, "Signature ..........................................", border=0, ln=True)
    pdf.cell(80, 8, "", border=0)
    pdf.cell(0, 8, "Designation ........................................", border=0, ln=True)

    return io.BytesIO(pdf.output(dest='S').encode('latin1'))

# Streamlit App
st.title("📑 Last Pay Certificate Generator")
st.write("Upload an Excel file containing employee details to generate Last Pay Certificates.")

uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Excel file loaded successfully!")
        st.dataframe(df.head(10))

        st.write("### 📥 Download Certificates:")
        for index, row in df.iterrows():
            name = str(row.get('Name', f"Employee_{index}"))
            department = str(row.get('Department', 'N/A'))
            file_name = f"{name.replace(' ', '_')}_last_pay_certificate.pdf"

            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"**{name} ({department})**")
            with col2:
                try:
                    pdf_buffer = generate_last_pay_certificate(row)
                    st.download_button(
                        label="Download PDF",
                        data=pdf_buffer,
                        file_name=file_name,
                        mime="application/pdf"
                    )
                except Exception as pdf_err:
                    st.error(f"PDF generation failed for {name}: {pdf_err}")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
