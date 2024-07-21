# If you get ModuleNotFoundError, install the necessary libraries by running the following commands in your terminal:
# pip install fpdf
# pip install xlsxwriter
# pip install pandas

import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
from io import BytesIO
from fpdf import FPDF


def get_on_off_days(start_date, num_days_on, num_days_off, year):
    on_off_days = []
    current_date = start_date
    while current_date.year == year:
        for _ in range(num_days_on):
            if current_date.year == year:
                on_off_days.append((current_date, current_date.strftime("%A"), "On"))
            current_date += timedelta(days=1)
        for _ in range(num_days_off):
            if current_date.year == year:
                on_off_days.append((current_date, current_date.strftime("%A"), "Off"))
            current_date += timedelta(days=1)
    return on_off_days


class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Work Schedule', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def create_table(self, data):
        self.set_font("Arial", size=12)
        col_width = (self.w - 2 * self.l_margin) / 3  # calculate effective page width and divide by number of columns

        # Table header
        self.set_font("Arial", 'B', 12)
        self.cell(col_width, 10, 'Date', 1)
        self.cell(col_width, 10, 'Day of Week', 1)
        self.cell(col_width, 10, 'Status', 1)
        self.ln()

        # Table rows
        self.set_font("Arial", size=12)
        for row in data:
            self.cell(col_width, 10, row['Date'].strftime('%d %B %Y'), 1)
            self.cell(col_width, 10, row['Day of Week'], 1)
            self.cell(col_width, 10, row['Status'], 1)
            self.ln()


def create_pdf(df):
    pdf = PDF()
    pdf.add_page()
    pdf.create_table(df.to_dict('records'))
    return pdf


# Streamlit App
st.title("Work Schedule Generator")

# Sidebar for inputs
st.sidebar.header("Input Parameters")
start_date = st.sidebar.date_input("Start Date", datetime(2024, 4, 8))
num_days_on = st.sidebar.number_input("Number of Days On", min_value=1, step=1, value=4)
num_days_off = st.sidebar.number_input("Number of Days Off", min_value=1, step=1, value=4)
year = start_date.year

if st.sidebar.button("Generate Schedule"):
    with st.spinner("Generating schedule..."):
        on_off_days = get_on_off_days(start_date, num_days_on, num_days_off, year)
        df = pd.DataFrame(on_off_days, columns=["Date", "Day of Week", "Status"])

        # Display the date range
        st.write(f"### Work Schedule for {year}")
        st.write(f"From {start_date.strftime('%d %B %Y')} to {on_off_days[-1][0].strftime('%d %B %Y')}")

        # Display the schedule in a table
        st.table(df)

        # Download buttons
        csv = df.to_csv(index=False)
        st.download_button(label="Download Schedule as CSV", data=csv, file_name='work_schedule.csv', mime='text/csv')

        # Create and download Excel file
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Schedule')
            writer.close()
        excel_buffer.seek(0)
        st.download_button(label="Download Schedule as Excel", data=excel_buffer, file_name='work_schedule.xlsx',
                           mime='application/vnd.ms-excel')

        # Create and download PDF file
        pdf = create_pdf(df)
        pdf_buffer = BytesIO()
        pdf_str = pdf.output(dest='S').encode('latin1')  # Output PDF as string
        pdf_buffer.write(pdf_str)  # Write PDF string to BytesIO buffer
        pdf_buffer.seek(0)
        st.download_button(label="Download Schedule as PDF", data=pdf_buffer, file_name='work_schedule.pdf',
                           mime='application/pdf')
