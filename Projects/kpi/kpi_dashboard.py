import pandas as pd  # Importing pandas for data manipulation and analysis
import matplotlib.pyplot as plt  # Importing matplotlib for data visualization
from fpdf import FPDF  # Importing FPDF for creating PDF documents
import os  # Importing os for interacting with the operating system (e.g., creating directories)
import io  # Importing io for handling in-memory file-like objects
from datetime import datetime  # Importing datetime for working with dates and times


# Path to the Excel file containing sales and marketing data
file_path = 'Sales Data.xls'

# Reading data from different sheets in the Excel file
sales_data = pd.read_excel(file_path, sheet_name='SalesData')
marketing_data = pd.read_excel(file_path, sheet_name='MarketingCosts')

# Converting the 'Date' columns to datetime format and extracting the year
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
sales_data['Year'] = sales_data['Date'].dt.year

marketing_data['Date'] = pd.to_datetime(marketing_data['Date'])
marketing_data['Year'] = marketing_data['Date'].dt.year

# 1. Calculate total sales/revenue per category per year
total_sales_per_category = sales_data.groupby(['Year', 'Category'])['TotalSales'].sum().reset_index()

# 2. Calculate Return on Marketing Spend (RoMS) per category per year
# Sum marketing costs per category and year
marketing_spend_per_category = marketing_data.groupby(['Year', 'Category'])['Cost'].sum().reset_index()
# Merge sales data with marketing data on Year and Category
kpi_data = pd.merge(total_sales_per_category, marketing_spend_per_category, on=['Year', 'Category'])
# Calculate RoMS as Total Sales divided by Marketing Costs
kpi_data['RoMS'] = kpi_data['TotalSales'] / kpi_data['Cost']

# 3. Calculate average order value per category per year
# Order Value is Total Sales divided by Quantity Sold
sales_data['OrderValue'] = sales_data['TotalSales'] / sales_data['QuantitySold']
# Group by Year and Category to find average Order Value
avg_order_value_per_category = sales_data.groupby(['Year', 'Category'])['OrderValue'].mean().reset_index()

# Create a custom PDF class to generate the KPI dashboard
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'KPI Dashboard', ln=True, align='C')
        self.ln(10)

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, ln=True)
        self.ln(5)

    def chapter_body(self, body):
        self.set_font('Arial', '', 10)
        self.multi_cell(0, 10, body)
        self.ln()

    def add_chart(self, buffer):
        buffer.seek(0)
        self.image(buffer, w=180)
        self.ln(10)

# Instantiate the PDF object
pdf = PDF()

# Function to create visualizations and add them to the PDF
def create_visualizations():
    # Loop through each unique year in the data
    for year in total_sales_per_category['Year'].unique():
        pdf.add_page()  # Add a new page for each year

        # 1. Visualization for Total Sales/Revenue per Category
        yearly_sales_data = total_sales_per_category[total_sales_per_category['Year'] == year]
        plt.figure(figsize=(10, 6))
        plt.bar(yearly_sales_data['Category'], yearly_sales_data['TotalSales'], color='skyblue')
        plt.title(f'Total Sales/Revenue per Category - {year}')
        plt.xlabel('Category')
        plt.ylabel('Total Sales')
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Save the plot to an in-memory buffer and add it to the PDF
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        pdf.add_chart(buf)
        buf.close()
        plt.close()

        # Add Total Sales/Revenue per Category Calculation to the PDF
        total_sales_text = yearly_sales_data.to_string(index=False)
        pdf.chapter_title(f'Total Sales/Revenue per Category - {year}')
        pdf.chapter_body(total_sales_text)

        # 2. Visualization for Return on Marketing Spend (RoMS) per Category
        yearly_roms_data = kpi_data[kpi_data['Year'] == year]
        plt.figure(figsize=(10, 6))
        plt.bar(yearly_roms_data['Category'], yearly_roms_data['RoMS'], color='lightgreen')
        plt.title(f'Return on Marketing Spend per Category - {year}')
        plt.xlabel('Category')
        plt.ylabel('RoMS')
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Save the plot to an in-memory buffer and add it to the PDF
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        pdf.add_chart(buf)
        buf.close()
        plt.close()

        # Add RoMS Calculation to the PDF
        roms_text = yearly_roms_data[['Year', 'Category', 'RoMS']].to_string(index=False)
        pdf.chapter_title(f'Return on Marketing Spend per Category - {year}')
        pdf.chapter_body(roms_text)

        # 3. Visualization for Average Order Value per Category
        yearly_aov_data = avg_order_value_per_category[avg_order_value_per_category['Year'] == year]
        plt.figure(figsize=(10, 6))
        plt.bar(yearly_aov_data['Category'], yearly_aov_data['OrderValue'], color='salmon')
        plt.title(f'Average Order Value per Category - {year}')
        plt.xlabel('Category')
        plt.ylabel('Average Order Value')
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Save the plot to an in-memory buffer and add it to the PDF
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        pdf.add_chart(buf)
        buf.close()
        plt.close()

        # Add Average Order Value Calculation to the PDF
        aov_text = yearly_aov_data.to_string(index=False)
        pdf.chapter_title(f'Average Order Value per Category - {year}')
        pdf.chapter_body(aov_text)

# Function to save the KPI dashboard as a PDF in a timestamped directory
def save_kpi_dashboard():
    # Create a directory named with the current date
    output_dir = os.path.join('output', datetime.now().strftime('%Y_%m_%d'))
    os.makedirs(output_dir, exist_ok=True)
    # Define the output path for the PDF
    pdf_output_path = os.path.join(output_dir, 'kpi_dashboard.pdf')
    # Save the PDF to the output path
    pdf.output(pdf_output_path)
    print(f"KPI dashboard with visuals and calculations has been created and saved as '{pdf_output_path}'.")

# Generate the KPI dashboard and save it
create_visualizations()
save_kpi_dashboard()
