import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from bidi.algorithm import get_display  # To handle RTL text correctly
import re
import sys


# Function to read the Excel file
def read_excel(file_path):
    df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets into a dictionary
    sheet_name = list(df.keys())[0]  # Assuming you want the first sheet
    return df[sheet_name]

# Function to split the tables based on the presence of "נוכחות"
def split_tables(df: pd.DataFrame):
    tables = []
    current_table = []
    column_names = df.columns
    for index, row in df.iterrows():
        # Start a new table if "נוכחות" is in the first column
        if "נוכחות" in str(row.iloc[0]).strip():  
            if current_table:  # If there's an existing table, save it
                tables.append(pd.DataFrame(current_table))  # Convert to DataFrame and append
            current_table = [column_names]  # Start a new table
            current_table.append(row)  # Add row to current table

        else:
            current_table.append(row)  # Add row to current table
    
    if current_table:  # Don't forget to add the last table
        tables.append(pd.DataFrame(current_table))
    
    return tables

def get_font_path():
    # Check if the script is running as a PyInstaller bundle
    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle
        font_path = os.path.join(sys._MEIPASS, 'fonts', 'NotoSansHebrew-Regular.ttf')

    else:
        # Running in a normal Python environment
        bundle = os.path.abspath(os.path.dirname(__file__))
        font_path = os.path.join(bundle, 'fonts', 'NotoSansHebrew-Regular.ttf')

    # Path to the font file
    return font_path

# Function to create the PDF from the tables
def create_pdf_from_table(table, title, output_pdf):
    # Set the document to landscape A4
    document = SimpleDocTemplate(output_pdf, pagesize=landscape(A4), 
                             topMargin=10, bottomMargin=10, leftMargin=10, rightMargin=10)
    
    font_path = get_font_path()
    # Register the Noto Sans Hebrew font
    pdfmetrics.registerFont(TTFont('NotoSansHebrew', font_path))  # Adjust path if necessary

    # Styles for the text and table
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    
    # Custom style for Hebrew support - we use "NotoSansHebrew" font
    custom_style = ParagraphStyle('Custom', parent=styleN, fontName="NotoSansHebrew", fontSize=22, alignment=1)

    # List to hold content for the PDF
    content = []
    
    # Add the title as a header on the page
    title_paragraph = Paragraph(get_display(title), custom_style)
    content.append(title_paragraph)
    # Add a Spacer to push the table down a bit
    content.append(Spacer(1, 20))  # 20 points of vertical space (adjust as needed)

    # Prepare data
    data = []
    
    # Add table rows (replacing NaN with empty string)
    for _, row in table.iterrows():
        # Reorder Hebrew text to display properly (right-to-left)
        row_data = [get_display(str(cell)) if not pd.isna(cell) else '' for cell in row]
        data.append(row_data)
    
    # Calculate column widths to fit the landscape A4
    num_columns = len(table.columns)
    first_col_width = 181.75  # 5 cm in points (1 cm = 28.35 points)
    other_col_width = 28.35   # 1 cm in points for other columns
    
    col_widths = [first_col_width] + [other_col_width] * (num_columns - 1)
    row_height = 15  # Height of each row
    
    # Create the table with data
    table_obj = Table(data, colWidths=col_widths, rowHeights=row_height)
    
    # Add table styling
    table_style = TableStyle([ 
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'NotoSansHebrew'),  # Use NotoSansHebrew font for Hebrew
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),  # Reduced bottom padding
        ('TOPPADDING', (0, 0), (-1, -1), 5),  # Reduced top padding
        ('LEFTPADDING', (0, 0), (-1, -1), 5),  # Reduced left padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),  # Reduced right padding
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])
    
    table_obj.setStyle(table_style)
    
    content.append(table_obj)
    
    # Add a page break if necessary
    document.build(content)

# Function to export tables to PDF
def export_tables_to_pdf(tables, output_pdf_base, combine_all=False):
    if combine_all:
        # Combine all tables into one PDF
        combined_output_pdf = os.path.join(output_pdf_base, "combined_output.pdf")
        document = SimpleDocTemplate(combined_output_pdf, pagesize=landscape(A4), 
                             topMargin=10, bottomMargin=10, leftMargin=10, rightMargin=10)
        font_path = get_font_path()
        # Register the Noto Sans Hebrew font
        pdfmetrics.registerFont(TTFont('NotoSansHebrew', font_path))  # Adjust path if necessary

        content = []
        
        for table in tables:
            # Extract the title from the third row, first column
            title = str(table.iloc[2, 0]) if len(table) > 1 else f"Table"
            # Styles for the text and table
            styles = getSampleStyleSheet()
            styleN = styles['Normal']
            
            # Custom style for Hebrew support - we use "NotoSansHebrew" font
            custom_style = ParagraphStyle('Custom', parent=styleN, fontName="NotoSansHebrew", fontSize=22, alignment=1)  # Right alignment for Hebrew

            title_paragraph = Paragraph(get_display(title), custom_style)
            content.append(title_paragraph)
            content.append(Spacer(1, 20))  # Space before table
            
            data = []
            for _, row in table.iterrows():
                row_data = [get_display(str(cell)) if not pd.isna(cell) else '' for cell in row]
                data.append(row_data)

            table_obj = Table(data, colWidths=[181.75] + [28.35] * (len(table.columns) - 1), rowHeights=15)
            table_style = TableStyle([ 
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'NotoSansHebrew'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ])
            table_obj.setStyle(table_style)
            content.append(table_obj)
            content.append(Spacer(1, 20))  # Space between tables
            content.append(PageBreak())
            
        document.build(content)
    else:
        # Export individual PDFs
        for index, table in enumerate(tables):
            title = str(table.iloc[2, 0]) if len(table) > 1 else f"Table_{index + 1}"
            sanitized_title = sanitize_filename(title)
            output_pdf = os.path.join(output_pdf_base, f"{sanitized_title}.pdf")
            create_pdf_from_table(table, title, output_pdf)
            print(f"PDF saved as {output_pdf}")

# Function to sanitize the title for use as a filename
def sanitize_filename(title):
    # Remove any characters that are not valid in a filename
    return re.sub(r'[\\/*?:"<>|]', "_", title)

# GUI for choosing files and folder
class PDFExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Exporter")
        self.root.geometry("600x300")  # Set a bigger window size

        # Add a label and button for selecting the input file
        self.input_file_label = tk.Label(self.root, text="קובץ לא נבחר", width=50, anchor="center")
        self.input_file_label.pack(pady=10)

        self.input_file_button = tk.Button(self.root, text="בחרו קובץ אקסל", command=self.choose_input_file)
        self.input_file_button.pack(pady=10)

        # Add a label and button for selecting the output folder
        self.output_folder_label = tk.Label(self.root, text="תיקייה לא נבחרה", width=50, anchor="center")
        self.output_folder_label.pack(pady=10)

        self.output_folder_button = tk.Button(self.root, text="בחרו תיקייה שבה הקבצים החדשים יווצרו", command=self.choose_output_folder)
        self.output_folder_button.pack(pady=10)

        # Checkbox for combining into one PDF
        self.combine_checkbox_var = tk.BooleanVar()
        self.combine_checkbox = tk.Checkbutton(root, text="ייצא קובץ אחד משולב", variable=self.combine_checkbox_var)
        self.combine_checkbox.pack(padx=100, pady=10)

        # Start button
        self.start_button = tk.Button(self.root, text="התחל", state=tk.DISABLED, command=self.start_export, font=("Helvetica", 16, "bold"), bg="green", fg="black", width=20, height=2)
        self.start_button.pack(pady=20)

    def choose_input_file(self):
        self.input_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
        if self.input_file:
            self.input_file_label.config(text=f"קובץ נבחר: {self.input_file}")
            self.check_start_button()

    def choose_output_folder(self):
        self.output_folder = filedialog.askdirectory(title="Select Output Folder")
        if self.output_folder:
            self.output_folder_label.config(text=f"תיקייה נבחרת: {self.output_folder}")
            self.check_start_button()

    def check_start_button(self):
        if hasattr(self, 'input_file') and hasattr(self, 'output_folder'):
            self.start_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.DISABLED)

    def start_export(self):
        try:
            # Read the Excel file
            df = read_excel(self.input_file)
            
            # Split the tables based on "נוכחות"
            tables = split_tables(df)
            
            # Export the tables to PDFs, with combine option from checkbox
            export_tables_to_pdf(tables, self.output_folder, combine_all=self.combine_checkbox_var.get())
            
            messagebox.showinfo("הצלחה", "הקבצים נוצרו בהצלחה!")
        except Exception as e:
            messagebox.showerror("תקלה", f"An error occurred: {str(e)}")


# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    gui = PDFExporterGUI(root)
    root.mainloop()
