import os
import tkinter as tk
from tkinter import filedialog
import pdfplumber
import pandas as pd

def get_unique_excel_path(base_path):
    """
    Generate a unique Excel file path by appending a number to the base path.
    If the base path already exists, increment the number until a unique path is found.
    """
    count = 1
    while os.path.exists(f"{base_path}_{count}.xlsx"):
        count += 1
    return f"{base_path}_{count}.xlsx"


def extract_data_from_pdf(pdf_path, bank_type):
    data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(1, len(pdf.pages)):  # Skip the first page
            page = pdf.pages[page_num]

            # Adjust table extraction parameters based on your analysis
            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "intersection_y_tolerance": 2  # adjust as needed
            }

            table = page.extract_table(table_settings)

            if table is not None:
                print(f"Table on page {page_num + 1} successfully extracted.")

                # Identify header row based on the number of columns
                header_row_index = None
                expected_columns = 3  # Assuming Date, Payer/Payee, and Amount columns

                for i, row in enumerate(table):
                    if len(row) == expected_columns:
                        header_row_index = i
                        break

                if header_row_index is not None:
                    for row in table[header_row_index + 1:]:  # Skip header row
                        date_index = 0  # Assuming Date is in the first column
                        description_index = 1  # Assuming Payer/Payee is in the second column
                        amount_index = 2  # Assuming Amount is in the third column

                        # Assuming Date, Payer/Payee, and Amount are present in the row
                        data.append([row[date_index], row[description_index], row[amount_index]])
                else:
                    print("Warning: Header row not found.")

            else:
                print(f"Table on page {page_num + 1} not found or extraction failed.")

    # Create a DataFrame with the extracted data
    columns = ["Date", "Payer/Payee", "Amount"]
    df = pd.DataFrame(data, columns=columns)

    return df

def process_pdf_file():
    # Ask the user to select a PDF file
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    if file_path:
        file_button.pack_forget()
        try:
            # Extract bank type from the filename based on the keywords
            file_name = os.path.basename(file_path)
            selected_bank_type = "Unknown"
            for bank_type, keyword in bank_keywords.items():
                if keyword.lower() in file_name.lower():
                    selected_bank_type = bank_type
                    break

            # Process the PDF file
            data_frame = extract_data_from_pdf(file_path, selected_bank_type)

            # Generate a unique Excel file path
            output_excel_path = get_unique_excel_path("output_data.xlsx")

            # Save the DataFrame to the unique Excel file
            data_frame.to_excel(output_excel_path, index=False)
            feedback_label.config(text=f"Data saved to {output_excel_path}")
        except Exception as e:
            feedback_label.config(text=f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Define the bank keywords for parsing
bank_keywords = {
    "RBC": "RBC",
    "Amex": "Amex"
}

# Create the main Tkinter window
root = tk.Tk()
root.title("PDF Parser")

# Create a label for feedback
feedback_label = tk.Label(root, text="")
feedback_label.pack(pady=10)

# Create and configure a button for file selection
file_button = tk.Button(root, text="Select PDF File", command=process_pdf_file)
file_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()