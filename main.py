import os
import tkinter as tk
from tkinter import filedialog
import pdfplumber
import pandas as pd

def get_unique_excel_path(base_path):
    # If the file already exists, append a number until a unique name is found
    file_count = 1
    while os.path.exists(base_path):
        base_path = base_path.rsplit('.', 1)[0] + f"_{file_count}.xlsx"
        file_count += 1
    return base_path

def extract_data_from_pdf(pdf_path, bank_type):
    data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]

            # Adjust table extraction parameters
            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 8
            }

            table = page.extract_table(table_settings)

            # Placeholder for bank-specific logic
            if bank_type == "RBC":
                # Apply RBC-specific processing
                pass
            elif bank_type == "Amex":
                # Skip the first page for Amex
                if page_num == 0:
                    continue

                # Apply Amex-specific processing
                print("AMEX detected. Applying AMEX-specific processing.")
                pass

            # Check if table extraction was successful
            if table is not None:
                print(f"Table on page {page_num + 1} successfully extracted.")
                # Assuming the data you need is in a specific row and column range
                # Adjust these indices based on the layout of your bank statement
                data += [row[1:4] for row in table[3:]]
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
        print(f"Data saved to {output_excel_path}")

# Define the bank keywords for parsing
bank_keywords = {
    "RBC": "RBC",
    "Amex": "Amex"
}

# Create the main Tkinter window
root = tk.Tk()
root.title("PDF Parser")

# Create and configure a button for file selection
file_button = tk.Button(root, text="Select PDF File", command=process_pdf_file)
file_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
