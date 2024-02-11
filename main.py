import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pdfplumber
import pandas as pd

def get_unique_excel_path(base_path):
    count = 1
    while os.path.exists(f"{base_path}_{count}.xlsx"):
        count += 1
    return f"{base_path}_{count}.xlsx"

def extract_data_from_pdf(pdf_path, bank_type):
    data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(1, len(pdf.pages) - 3):  # Skip the first page and last three pages
            page = pdf.pages[page_num]

            # Draw rectangles around detected cells for visual debugging
            im = page.to_image(resolution=150)
            im.debug_tablefinder()
            im.show()

            # Your existing table extraction logic
            table_settings = {
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "intersection_y_tolerance": 2  # adjust as needed
            }

            table = page.extract_table(table_settings)

            if table is not None:
                print(f"Table on page {page_num + 1} successfully extracted.")

                # Identify starting and ending rows of each table
                start_row_index = None
                end_row_index = None
                header_keywords = ["Transaction Date", "Posted Date", "Details", "Amount"]

                for i, row in enumerate(table):
                    if all(keyword in " ".join(row).lower() for keyword in header_keywords):
                        if start_row_index is None:
                            start_row_index = i + 1  # Start from the row after the header
                        elif end_row_index is not None:
                            # Start a new table after encountering a header, skipping the row of the previous table
                            start_row_index = i + 1
                            end_row_index = None
                    elif "Total of Payment Activity" in " ".join(row).lower():
                        end_row_index = i
                        break

                if start_row_index is not None and end_row_index is not None:
                    for row in table[start_row_index:end_row_index]:
                        # Ensure each row has exactly 4 columns
                        if len(row) == 4:
                            data.append(row)
                        else:
                            print("Warning: Row does not have 4 columns.")
                else:
                    print("Warning: No complete table found on this page.")
            else:
                print(f"Table on page {page_num + 1} not found or extraction failed.")

    columns = ["Transaction Date", "Posted Date", "Details", "Amount"]
    df = pd.DataFrame(data, columns=columns)

    return df

def process_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    if file_path:
        file_button.pack_forget()
        try:
            file_name = os.path.basename(file_path)
            selected_bank_type = "Unknown"
            for bank_type, keyword in bank_keywords.items():
                if keyword.lower() in file_name.lower():
                    selected_bank_type = bank_type
                    break

            data_frame = extract_data_from_pdf(file_path, selected_bank_type)
            output_excel_path = get_unique_excel_path("output_data.xlsx")

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
