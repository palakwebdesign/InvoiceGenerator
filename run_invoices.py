import subprocess
import os

# Path to your main invoice generator script
script_path = "C:/InvoiceGenerator/generate_Invoice.py"

def main():
    # Check if the Excel file is filled and exists
    if os.path.exists("C:/InvoiceGenerator/Input_InvoiceGenerator.xlsx"):
        print("Generating invoices...")
        subprocess.run(["python", script_path])
        print("Invoices generated successfully!")
    else:
        print("Error: Excel file not found or needs to be filled.")

if __name__ == "__main__":
    main()