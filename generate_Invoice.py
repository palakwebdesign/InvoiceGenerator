from PIL import Image, ImageDraw, ImageFont
import textwrap
import xlwings as xw
import inflect

# Configuration
FOLDER_PATH = "C:/InvoiceGenerator/Invoices"
EXCEL_FILE_PATH = "C:/InvoiceGenerator/Input_InvoiceGenerator.xlsx"
IMAGE_TEMPLATE_PATH = 'BipinEnterprises_CGST.jpg'
FONT_PATH = "arial"
FONT_SIZE_SMALL = 14
FONT_SIZE_LARGE = 16
LINE_SPACING = 5
TEXT_COLOR = (0, 0, 0)

# Initialize font
def load_fonts():
    try:
        myfont = ImageFont.truetype(FONT_PATH, FONT_SIZE_SMALL)
        header_font = ImageFont.truetype(FONT_PATH, FONT_SIZE_LARGE)
        return myfont, header_font
    except IOError as e:
        print(f"Font file not found: {e}")
        raise

# Load Excel data
def load_excel_data(file_path):
    try:
        wb = xw.Book(file_path)
        sheet_header = wb.sheets[0]
        sheet_products = wb.sheets[1]
        print("Excel data loaded successfully.")
        return sheet_header, sheet_products
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        raise

# Fetch header data from Excel for a specific invoice
def fetch_header_data(sheet, invoice_id):
    header_data = {
        "Name": {"cell": "B2", "coordinates": (68, 217)},
        "Address": {"cell": "C2", "coordinates": (68, 235)},
        "Invoice_ID": {"cell": "A2", "coordinates": (481, 213)},
        "Invoice_Date": {"cell": "D2", "coordinates": (625, 213)},
        "Challan_No": {"cell": "E2", "coordinates": (485, 235)},
        "Challan_Date": {"cell": "F2", "coordinates": (625, 235)},
        "PO_No": {"cell": "G2", "coordinates": (481, 253)},
        "PO_Date": {"cell": "H2", "coordinates": (625, 255)},
        "Dispatched_through": {"cell": "I2", "coordinates": (545, 275)},
        "LR_No": {"cell": "J2", "coordinates": (484, 293)},
        "LR_Date": {"cell": "K2", "coordinates": (628, 293)},
        "GST_No": {"cell": "L2", "coordinates": (513, 313)}
    }
    wrapper = textwrap.TextWrapper(width=50)
    for key, value in header_data.items():
        cell_value = sheet[value["cell"]].value
        header_data[key]["value"] = wrapper.fill(text=str(cell_value)) if key == "Address" else str(cell_value) if cell_value else ""

    print(f"Fetched header data for Invoice ID {invoice_id}: {header_data}")

    return header_data

# Calculate totals for products under the specific Invoice_ID
def calculate_totals(sheet, invoice_id):
    total_value = 0.0
    packing_charges_total = 0.0
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

    for i in range(2, last_row + 1):
        if sheet[f"A{i}"].value == invoice_id:  # Check for matching Invoice_ID
            rs_value = sheet[f"F{i}"].value  # Rs column
            packing_chrg_value = sheet[f"G{i}"].value  # Packing charges column

            if rs_value is not None:
                total_value += float(rs_value)

            # Add packing charges if available
            if packing_chrg_value is not None:
                packing_charges_total += float(packing_chrg_value)

    # Calculate taxes and grand total
    total_with_packing = total_value + packing_charges_total
    CGST = SGST = total_with_packing * 0.09  # 9% GST
    grand_total = total_with_packing + CGST + SGST

    # Round off the grand total and adjust packing charges
    rounding_difference = round(grand_total) - grand_total
    packing_charges_total += rounding_difference
    grand_total = round(grand_total)

    return "{:.2f}".format(CGST), "{:.2f}".format(SGST), total_with_packing, grand_total

# Convert number to words
def number_to_words(n):
    p = inflect.engine()
    return p.number_to_words(n).capitalize() + " Only"

# Draw text on image
def draw_text(draw, font, coordinates, text, line_spacing=LINE_SPACING):
    wrapper = textwrap.TextWrapper(width=40)
    wrapped_text = wrapper.wrap(text=text)
    y_position = coordinates[1]
    for line in wrapped_text:
        bbox = draw.textbbox((coordinates[0], y_position), line, font=font)
        text_height = bbox[3] - bbox[1]
        draw.text((coordinates[0], y_position), line, fill=(0, 0, 0), font=font)
        y_position += text_height + line_spacing
    return y_position

def draw_products(draw, sheet, invoice_id, myfont):
    product_columns = {
        # "SRNO": {"col": "A", "x_offset": 33},
        "Prod_Desc": {"col": "B", "x_offset": 75},
        "Qty": {"col": "C", "x_offset": 412},
        "HSN": {"col": "D", "x_offset": 457},
        "Rate_PP": {"col": "E", "x_offset": 525},
        "Rs": {"col": "F", "x_offset": 581},
        "packing_chrg": {"col": "G", "x_offset": 581}
    }

    start_y = 375  # Starting y position for products
    vertical_spacing = 20
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

    product_count = 0  # Count how many products have been drawn


    for i in range(2, last_row + 1):  # Start from row 2 (first product)
        if sheet[f"A{i}"].value == invoice_id:  # Only draw products for this Invoice_ID
            product_count += 1  # Increment product count

            current_y = start_y + (product_count - 1) * vertical_spacing  # Calculate y-coordinate based on count
            draw.text((33, current_y), str(product_count), fill=(0, 0, 0), font=myfont)
            print(f"Drawing product for Invoice ID {invoice_id} at row {i}")
            
            for key, col_info in product_columns.items():
                cell_value = sheet[f"{col_info['col']}{i}"].value

                if key == "Rs":
                    if cell_value is not None:
                        parts = str(cell_value).split(".")
                        rupee = parts[0]
                        paisa = parts[1] if len(parts) > 1 else "00"
                        draw.text((col_info["x_offset"], current_y), str(rupee), fill=(0, 0, 0), font=myfont)
                        draw.text((col_info["x_offset"] + 100, current_y), str(paisa), fill=(0, 0, 0), font=myfont)

                elif key == "packing_chrg":
                    if cell_value is not None:
                        draw.text((col_info["x_offset"], 697), str(float(cell_value)), fill=(0, 0, 0), font=myfont)
                else:
                    value = str(cell_value) if cell_value is not None else ""
                    draw.text((col_info["x_offset"], current_y), value, fill=(0, 0, 0), font=myfont)

# Main function
def main():
    sheet_header, sheet_products = load_excel_data(EXCEL_FILE_PATH)
    myfont, header_font = load_fonts()
    
    last_row = sheet_header.range("A" + str(sheet_header.cells.last_cell.row)).end("up").row
    invoice_ids = {sheet_header[f"A{i}"].value for i in range(2, last_row + 1)}  # Collect distinct Invoice_IDs

    for invoice_id in invoice_ids:
        img = Image.open(IMAGE_TEMPLATE_PATH)
        draw = ImageDraw.Draw(img)

        # Fetch and draw header data for the specific invoice
        header_data = fetch_header_data(sheet_header, invoice_id)
        for key, value in header_data.items():
            draw.text(value["coordinates"], value["value"], fill=(0, 0, 0), font=header_font if key != "Address" else myfont)

        # Calculate and draw totals
        CGST, SGST, total_with_packing, grand_total = calculate_totals(sheet_products, invoice_id)

        # Split CGST and SGST into Rupee and Paisa
        def split_rupee_paisa(value):
            parts = value.split(".")
            rupee = parts[0]
            paisa = parts[1] if len(parts) > 1 else "00"
            return rupee, paisa

        CGST_rupee, CGST_paisa = split_rupee_paisa(CGST)
        SGST_rupee, SGST_paisa = split_rupee_paisa(SGST)
        total_rupee, total_paisa = split_rupee_paisa("{:.2f}".format(total_with_packing))
        grand_total_rupee, grand_total_paisa = split_rupee_paisa(str(grand_total))

        # Draw totals on the image
        totals = {
            "Total": (total_rupee, total_paisa, (585, 764)),
            "CGST": (CGST_rupee, CGST_paisa, (585, 794)),
            "SGST": (SGST_rupee, SGST_paisa, (585, 825)),
            "Grand Total": (grand_total_rupee, grand_total_paisa, (585, 880)),
            # "Inwords": {"value": number_to_words(grand_total).upper(), "coordinate": (93, 765)},
        }

        for label, (rupee, paisa, coords) in totals.items():
            draw.text(coords, rupee, fill=(0, 0, 0), font=myfont)
            if paisa:  # Only draw paisa if it exists
                draw.text((coords[0] + 100, coords[1]), paisa, fill=(0, 0, 0), font=myfont)
        
        draw.text((525, 793), "9", fill=(0, 0, 0), font=myfont)  # CGST percentage
        draw.text((525, 820), "9", fill=(0, 0, 0), font=myfont)  # SGST percentage

        # Draw the invoice products
        draw_products(draw, sheet_products, invoice_id, myfont)

        grand_total_in_words = number_to_words(grand_total)
        draw.text((93, 765), grand_total_in_words, fill=(0, 0, 0), font=myfont)

        try:
            img.save(f"{FOLDER_PATH}/Invoice_{invoice_id}.jpg")
            print(f"Invoice saved as Invoice_{invoice_id}.jpg")
        except Exception as e:
            print(f"Error saving image for Invoice ID {invoice_id}: {e}")
        img.close()

if __name__ == "__main__":
    main()
