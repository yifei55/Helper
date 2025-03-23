import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import re
from datetime import datetime, date  # Import date class

def process_mb_files(input_dir, output_file):
    all_data = []
    for filename in os.listdir(input_dir):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith("mb_extracted_data_"):
            filepath = os.path.join(input_dir, filename)
            try:
                df = pd.read_excel(filepath, sheet_name="Zeitraum bis Bedarfsende", header=None)
                customer_item_raw = df.iloc[1, 0]
                match = re.search(r"Sachnummer:\s+(\S+)", customer_item_raw)
                customer_item = match.group(1) if match else None
                if customer_item is None:
                    print(f"Warning: Could not extract Customer Item from {filename}")
                    continue

                calendar_weeks = df.iloc[5, 1:].tolist()
                quantities = df.iloc[7, 1:].tolist()

                temp_data = []
                for cw, qty in zip(calendar_weeks, quantities):
                    if pd.notna(cw) and pd.notna(qty):
                        temp_data.append({"customer_item": customer_item, "calendar_week": cw, "quantity": int(qty)})

                # Sort by year, then month
                temp_data.sort(key=lambda x: (int(x['calendar_week'].split('/')[1]), int(x['calendar_week'].split('/')[0])))
                all_data.extend(temp_data)

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    if all_data:
        create_output_excel(all_data, output_file)
        print(f"Output saved to: {output_file}")
    else:
        print("No data found.")

def create_output_excel(data_list, output_file):
    if not data_list:
        return

    all_calendar_weeks = sorted(list(set([item["calendar_week"] for item in data_list])), key=lambda x: (int(x.split('/')[1]), int(x.split('/')[0])))

    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Yellow fill

    header_row = ["Customer Item"] + all_calendar_weeks
    ws.append(header_row)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    customer_items = sorted(list(set([item["customer_item"] for item in data_list])))
    for customer_item in customer_items:
        row_data = [customer_item]
        for cw in all_calendar_weeks:
            quantity = next((item["quantity"] for item in data_list if item["customer_item"] == customer_item and item["calendar_week"] == cw), 0)
            row_data.append(quantity)
        ws.append(row_data)
    ws.freeze_panes = "B2"

    # Highlight current week
    now = date.today()
    year, week, _ = now.isocalendar()
    current_week = f"{week:02d}/{year}"

    try:
        col_index = all_calendar_weeks.index(current_week) + 2
        ws.cell(row=1, column=col_index).fill = yellow_fill
    except ValueError:
        print(f"Warning: Current week {current_week} not found in data.")

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    wb.save(output_file)

if __name__ == "__main__":
    input_directory = os.getcwd()
    timestamp = datetime.now().strftime("%y%m%d_%H%M")
    output_excel_file = f"mb_extracted_data_{timestamp}.xlsx"
    process_mb_files(input_directory, output_excel_file)
