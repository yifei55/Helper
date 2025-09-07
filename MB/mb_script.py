import pandas as pd
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
import re
from datetime import datetime, date
import zipfile
import shutil
def get_current_calendar_week():
    """Calculates the current calendar week in the format 'WW/YYYY'."""
    now = date.today()
    year, week, _ = now.isocalendar()
    return f"{week:02d}/{year}"

def process_mb_files(input_dir, output_file, mercedes_file):
    """Processes MB files, extracts data, and updates the Mercedes file."""
    all_data = []
    all_data_bedarfs = []
    all_data_ruckstand = []
    print(f"Processing files in directory: {input_dir}")  # Add this line
    # --- MODIFICATION: Only process files that start with 'BKM Lieferbeziehung' ---
    for filename in os.listdir(input_dir):
        # This check is now more specific to avoid processing output/master files.
        if filename.startswith('BKM Lieferbeziehung') and filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            print(f"Processing MB report file: {filename}")
            filepath = os.path.join(input_dir, filename)
            try:
                df = pd.read_excel(filepath, sheet_name="Zeitraum bis Bedarfsende", header=None)
                customer_item_raw = df.iloc[1, 0]
                match = re.search(r"Sachnummer:\s+(\S+)", customer_item_raw)
                customer_item = match.group(1) if match else None
                if customer_item is None:
                    print(f"Warning: Could not extract Customer Item from {filename}")
                    continue
                print(f"Sachnummer found: {customer_item}")

                # Extract calendar weeks from row 6 (index 5)
                calendar_weeks = df.iloc[5, 1:].tolist()

                # Extract data for Bedarf
                quantities_bedarf = df.iloc[7, 1:].tolist()
                temp_data_bedarf = []
                for cw, qty in zip(calendar_weeks, quantities_bedarf):
                    if pd.notna(cw) and pd.notna(qty):
                        temp_data_bedarf.append({"customer_item": customer_item, "calendar_week": cw, "quantity": int(qty)})
                temp_data_bedarf.sort(key=lambda x: (int(x['calendar_week'].split('/')[1]), int(x['calendar_week'].split('/')[0])))
                all_data_bedarfs.extend(temp_data_bedarf)

                # Find rows with "ABS" followed by a number
                abs_rows = []
                for index, row in df.iterrows():
                    first_cell_value = row.iloc[0]
                    if isinstance(first_cell_value, str) and re.match(r"^\s*ABS\s+\d+\w*", first_cell_value):
                        abs_rows.append(index)
                if abs_rows:
                    print(f"ABS rows found in {filename}")
                else:
                    print(f"No ABS rows found in {filename}")

                # Extract data for each ABS row
                for row_index in abs_rows:
                    abs_value_raw = df.iloc[row_index, 0].strip()
                    if isinstance(abs_value_raw, str) and abs_value_raw.startswith("ABS "):
                        abs_value = abs_value_raw.split(" ", 1)[1]
                    else:
                        abs_value = abs_value_raw
                    quantities = df.iloc[row_index, 1:].tolist()

                    # Get current week index
                    current_week = get_current_calendar_week()
                    try:
                        current_week_index = calendar_weeks.index(current_week)
                    except ValueError:
                        print(f"Warning: Current week {current_week} not found in {filename}")
                        continue

                    # Extract the 5 data points
                    extracted_quantities = quantities[current_week_index:current_week_index + 5]

                    # Ensure we have 5 data points, fill with None if not enough
                    while len(extracted_quantities) < 5:
                        extracted_quantities.append(None)

                    all_data.append({
                        "customer_item": customer_item,
                        "abs_value": abs_value,
                        "quantities": extracted_quantities,
                        "calendar_weeks": [calendar_weeks[i] if i < len(calendar_weeks) else None for i in range(current_week_index, current_week_index + 5)]
                    })
                # Extract data for Rückstand
                df_bkm = pd.read_excel(filepath, sheet_name="BKM Lieferbeziehung", header=None)
                for row_index in abs_rows:
                    abs_value_raw = df.iloc[row_index, 0].strip()
                    if isinstance(abs_value_raw, str) and abs_value_raw.startswith("ABS "):
                        abs_value = abs_value_raw.split(" ", 1)[1]
                    else:
                        abs_value = abs_value_raw
                    ruckstand = df_bkm.iloc[row_index, 21]
                    all_data_ruckstand.append({
                        "customer_item": customer_item,
                        "abs_value": abs_value,
                        "ruckstand":ruckstand
                    })

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    if all_data:
        # --- New Requirement: Create a timestamped copy and update it ---
        if not os.path.exists(mercedes_file):
            print(f"Error: Original Mercedes file '{mercedes_file}' not found. Cannot create a copy.")
        else:
            try:
                # Generate timestamp YYMMDDHH
                timestamp = datetime.now().strftime("%y%m%d%H")
                file_name, file_extension = os.path.splitext(mercedes_file)
                
                # Create the new filename for the copy
                mercedes_copy_file = f"{file_name}_{timestamp}{file_extension}"
                
                # Copy the file
                shutil.copy2(mercedes_file, mercedes_copy_file)
                print(f"Created a copy for updating: '{mercedes_copy_file}'")
                update_mercedes_file(all_data, all_data_ruckstand, mercedes_copy_file)
            except Exception as e:
                print(f"Error creating or updating the Mercedes file copy: {e}")
    if all_data_bedarfs:
        create_output_excel(all_data_bedarfs, output_file)
        print(f"Output saved to: {output_file}")
    else:
        print("No data found.")

def update_mercedes_file(data_list, data_list_ruckstand, mercedes_file):
    if not os.path.exists(mercedes_file):
        print(f"Error: Mercedes file '{mercedes_file}' not found.")
        return
    print(f"Mercedes file exists: {mercedes_file}")
    try:
        wb = load_workbook(mercedes_file)
        print(f"Mercedes file loaded successfully: {mercedes_file}")
        ws = wb["EDI"]  # Access the "EDI" sheet
    except FileNotFoundError:
        print(f"Error: Mercedes file '{mercedes_file}' not found.")
        return
    except KeyError:
        print(f"Error: Sheet 'EDI' not found in '{mercedes_file}'.")
        return
    except PermissionError:
        print(f"Error: Permission denied to access '{mercedes_file}'. Is the file open in another program?")
        return
    except zipfile.BadZipFile:
        print(f"Error: '{mercedes_file}' is corrupted or not a valid Excel file.")
        return
    except Exception as e:
        print(f"Error loading Mercedes file: {e}")
        return

    header_font = Font(bold=True)
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Create a dictionary to store existing data for faster lookup
    existing_data = {}
    for row in ws.iter_rows(min_row=2):
        sachnummer = row[3+10].value
        abs_value_raw = row[6+9].value
        if sachnummer and abs_value_raw:
            if isinstance(abs_value_raw, str) and abs_value_raw.startswith("ABS "):
                abs_value = abs_value_raw.split(" ", 1)[1]
                row[6+9].value = abs_value
            else:
                abs_value = abs_value_raw
            existing_data[(sachnummer, str(abs_value))] = row
    
    # Create a dictionary for faster lookup of Rückstand
    ruckstand_data = {}
    for item in data_list_ruckstand:
        ruckstand_data[(item["customer_item"], item["abs_value"])] = item["ruckstand"]

    for data in data_list:
        sachnummer = data["customer_item"]
        abs_value = data["abs_value"]
        quantities = data["quantities"]
        calendar_weeks = data["calendar_weeks"]

        # Find the row or create a new one
        match_key = (sachnummer, str(abs_value))
        if match_key in existing_data:
            row = existing_data[match_key]
        else:
            # Add a new row
            new_row = [None] * ws.max_column
            ws.append(new_row)
            rows_list = list(ws.rows)
            row = rows_list[-1]
            row[3+10].value = sachnummer  # Sachnummer
            row[6+9].value = abs_value  # ABS
            # Highlight the new row in yellow
            for cell in row:
                cell.fill = yellow_fill
        # Fill Rückstand
        if match_key in ruckstand_data:
            row[9+8].value = ruckstand_data[match_key]
        # Fill in the quantities and calendar weeks
        columns = [11+8, 13+8, 15+8, 17+8, 19+8]  #  (Corrected indices)
        for i, (qty, cw) in enumerate(zip(quantities, calendar_weeks)):
            if qty is not None:
                row[columns[i]-1].value = qty
            if cw is not None:
                ws.cell(row=1, column=columns[i]).value = cw
    try:
        wb.save(mercedes_file)
    except PermissionError:
        print(f"Error: Permission denied to save '{mercedes_file}'. Is the file open in another program?")
        return
    except Exception as e:
        print(f"Error saving Mercedes file: {e}")
        return


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

    # Define the desired order
    desired_order = [
        "A2238305705",
        "A2238305706",
        "A2068305905",
        "A2548302703",
        "A2148308201",
        "A2978306501",
        "A2979970200",
        "A2979970600",
        "A0005003700",
        "A0005003101",
        "A0005004901",
        "A0005005301",
        "A0005002901",
    ]

    # Extract customer items and sort them based on the desired order or alphabetically
    customer_items = sorted(list(set([item["customer_item"] for item in data_list])), key=lambda x: (desired_order.index(x) if x in desired_order else len(desired_order)))

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

def update_forecast_file(source_data_file, forecast_file_path):
    """
    Updates the 'forecast for all projects.xlsx' with data from the generated
    'mb_extracted_data_...' file.
    """
    print(f"\nStarting update of master forecast file: '{forecast_file_path}'")
    if not os.path.exists(source_data_file):
        print(f"Error: Source data file '{source_data_file}' not found. Skipping forecast update.")
        return
    if not os.path.exists(forecast_file_path):
        print(f"Error: Master forecast file '{forecast_file_path}' not found. Skipping forecast update.")
        return

    try:
        # 1. Load the source data (mb_extracted_data_...)
        source_df = pd.read_excel(source_data_file)

        # 2. Load the destination forecast workbook and get the active sheet
        wb = load_workbook(forecast_file_path)
        ws = wb.active

        # 3. Get the week numbers from the header of the forecast file
        # Assuming headers are in the first row and week numbers start from the second column
        forecast_headers = [cell.value for cell in ws[1]]
        # Find the column index for 'Customer Item' or a similar identifier
        try:
            item_col_name = 'Customer Item' # Adjust if the name is different
            item_col_idx = forecast_headers.index(item_col_name) + 1
        except ValueError:
            print(f"Error: Column '{item_col_name}' not found in '{forecast_file_path}'. Cannot match items.")
            return

        # --- MODIFICATION: Make week number detection more robust ---
        # It will now try to convert header values to integers.
        week_cols = {}
        for i, h in enumerate(forecast_headers):
            try:
                # --- MODIFICATION: Handle headers like "week 36" ---
                # Use regex to find any number in the header string.
                match = re.search(r'\d+', str(h))
                if match:
                    week_num = int(match.group(0))
                    week_cols[week_num] = i + 1
            except (ValueError, TypeError):
                continue # Ignore headers that can't be processed

        if not week_cols:
            print(f"Error: No integer week number columns found in '{forecast_file_path}'.")
            return

        start_week = min(week_cols.keys())
        print(f"Master forecast starts at week {start_week}. Weeks before this will be ignored.")

        # 4. Create a map of 'Customer Item' to its row number in the forecast sheet
        item_row_map = {str(ws.cell(row=r, column=item_col_idx).value): r for r in range(2, ws.max_row + 1)}

        # 5. Iterate through the source data and update the forecast sheet
        for _, source_row in source_df.iterrows():
            customer_item = str(source_row['Customer Item'])
            if customer_item in item_row_map:
                target_row_idx = item_row_map[customer_item]
                # Iterate through the week columns in the source data
                for col_name in source_df.columns:
                    if '/' in str(col_name): # Identifies week columns like '36/2024'
                        try:
                            week_num = int(col_name.split('/')[0])
                            if week_num >= start_week and week_num in week_cols:
                                target_col_idx = week_cols[week_num]
                                quantity = source_row[col_name]
                                ws.cell(row=target_row_idx, column=target_col_idx).value = quantity
                        except (ValueError, IndexError):
                            continue # Ignore columns that are not in the 'WW/YYYY' format

        # 6. Save the updated workbook
        wb.save(forecast_file_path)
        print(f"Successfully updated and saved '{forecast_file_path}'.")

    except Exception as e:
        print(f"An unexpected error occurred during the forecast update: {e}")

if __name__ == "__main__":
    input_directory = os.path.dirname(os.path.abspath(__file__))
    timestamp = datetime.now().strftime("%y%m%d_%H%M")
    output_excel_file = f"mb_extracted_data_{timestamp}.xlsx"
    mercedes_excel_file = "Mercedes_Shipping_Plan_EDI.xlsx"
    process_mb_files(input_directory, output_excel_file, mercedes_excel_file)

    # --- New Step: Update the master forecast file ---
    forecast_master_file = "forecast for all projects.xlsx"
    forecast_file_path = os.path.join(input_directory, forecast_master_file)
    update_forecast_file(output_excel_file, forecast_file_path)
