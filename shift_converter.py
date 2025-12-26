import pandas as pd
from datetime import datetime, timedelta
from ics import Calendar, Event
import tempfile
import logging
from zoneinfo import ZoneInfo
import time # Add timing for debugging

# Configure logging with more detail
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Set timezone to Eastern Time
LOCAL_TZ = ZoneInfo("America/New_York")

# NOTE: Consider loading this from a config file or environment variables
SHIFT_MAP = {
    "IV": "0600-1400",
    "A": "0700-1500", "BH": "0700-1500", "C": "0700-1500",
    "D": "0700-1500", "HDmix": "0700-1500",
    "W": "0700-1500", "R": "0700-1500", "B": "0700-1500", "F": "0700-1500",
    "G": "0700-1500", "YC": "0700-1500",
    "2ed": "0800-1600",
    "CF": "0900-1700", "CF*": "0900-1700", # Added CF* based on log
    "6": "0900-1700", "6FT": "0900-1700", # Added 6FT based on log
    "9": "0900-2100", "9-5FT": "0900-1700", # Added 9-5Ft based on log (assuming 9-5)
    "E1": "1300-2100",
    "E": "1500-2300", "EC": "1500-2300", "EIV": "1500-2300", "ECT": "1500-2300", # Added ECT based on log
    "ED": "1600-0000", "EDT": "1600-0000", # Added EDT based on log
    "N": "2100-0700",
    "13": "2300-0700",
    "5": "0700-1700",
    "7": "0700-1900",
    "IP": "0800-1600",
    "IH": "0800-1600",
    "T": "0800-1400",
    "V": "OFF",
    "-": "OFF",
    "CT": "OFF",
    "PL": "OFF",
    "S": "OFF",
    "CL": "0800-1600",
    "HD": "0715-1515",
    "IM": "0800-1400",
    "PJ": "0700-1300",
    # Added based on logs (assuming they are shifts or need mapping)
    "BHt": "0700-1500", # Assuming similar to BH
    "Ft": "0700-1500",  # Assuming Full Time Day? Needs verification
    "Bt": "0700-1500",   # Needs verification
}


def parse_time(time_str: str, date_obj: datetime):
    """Parse time string in format 'HHMM' and combine with date_obj."""
    if time_str == "OFF":
        return None

    try:
        hours = int(time_str[:2]) # This should be line 61 or near it
        minutes = int(time_str[2:])
        # Create a timezone-aware datetime using the LOCAL_TZ defined globally
        aware_date = date_obj.replace(
            hour=hours,
            minute=minutes,
            second=0, # Ensure seconds are zero
            microsecond=0, # Ensure microseconds are zero
            tzinfo=LOCAL_TZ # Apply timezone
        )
        return aware_date
    except Exception as e:
        logger.error(f"Error parsing time string '{time_str}' for date {date_obj}: {str(e)}")
        return None

def _get_schedule_year(df_raw):
    """
    Get the schedule year from cell A1 (top-left).
    This year represents the PRIMARY year (usually January+).
    December dates should use year-1.
    """
    try:
        # Cell A1 = iloc[0, 0] in raw (headerless) dataframe
        year_val = df_raw.iloc[0, 0]
        logger.info(f"Cell A1 raw value: {year_val} (type: {type(year_val)})")

        if pd.notna(year_val):
            # Handle numeric year
            if isinstance(year_val, (int, float)):
                year_int = int(year_val)
                if 2020 <= year_int <= 2050:
                    logger.info(f"Schedule year from A1: {year_int}")
                    return year_int
            # Handle string year
            elif isinstance(year_val, str):
                # Try to extract a 4-digit year from the string
                import re
                match = re.search(r'(20[2-5]\d)', year_val)
                if match:
                    year_int = int(match.group(1))
                    logger.info(f"Schedule year extracted from A1 string: {year_int}")
                    return year_int
    except Exception as e:
        logger.error(f"Error reading year from A1: {e}")

    # Fallback to next year
    fallback = datetime.now().year + 1
    logger.warning(f"Could not read year from A1. Using fallback: {fallback}")
    return fallback


def _adjust_date_with_year(date_obj, schedule_year, first_month=None):
    """
    Simple year adjustment:
    - A1 contains the PRIMARY year (January's year)
    - December = A1 - 1 (previous year)
    - January-November = A1

    Example: A1=2026 means Dec 2025, Jan-Nov 2026
    """
    if date_obj is None:
        return None

    if date_obj.month == 12:
        correct_year = schedule_year - 1  # December is previous year
    else:
        correct_year = schedule_year  # Jan-Nov is the A1 year

    if date_obj.year != correct_year:
        logger.debug(f"Adjusting {date_obj.strftime('%b %d')} from {date_obj.year} to {correct_year}")
        return date_obj.replace(year=correct_year)

    return date_obj


def process_excel_file(file_path: str, timeout=30):
    """Process the Excel file and return the list of employees and start date."""
    start_time_process = time.time() # Use unique name for this timer
    logger.info(f"Processing Excel file: {file_path}")

    try:
        # Read Excel file
        try:
            df_raw = pd.read_excel(file_path, header=None) # Read raw first
            logger.debug(f"Raw Excel data (no header):\n{df_raw.head()}")
            df = pd.read_excel(file_path, parse_dates=True) # Read again potentially with headers
        except Exception as read_err:
             logger.error(f"Pandas read_excel failed: {read_err}")
             raise

        logger.info(f"Excel file loaded successfully. Shape: {df.shape}")
        logger.info(f"Columns: {df.columns.tolist()}")
        logger.debug(f"First 5 rows of data:\n{df.head(5)}")

        # *** Get Schedule Year ***
        schedule_year = _get_schedule_year(df_raw) # Use raw read to get year before potential header shift

        if time.time() - start_time_process > timeout: raise TimeoutError("Timeout after loading Excel")

        # STEP 1: Collect raw dates first (to find first month)
        raw_dates = []  # List of (column_key, date_obj) tuples
        dates_in_header = False

        # Check column headers for dates
        logger.debug(f"Checking column headers for dates")
        for col in df.columns:
            if time.time() - start_time_process > timeout: raise TimeoutError("Timeout processing columns")
            try:
                date_obj = None
                if isinstance(col, (datetime, pd.Timestamp)): date_obj = col
                elif isinstance(col, str):
                    try: date_obj = pd.to_datetime(col)
                    except: continue
                else: continue

                if date_obj:
                    raw_dates.append((col, date_obj))
                    dates_in_header = True
            except Exception as e: logger.debug(f"Error checking header '{col}': {e}")

        # If no dates in headers, check first row
        if not dates_in_header and len(df) > 0:
            logger.debug(f"No dates in headers, checking first row")
            first_row = df.iloc[0]
            for col_idx, (orig_col_name, val) in enumerate(first_row.items()):
                if time.time() - start_time_process > timeout: raise TimeoutError("Timeout processing first row")
                try:
                    date_obj = None
                    if isinstance(val, (datetime, pd.Timestamp)): date_obj = val
                    elif isinstance(val, str):
                        try: date_obj = pd.to_datetime(val)
                        except: continue
                    else: continue

                    if date_obj:
                        raw_dates.append((orig_col_name, date_obj))
                except Exception as e: logger.debug(f"Error checking row 0 val '{val}': {e}")

        # STEP 2: Find the first month in the schedule (first column's month)
        first_month = None
        if raw_dates:
            first_month = raw_dates[0][1].month
            logger.info(f"First month in schedule: {first_month} (1=Jan, 12=Dec)")

        # STEP 3: Now adjust years and build final date structures
        date_columns = []
        date_col_map = {}

        for col_key, date_obj in raw_dates:
            # Adjust year based on first month context
            date_obj = _adjust_date_with_year(date_obj, schedule_year, first_month)
            date_obj_aware = date_obj.replace(tzinfo=LOCAL_TZ)

            if date_obj_aware not in date_col_map.values():
                date_columns.append(date_obj_aware)
                date_col_map[col_key] = date_obj_aware
                logger.debug(f"Date column: {col_key} -> {date_obj_aware}")

        if not dates_in_header and date_columns:
            logger.info("Found dates in first row, data starts from second row.")

        if not date_columns: raise ValueError("No valid date columns identified.")

        date_columns = sorted(date_columns)
        start_date = min(date_columns) if date_columns else None
        logger.info(f"Start date identified: {start_date}")
        logger.info(f"All unique dates found: {date_columns}")
        logger.debug(f"Date column mapping (orig_col -> datetime): {date_col_map}")

        # --- Find employees ---
        employees = set()
        potential_employee_cols = []
        date_col_keys = set(date_col_map.keys())
        for col in df.columns:
            if col not in date_col_keys: potential_employee_cols.append(col)
            if len(potential_employee_cols) >= 2: break
        logger.debug(f"Potential employee name columns: {potential_employee_cols}")
        if not potential_employee_cols: potential_employee_cols = [col for col in df.columns if col not in date_col_keys]

        employee_search_limit = min(50, len(df))
        logger.debug(f"Searching for employees in first {employee_search_limit} rows within columns: {potential_employee_cols}")

        # *** Use the initialized dates_in_header variable ***
        df_data_start_row = 1 if (not dates_in_header and date_columns) else 0
        logger.debug(f"Employee search starting from data row index: {df_data_start_row}")

        # Words that indicate headers, not employee names
        HEADER_WORDS = {'shift', 'code', 'codes', 'hour', 'hours', 'date', 'name', 'employee', 'schedule', 'time', 'day', 'week', 'nan', 'none', 'null', ''}

        for idx in range(df_data_start_row, employee_search_limit):
            if time.time() - start_time_process > timeout: raise TimeoutError("Timeout finding employees")
            if idx >= len(df): break # Ensure index is within bounds
            row = df.iloc[idx]
            for emp_col in potential_employee_cols:
                 if emp_col not in row: continue
                 val = str(row[emp_col]).strip()
                 val_lower = val.lower()

                 # Skip if empty, nan, or contains header-like words
                 if not val or val_lower in HEADER_WORDS or 'nan' in val_lower:
                     continue

                 # Skip if any word in the value is a header word
                 val_words = set(val_lower.split())
                 if val_words & HEADER_WORDS:
                     continue

                 is_potential_name = (
                     any(c.islower() for c in val) and
                     not val.isupper() and not val.isdigit() and
                     val.upper() not in SHIFT_MAP and
                     len(val) >= 2  # Names should be at least 2 chars
                 )
                 if is_potential_name:
                    employees.add(val)
                    logger.debug(f"Found potential employee in row {idx}, col '{emp_col}': {val}")

        employees = sorted(list(employees))
        logger.info(f"Found {len(employees)} potential employees: {employees}")
        if not employees: logger.warning("No potential employees found based on heuristics.")

        return employees, start_date

    except Exception as e:
        elapsed = time.time() - start_time_process
        logger.error(f"Error processing Excel file after {elapsed:.2f}s: {str(e)}", exc_info=True)
        raise

def generate_ics_file(file_path: str, employee: str, shift_map=None, timeout=60):
    """Generate ICS file for an employee's schedule."""
    func_start_time = time.time() # Use unique name for the main timer
    logger.info(f"Generating calendar for employee: {employee}")
    current_shift_map = shift_map if shift_map is not None else SHIFT_MAP
    logger.debug(f"Using shift map: {current_shift_map}")

    try:
        # Read Excel file
        try:
            df_raw = pd.read_excel(file_path, header=None) # Read raw first
            logger.debug(f"Raw Excel data (no header) for ICS:\n{df_raw.head()}")
            df = pd.read_excel(file_path, parse_dates=True) # Read again
        except Exception as read_err:
             logger.error(f"Pandas read_excel failed during ICS generation: {read_err}")
             raise
        logger.info("Excel file re-loaded for calendar generation")

        # *** Get Schedule Year ***
        schedule_year = _get_schedule_year(df_raw) # Use raw read

        if time.time() - func_start_time > timeout: raise TimeoutError("Timeout after loading Excel for ICS")

        # STEP 1: Collect raw dates first (to find first month)
        raw_dates = []
        dates_in_header = False

        for col_header in df.columns:
            if time.time() - func_start_time > timeout: raise TimeoutError("Timeout finding date headers")
            try:
                date_obj = None
                if isinstance(col_header, (datetime, pd.Timestamp)): date_obj = col_header
                elif isinstance(col_header, str):
                    try: date_obj = pd.to_datetime(col_header)
                    except: continue
                else: continue

                if date_obj:
                    raw_dates.append((col_header, date_obj))
                    dates_in_header = True
            except Exception as e: logger.debug(f"ICS Gen Err checking header {col_header}: {e}")

        if not dates_in_header and len(df) > 0:
            first_row = df.iloc[0]
            for col_idx, (orig_col_name, val) in enumerate(first_row.items()):
                if time.time() - func_start_time > timeout: raise TimeoutError("Timeout finding dates in first row")
                try:
                    date_obj = None
                    if isinstance(val, (datetime, pd.Timestamp)): date_obj = val
                    elif isinstance(val, str):
                        try: date_obj = pd.to_datetime(val)
                        except: continue
                    else: continue

                    if date_obj:
                        raw_dates.append((orig_col_name, date_obj))
                except Exception as e: logger.debug(f"ICS Gen Err checking row 0 val {val}: {e}")

        # STEP 2: Find first month
        first_month = raw_dates[0][1].month if raw_dates else 1
        logger.info(f"ICS Gen: First month in schedule: {first_month}")

        # STEP 3: Adjust years and build date structures
        date_columns = []
        date_col_map = {}

        for col_key, date_obj in raw_dates:
            date_obj = _adjust_date_with_year(date_obj, schedule_year, first_month)
            date_obj_aware = date_obj.replace(tzinfo=LOCAL_TZ)

            if date_obj_aware not in date_col_map.values():
                date_columns.append(date_obj_aware)
                date_col_map[col_key] = date_obj_aware
                logger.debug(f"ICS Gen: Date {col_key} -> {date_obj_aware}")

        # Determine where actual data starts
        if dates_in_header:
            logger.info("ICS Gen: Dates in header. Data starts from row 0.")
            df_data = df
        elif date_columns:
            logger.info("ICS Gen: Dates in first row. Data starts from row 1.")
            df_data = df.iloc[1:].reset_index(drop=True)
        else:
            raise ValueError("ICS Gen: No date columns found.")

        if not date_columns: raise ValueError("ICS Gen: No date columns identified.")

        date_columns = sorted(date_columns)
        logger.info(f"ICS Gen: Found {len(date_columns)} unique dates.")
        logger.debug(f"ICS Gen: Date column mapping (orig_col -> datetime): {date_col_map}")

        # --- Employee Row Identification ---
        employee_row_idx = None
        employee_search_limit = min(50, len(df_data))
        logger.debug(f"ICS Gen: Searching for employee '{employee}' in first {employee_search_limit} rows of data frame with columns: {df_data.columns.tolist()}")

        employee_col_name = None
        date_col_keys = set(date_col_map.keys())
        # Identify potential employee column name from the ORIGINAL df columns
        for col in df.columns:
            if col not in date_col_keys:
                employee_col_name = col
                logger.debug(f"ICS Gen: Identified potential employee name column: '{employee_col_name}'")
                break
        if not employee_col_name: raise ValueError("ICS Gen: Could not identify employee name column.")
        # Check if this column exists in df_data (might be different if dates were in first row)
        if employee_col_name not in df_data.columns:
            try:
                 original_cols = df.columns.tolist()
                 employee_col_idx = original_cols.index(employee_col_name)
                 if employee_col_idx < len(df_data.columns):
                      employee_col_name = df_data.columns[employee_col_idx] # Get the corresponding column name in df_data
                      logger.debug(f"ICS Gen: Adjusted employee column name for df_data: '{employee_col_name}'")
                 else: raise ValueError(f"ICS Gen: Employee column index {employee_col_idx} out of bounds.")
            except (ValueError, IndexError) as e:
                 logger.error(f"ICS Gen: Error mapping employee column name to df_data: {e}")
                 raise ValueError("ICS Gen: Could not map employee name column to data frame.")

        for idx in range(employee_search_limit):
            if time.time() - func_start_time > timeout: raise TimeoutError("Timeout finding employee row")
            if idx >= len(df_data): break # Check bounds
            try:
                 current_name = str(df_data.iloc[idx][employee_col_name]).strip()
                 if current_name == employee:
                    employee_row_idx = idx
                    logger.info(f"ICS Gen: Found employee '{employee}' at data row index {idx}")
                    break
            except KeyError: logger.error(f"ICS Gen: Column '{employee_col_name}' not found in df_data at index {idx}.") ; raise
            except IndexError: logger.error(f"ICS Gen: Index {idx} out of bounds for df_data."); break

        if employee_row_idx is None: raise ValueError(f"ICS Gen: Could not find schedule data for employee '{employee}'")

        # --- Calendar Generation ---
        cal = Calendar()
        processed_count = 0

        for date_obj in date_columns:
             orig_col_for_date = None
             for name, dt in date_col_map.items():
                 if dt.date() == date_obj.date() and dt.tzinfo == date_obj.tzinfo:
                     orig_col_for_date = name
                     break
             if orig_col_for_date is None:
                 logger.warning(f"ICS Gen: Could not find original column name for date {date_obj}. Skipping.")
                 continue
             # Check if this original column name still exists in df_data
             if orig_col_for_date not in df_data.columns:
                  found_matching_col = False
                  for data_col in df_data.columns:
                       try:
                           if isinstance(data_col, (datetime, pd.Timestamp)):
                                data_col_aware = data_col.replace(tzinfo=LOCAL_TZ) if data_col.tzinfo is None else data_col.astimezone(LOCAL_TZ)
                                date_obj_aware = date_obj # Already aware
                                if data_col_aware.date() == date_obj_aware.date():
                                     orig_col_for_date = data_col # Use the actual column name from df_data
                                     found_matching_col = True
                                     break
                       except Exception: continue
                  if not found_matching_col:
                     logger.warning(f"ICS Gen: Orig col '{orig_col_for_date}' missing & couldn't map date {date_obj.date()} to df_data cols. Skipping.")
                     continue

             # *** Variable Renaming and Timeout Check ***
             logger.debug(f"ICS Gen: Checking timeout before processing date {date_obj}. Type of func_start_time: {type(func_start_time)}")
             if time.time() - func_start_time > timeout: # <-- Check the outer timer
                 logger.error(f"Timeout exceeded during event loop: {timeout}s")
                 raise TimeoutError(f"Processing exceeded {timeout} seconds during event loop")

             try:
                 # Use the potentially updated orig_col_for_date if it was remapped above
                 shift_code = str(df_data.iloc[employee_row_idx][orig_col_for_date]).strip().upper()
                 logger.debug(f"Processing date {date_obj}: Employee '{employee}', Data Col '{orig_col_for_date}', Shift code '{shift_code}'")

                 if not shift_code or pd.isna(shift_code) or shift_code.lower() == 'nan':
                     logger.debug(f"Empty or NaN shift code for {date_obj}, skipping.")
                     continue

                 event = Event()

                 if shift_code in current_shift_map:
                     shift_times = current_shift_map[shift_code]
                     logger.debug(f"Shift times for '{shift_code}': {shift_times}")

                     if shift_times == "OFF":
                         event.name = "OFF"
                         event.begin = date_obj
                         event.make_all_day()
                         cal.events.add(event)
                         processed_count += 1
                         logger.debug(f"Added OFF day event for {date_obj.date()}")
                     else:
                         try:
                             start_str, end_str = shift_times.split("-")
                             event.name = f"Work: {shift_code}"

                             event_start_dt = parse_time(start_str, date_obj) # Renamed
                             start_hour = int(start_str[:2])
                             end_hour = int(end_str[:2])

                             if end_hour < start_hour: event_end_dt = parse_time(end_str, date_obj + timedelta(days=1)) # Renamed
                             else: event_end_dt = parse_time(end_str, date_obj) # Renamed

                             if event_start_dt and event_end_dt:
                                 event.begin = event_start_dt
                                 event.end = event_end_dt
                                 cal.events.add(event)
                                 processed_count += 1
                                 logger.debug(f"Added work shift event: {shift_code} from {event_start_dt} to {event_end_dt}")
                             else: logger.warning(f"Could not parse start/end time for shift '{shift_code}' ({shift_times}) on {date_obj.date()}. Skipping.")
                         except Exception as e: logger.error(f"Error processing known shift '{shift_code}' on {date_obj.date()}: {str(e)}")
                 else:
                     event.name = f"Unknown Shift: {shift_code}"
                     event.begin = date_obj
                     event.make_all_day()
                     cal.events.add(event)
                     processed_count += 1
                     logger.warning(f"Added 'Unknown Shift: {shift_code}' all-day event for {date_obj.date()}. Add mapping?")

             except KeyError: logger.error(f"ICS Gen: Column '{orig_col_for_date}' key error for row {employee_row_idx} on {date_obj}. Skipping date.") ; continue
             except IndexError: logger.error(f"ICS Gen: Employee row index {employee_row_idx} out of bounds on {date_obj}. Stopping.") ; break
             except Exception as e: logger.error(f"Error processing data for date {date_obj}: {str(e)}", exc_info=True) ; continue

        logger.info(f"Processed {processed_count} calendar entries for {employee}.")

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.ics', mode='w', encoding='utf-8')
        logger.debug(f"Writing calendar to temp file: {temp_file.name}")
        try:
            for line in cal.serialize_iter(): temp_file.write(line)
            logger.debug("Finished writing lines.")
        finally:
             temp_file.close()
             logger.debug("Temp file closed.")

        # Verify file (optional)
        try:
            with open(temp_file.name, 'r', encoding='utf-8') as f_verify:
                 content_sample = f_verify.read(200)
                 logger.debug(f"ICS file sample:\n{content_sample}...")
                 if not content_sample.startswith("BEGIN:VCALENDAR"): logger.error("ICS does not start with BEGIN:VCALENDAR!")
        except Exception as verify_err: logger.error(f"Could not verify ICS: {verify_err}")

        elapsed = time.time() - func_start_time
        logger.info(f"Calendar file generated successfully in {elapsed:.2f}s: {temp_file.name}")
        return temp_file.name

    except Exception as e:
        elapsed = time.time() - func_start_time
        logger.debug(f"ICS Gen Error Handler: Type of func_start_time: {type(func_start_time)}") # Added debug log
        logger.error(f"Error generating calendar after {elapsed:.2f}s: {str(e)}", exc_info=True)
        raise