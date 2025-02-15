import pandas as pd
from datetime import datetime, timedelta
from ics import Calendar, Event
import tempfile
import logging

# Configure logging
logger = logging.getLogger(__name__)

SHIFT_MAP = {
    "IV": "0600-1400",
    "A": "0700-1500", "BH": "0700-1500", "C": "0700-1500",
    "D": "0700-1500", "HDmix": "0700-1500",
    "W": "0700-1500", "R": "0700-1500", "B": "0700-1500", "F": "0700-1500",
    "G": "0700-1500", "YC": "0700-1500",
    "2ed": "0800-1600",
    "CF": "0800-1600",
    "6": "0900-1700",
    "9": "0900-2100",
    "E1": "1300-2100",
    "E": "1500-2300", "EC": "1500-2300", "EIV": "1500-2300",
    "ED": "1600-0000",
    "N": "2100-0700",
    "13": "2300-0700",
    "5": "0700-1700",
    "7": "0700-1900",
    "IP": "0800-1600",
    "IH": "0800-1600",
    "T": "0800-1400",
    "V": "OFF",
    "-": "OFF",
    "CL": "0800-1600",
    "HD": "0800-1600",
    "IM": "0800-1400",
    "PJ": "0700-1300"
}

def parse_time(time_str: str, date_obj: datetime):
    """Parse time string in format 'HHMM' and combine with date_obj."""
    if time_str == "OFF":
        return None

    try:
        hours = int(time_str[:2])
        minutes = int(time_str[2:])
        return date_obj.replace(hour=hours, minute=minutes)
    except Exception as e:
        logger.error(f"Error parsing time {time_str}: {str(e)}")
        return None

def process_excel_file(file_path: str):
    """Process the Excel file and return the list of employees and start date."""
    logger.info(f"Processing Excel file: {file_path}")

    try:
        # Read Excel file with date parsing
        df = pd.read_excel(file_path, parse_dates=True)
        logger.info(f"Excel file loaded successfully. Shape: {df.shape}")
        logger.info(f"Columns: {df.columns.tolist()}")

        # Try to identify date columns by checking each column
        date_columns = []
        current_year = datetime.now().year
        for col in df.columns:
            # Check if column name can be parsed as a date
            try:
                if isinstance(col, (datetime, pd.Timestamp)):
                    # Ensure the year is set to the current year if not specified
                    if col.year != current_year:
                        col = col.replace(year=current_year)
                    date_columns.append(col)
                elif isinstance(col, str):
                    try:
                        date = pd.to_datetime(col)
                        if date.year != current_year:
                            date = date.replace(year=current_year)
                        date_columns.append(date)
                    except:
                        pass
            except (ValueError, TypeError):
                continue

        logger.info(f"Found {len(date_columns)} date columns: {date_columns}")

        if not date_columns:
            # If no date columns found in headers, try first row
            first_row = df.iloc[0]
            for col, val in first_row.items():
                try:
                    if isinstance(val, (datetime, pd.Timestamp)):
                        if val.year != current_year:
                            val = val.replace(year=current_year)
                        date_columns.append(val)
                    elif isinstance(val, str):
                        try:
                            date = pd.to_datetime(val)
                            if date.year != current_year:
                                date = date.replace(year=current_year)
                            date_columns.append(date)
                        except:
                            pass
                except (ValueError, TypeError):
                    continue

            if date_columns:
                logger.info("Found dates in first row, adjusting dataframe")
                df.columns = df.iloc[0]
                df = df.iloc[1:]

        if not date_columns:
            raise ValueError("No date columns found in the Excel file. Please ensure the file contains dates in either the header row or first data row.")

        date_columns = sorted(date_columns)
        start_date = min(date_columns)
        logger.info(f"Start date identified: {start_date}")

        # Find employees (looking in non-date columns)
        employees = []
        for idx in range(len(df)):
            row = df.iloc[idx]
            for col in df.columns:
                if col not in date_columns:
                    val = str(row[col]).strip()
                    if val and not val.isupper() and val not in SHIFT_MAP:
                        employees.append(val)

        employees = sorted(list(set(employees)))
        logger.info(f"Found {len(employees)} employees: {employees}")

        return employees, start_date

    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}", exc_info=True)
        raise

def generate_ics_file(file_path: str, employee: str):
    """Generate ICS file for an employee's schedule."""
    logger.info(f"Generating calendar for employee: {employee}")

    try:
        # Read Excel file with date parsing
        df = pd.read_excel(file_path, parse_dates=True)
        logger.info("Excel file loaded for calendar generation")

        # Find employee's row and handle first row dates if needed
        employee_row = None
        date_columns = []
        current_year = datetime.now().year

        # Check if dates are in the first row
        first_row = df.iloc[0]
        dates_in_first_row = False
        for col, val in first_row.items():
            try:
                if isinstance(val, (datetime, pd.Timestamp)):
                    dates_in_first_row = True
                    break
                elif isinstance(val, str):
                    pd.to_datetime(val)
                    dates_in_first_row = True
                    break
            except:
                continue

        if dates_in_first_row:
            df.columns = df.iloc[0]
            df = df.iloc[1:]

        # Find employee's row
        for idx in range(len(df)):
            row_values = [str(val).strip() for val in df.iloc[idx].values]
            if employee in row_values:
                employee_row = df.iloc[idx]
                break

        if employee_row is None:
            raise ValueError(f"Could not find schedule for employee '{employee}'")

        # Find date columns
        for col in df.columns:
            try:
                if isinstance(col, (datetime, pd.Timestamp)):
                    # Ensure correct year
                    if col.year != current_year:
                        col = col.replace(year=current_year)
                    date_columns.append(col)
                elif isinstance(col, str):
                    try:
                        date = pd.to_datetime(col)
                        if date.year != current_year:
                            date = date.replace(year=current_year)
                        date_columns.append(date)
                    except:
                        pass
            except:
                continue

        date_columns.sort()
        logger.info(f"Processing {len(date_columns)} date columns")

        if not date_columns:
            raise ValueError("No date columns found in the Excel file")

        cal = Calendar()

        # Process each date column
        for date_col in date_columns:
            # Ensure correct year for the date
            if date_col.year != current_year:
                date_col = date_col.replace(year=current_year)

            shift_code = str(employee_row[date_col]).strip().upper()
            logger.debug(f"Processing date {date_col}: shift code '{shift_code}'")

            if not shift_code or pd.isna(shift_code):
                continue

            if shift_code in SHIFT_MAP:
                shift_times = SHIFT_MAP[shift_code]
                logger.debug(f"Shift times for {shift_code}: {shift_times}")

                if shift_times == "OFF":
                    event = Event()
                    event.name = "OFF"
                    event.begin = date_col
                    event.make_all_day()
                    cal.events.add(event)
                    logger.debug(f"Added OFF day event for {date_col}")
                else:
                    try:
                        start_str, end_str = shift_times.split("-")
                        event = Event()
                        event.name = f"Work Shift: {shift_code}"

                        start_time = parse_time(start_str, date_col)
                        start_hour = int(start_str[:2])
                        end_hour = int(end_str[:2])

                        if end_hour < start_hour:
                            end_time = parse_time(end_str, date_col + timedelta(days=1))
                        else:
                            end_time = parse_time(end_str, date_col)

                        if start_time and end_time:
                            event.begin = start_time
                            event.end = end_time
                            cal.events.add(event)
                            logger.debug(f"Added work shift event: {shift_code} from {start_time} to {end_time}")
                    except Exception as e:
                        logger.error(f"Error processing shift {shift_code}: {str(e)}")
            else:
                event = Event()
                event.name = f"Unknown Shift: {shift_code}"
                event.begin = date_col
                event.make_all_day()
                cal.events.add(event)
                logger.debug(f"Added unknown shift event for {date_col}: {shift_code}")

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.ics')
        temp_file.close()  # Close the file before writing

        with open(temp_file.name, 'w') as f:
            f.writelines(cal.serialize_iter())  # Use serialize_iter instead of str()

        logger.info(f"Calendar file generated: {temp_file.name}")
        return temp_file.name

    except Exception as e:
        logger.error(f"Error generating calendar: {str(e)}", exc_info=True)
        raise