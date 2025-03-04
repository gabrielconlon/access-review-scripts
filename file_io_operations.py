import openpyxl
import sqlite3
import json
import logging

# Load configuration
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

# Configure logging
logging.basicConfig(filename=config['log_file'], level=getattr(logging, config['log_level']), format=config['log_format'])

def write_audit_to_rollup(file_path, db_conn, verbosity, debug):
    print(f"Writing audit data to Rollup sheet in workbook: {file_path}")
    try:
        workbook = openpyxl.load_workbook(file_path)
    except PermissionError:
        print(f"Error: The file {file_path} is open. Please close the file and try again.")
        return

    rollup_sheet = workbook["Rollup"]

    # Clear existing data in Rollup sheet
    rollup_sheet.delete_rows(2, rollup_sheet.max_row)

    # Add headers
    headers = ["User", "Email", "Created From", "Needs Review", "Admin Privileges Review", "Comments"]
    cursor = db_conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    service_columns = []
    for table in tables:
        table_name = table[0]
        if table_name not in config['excluded_sheets']:
            cursor.execute(f'PRAGMA table_info("{table_name}")')
            columns = cursor.fetchall()
            for column in columns:
                if column[1] in config['review_columns']:
                    service_columns.append((table_name, column[1]))
                    headers.append(f"{table_name} - {column[1]}")

    if verbosity > 0:
        print(f"Headers: {headers}")
    rollup_sheet.append(headers)

    # Write audit data to Rollup sheet
    cursor.execute('SELECT mail, display_name, created_from, needs_review, is_admin, comments FROM users')
    user_data = cursor.fetchall()

    for user in user_data:
        user_email, display_name, created_from, needs_review, is_admin, comments = user

        user_row = [
            display_name,
            user_email,
            created_from,
            "TRUE" if needs_review else "",
            "TRUE" if is_admin else "",
            comments
        ]
        for service_name, field_name in service_columns:
            cursor.execute(f'PRAGMA table_info("{service_name}")')
            columns = cursor.fetchall()
            login_column = next((col for col in config['login_columns'] if col in [desc[1] for desc in columns]), None)
            if login_column:
                try:
                    cursor.execute(f'SELECT {field_name} FROM "{service_name}" WHERE LOWER({login_column}) = LOWER(?)', (user_email,))
                    value = cursor.fetchone()
                    user_row.append(value[0] if value else "N/A")
                except sqlite3.OperationalError as e:
                    user_row.append("N/A")
                    if verbosity > 0 or debug:
                        print(f"Error querying {service_name} for column {field_name}: {e}")
                        logging.debug(f"Error querying {service_name} for column {field_name}: {e}")
            else:
                user_row.append("N/A")

        if verbosity > 0 or debug:
            print(f"Appending user row to Rollup sheet: {user_row}")
            logging.debug(f"Appending user row to Rollup sheet: {user_row}")

        rollup_sheet.append(user_row)

    workbook.save(file_path)
    print("Audit data written to Rollup sheet successfully.")
    if debug:
        logging.debug("Audit data written to Rollup sheet successfully.")