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
    cursor.execute("SELECT DISTINCT service_name, field_name FROM audit")
    service_columns = cursor.fetchall()
    for service_name, field_name in service_columns:
        if field_name in config['review_columns']:  # Only add columns in review_columns
            headers.append(f"{service_name} - {field_name}")

    if verbosity > 0:
        print(f"Headers: {headers}")
    rollup_sheet.append(headers)

    # Write audit data to Rollup sheet
    cursor.execute('SELECT username, service_name, field_name, field_value, needs_review, admin_privileges_review, comments FROM audit')
    audit_data = cursor.fetchall()

    user_data = {}
    for username, service_name, field_name, field_value, needs_review, admin_privileges_review, comments in audit_data:
        if username not in user_data:
            user_data[username] = {
                "services": {},
                "comments": set(),  # Use a set to avoid duplicate comments
                "needs_review": needs_review,
                "admin_privileges_review": admin_privileges_review
            }
        if service_name not in user_data[username]["services"]:
            user_data[username]["services"][service_name] = {}
        user_data[username]["services"][service_name][field_name] = field_value
        if needs_review:
            user_data[username]["comments"].add(comments)

    for username, data in user_data.items():
        cursor.execute('SELECT mail, display_name, created_from FROM users WHERE mail = ?', (username,))
        user_info = cursor.fetchone()
        if user_info:
            user_email, display_name, created_from = user_info
        else:
            user_email, display_name, created_from = username, username, "N/A"

        user_row = [
            display_name,
            user_email,
            created_from,
            "TRUE" if data["needs_review"] else "",
            "TRUE" if data["admin_privileges_review"] else "",
            "; ".join(data["comments"])
        ]
        for service_name, field_name in service_columns:
            if field_name in config['review_columns']:  # Only process columns in review_columns
                value = data["services"].get(service_name, {}).get(field_name, "N/A")
                user_row.append(value)

        if verbosity > 0 or debug:
            print(f"Appending user row to Rollup sheet: {user_row}")
            logging.debug(f"Appending user row to Rollup sheet: {user_row}")

        rollup_sheet.append(user_row)

    workbook.save(file_path)
    print("Audit data written to Rollup sheet successfully.")
    if debug:
        logging.debug("Audit data written to Rollup sheet successfully.")