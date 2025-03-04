import openpyxl
import sqlite3
import json
import logging

# Load configuration
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

# Configure logging
logging.basicConfig(filename=config['log_file'], level=getattr(logging, config['log_level']), format=config['log_format'])

def perform_audit(file_path, db_conn, verbosity, debug):
    print(f"Performing audit on workbook: {file_path}")
    try:
        workbook = openpyxl.load_workbook(file_path)
    except PermissionError:
        print(f"Error: The file {file_path} is open. Please close the file and try again.")
        return

    # Configure logging
    logging.basicConfig(filename=config['log_file'], level=config['log_level'], format=config['log_format'])
    if debug:
        logging.getLogger().setLevel(logging.DEBUG)

    # Create audit table
    cursor = db_conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS audit (
                        username TEXT,
                        service_name TEXT,
                        field_name TEXT,
                        field_value TEXT,
                        needs_review BOOLEAN,
                        admin_privileges_review BOOLEAN,
                        comments TEXT,
                        admin_comments TEXT,
                        FOREIGN KEY(username) REFERENCES users(mail)
                      )''')
    db_conn.commit()
    print("Table 'audit' created or already exists.")

    # Add user data
    cursor.execute('SELECT * FROM users')
    users = cursor.fetchall()
    audit_data = {}

    for user in users:
        user_name = user[2]  # Correctly map the display name
        user_email = user[1].lower() if user[1] else None  # Convert email to lowercase for case-insensitive comparison
        created_from = user[4]
        audit_data[user_name] = {
            "email": user_email,
            "services": {},
            "needs_review": False,
            "admin_privileges_review": False,
            "created_from": created_from,
            "comments": set(),  # Use a set to avoid duplicate comments
            "admin_comments": set()  # Use a set to avoid duplicate admin comments
        }

        if verbosity > 0 or debug:
            print(f"Processing user: {user_name} (Email: {user_email})")
            logging.debug(f"Processing user: {user_name} (Email: {user_email})")

        match_count = 0  # Reset match_count for each user
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        for table in tables:
            table_name = table[0]
            if table_name not in config['excluded_sheets']:
                cursor.execute(f'PRAGMA table_info("{table_name}")')
                columns = cursor.fetchall()
                login_column = next((column[1] for column in columns if column[1].lower() in config['login_columns']), None)

                if login_column:
                    try:
                        cursor.execute(f'SELECT * FROM "{table_name}" WHERE LOWER({login_column}) = LOWER(?)', (user_email,))
                        service_data = cursor.fetchone()
                        if service_data:
                            for column_name, value in zip([col[1] for col in columns], service_data):
                                if column_name in config['review_columns']:  # Only process columns in review_columns
                                    if table_name not in audit_data[user_name]["services"]:
                                        audit_data[user_name]["services"][table_name] = {}
                                    audit_data[user_name]["services"][table_name][column_name] = value
                                    if value in config['review_values']:
                                        match_count += 1
                                        audit_data[user_name]["comments"].add(f"{table_name} - {column_name}: {value}")
                                        if verbosity > 0 or debug:
                                            print(f"Match from review variables for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                            logging.debug(f"Review needed for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                    if any(role in value.lower() for role in config['admin_roles']):
                                        audit_data[user_name]["admin_privileges_review"] = True
                                        audit_data[user_name]["admin_comments"].add(f"{table_name}")
                                        if verbosity > 0 or debug:
                                            print(f"Admin role found for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                            logging.debug(f"Admin role found for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                    # Insert into audit table
                                    cursor.execute('''INSERT INTO audit (username, service_name, field_name, field_value, needs_review, admin_privileges_review, comments, admin_comments)
                                                      VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', (user_email, table_name, column_name, value, match_count >= 2, audit_data[user_name]["admin_privileges_review"], "; ".join(audit_data[user_name]["comments"]), "; ".join(audit_data[user_name]["admin_comments"])))
                                    if debug:
                                        logging.debug(f"Inserted into audit table: {user_email, table_name, column_name, value, match_count >= 2, audit_data[user_name]["admin_privileges_review"], "; ".join(audit_data[user_name]["comments"]), "; ".join(audit_data[user_name]["admin_comments"])}")
                        else:
                            if table_name not in audit_data[user_name]["services"]:
                                audit_data[user_name]["services"][table_name] = {}
                            audit_data[user_name]["services"][table_name][column_name] = "N/A"
                            if debug:
                                logging.debug(f"No service data found for user: {user_name}, table: {table_name}, column: {column_name}")
                    except sqlite3.OperationalError as e:
                        logging.error(f"Error querying {table_name} for column {column_name}: {e}")
                        if table_name not in audit_data[user_name]["services"]:
                            audit_data[user_name]["services"][table_name] = {}
                        audit_data[user_name]["services"][table_name][column_name] = "N/A"
                        if debug:
                            logging.debug(f"OperationalError for user: {user_name}, table: {table_name}, column: {column_name}")

        if match_count >= 2:
            audit_data[user_name]["needs_review"] = True
            if debug:
                logging.debug(f"User {user_name} marked for review with match_count: {match_count}")

    db_conn.commit()
    print("Audit completed and saved to database.")
    if debug:
        logging.debug("Audit completed and saved to database.")