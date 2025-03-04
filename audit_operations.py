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

    # Check if columns exist and add them if they don't
    cursor = db_conn.cursor()
    cursor.execute("PRAGMA table_info(users)")
    columns = cursor.fetchall()
    column_names = [column[1] for column in columns]

    if 'needs_review' not in column_names:
        cursor.execute('''ALTER TABLE users ADD COLUMN needs_review BOOLEAN DEFAULT 0''')
    if 'is_admin' not in column_names:
        cursor.execute('''ALTER TABLE users ADD COLUMN is_admin BOOLEAN DEFAULT 0''')
    if 'comments' not in column_names:
        cursor.execute('''ALTER TABLE users ADD COLUMN comments TEXT''')
    db_conn.commit()

    # Add user data
    cursor.execute('SELECT * FROM users')
    users = cursor.fetchall()

    for user in users:
        user_name = user[2]  # Correctly map the display name
        user_email = user[1].lower() if user[1] else None  # Convert email to lowercase for case-insensitive comparison
        created_from = user[4]
        needs_review = False
        is_admin = False
        comments = set()  # Use a set to avoid duplicate comments
        admin_comments = set()  # Use a set to avoid duplicate admin comments

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
                login_column = next((col for col in config['login_columns'] if col in [desc[1] for desc in columns]), None)

                if login_column:
                    try:
                        cursor.execute(f'SELECT * FROM "{table_name}" WHERE LOWER({login_column}) = LOWER(?)', (user_email,))
                        service_data = cursor.fetchone()
                        if service_data:
                            for column_name, value in zip([col[1] for col in columns], service_data):
                                if column_name in config['review_columns']:  # Only process columns in review_columns
                                    if value in config['review_values']:
                                        match_count += 1
                                        comments.add(f"{table_name} - {column_name}: {value}")
                                        if verbosity > 0 or debug:
                                            print(f"Match from review variables for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                            logging.debug(f"Review needed for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                    if any(role in value.lower() for role in config['admin_roles']):
                                        is_admin = True
                                        admin_comments.add(f"{table_name}")
                                        if verbosity > 0 or debug:
                                            print(f"Admin role found for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                                            logging.debug(f"Admin role found for user: {user_name}, table: {table_name}, column: {column_name}, value: {value}")
                        else:
                            if debug:
                                logging.debug(f"No service data found for user: {user_name}, table: {table_name}, column: {column_name}")
                    except sqlite3.OperationalError as e:
                        logging.error(f"Error querying {table_name} for column {column_name}: {e}")
                        if debug:
                            logging.debug(f"OperationalError for user: {user_name}, table: {table_name}, column: {column_name}")

        if match_count >= 2:
            needs_review = True
            if debug:
                logging.debug(f"User {user_name} marked for review with match_count: {match_count}")

        # Update users table
        cursor.execute('''UPDATE users SET needs_review = ?, is_admin = ?, comments = ? WHERE mail = ?''',
                       (needs_review, is_admin, "; ".join(comments), user_email))
        db_conn.commit()

    print("Audit completed and saved to database.")
    if debug:
        logging.debug("Audit completed and saved to database.")