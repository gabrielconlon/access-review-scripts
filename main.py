import argparse
import sqlite3
import json
import logging
from db_operations import process_workbook
from audit_operations import perform_audit
from file_io_operations import write_audit_to_rollup

# Load configuration
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

# Configure logging
logging.basicConfig(filename=config['log_file'], level=getattr(logging, config['log_level']), format=config['log_format'])

def main():
    parser = argparse.ArgumentParser(description="Access Review Script")
    parser.add_argument("action", choices=["update_db", "print_user", "print_schema", "list_services", "run_query", "perform_audit", "write_audit"], help="Action to perform. Choices are: update_db, print_user, print_schema, list_services, run_query, perform_audit, write_audit")
    parser.add_argument("-f", "--file", help="Path to the Excel file for updating the database or writing the audit data")
    parser.add_argument("-e", "--email", help="Email of the user to print information")
    parser.add_argument("-o", "--output", help="Path to the output file for saving results")
    parser.add_argument("-q", "--query", help="SQL query to run")
    parser.add_argument("-v", "--verbosity", type=int, default=0, help="Verbosity level (0, 1, 2)")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode to include source information and N/A for empty fields")

    args = parser.parse_args()

    db_conn = sqlite3.connect(config['database_file'])

    if args.action == "update_db":
        if not args.file:
            print("Error: --file argument is required for updating the database")
            return
        process_workbook(args.file, db_conn, args.verbosity)
    elif args.action == "print_user":
        if not args.email:
            print("Error: --email argument is required for printing user information")
            return
        print_user(db_conn, args.email, args.output)
    elif args.action == "print_schema":
        print_table_schema(db_conn, args.output)
    elif args.action == "list_services":
        list_services(db_conn, args.output)
    elif args.action == "run_query":
        if not args.query:
            print("Error: --query argument is required for running a SQL query")
            return
        run_query(db_conn, args.query, args.output)
    elif args.action == "perform_audit":
        if not args.file:
            print("Error: --file argument is required for performing the audit")
            return
        perform_audit(args.file, db_conn, args.verbosity, args.debug)
    elif args.action == "write_audit":
        if not args.file:
            print("Error: --file argument is required for writing the audit data to the Rollup sheet")
            return
        write_audit_to_rollup(args.file, db_conn, args.verbosity, args.debug)

    db_conn.close()

if __name__ == "__main__":
    main()