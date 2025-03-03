# Access Review Scripts

This project contains scripts to perform access reviews using data from Excel files and store the results in a database.

## Configuration

The project uses a configuration file to set various parameters. An example configuration file is provided as `example-config.json`. You can copy this file to `config.json` and modify it as needed.

### Configuration Options

- `log_file`: The file where logs will be written.
- `log_level`: The logging level (e.g., `ERROR`, `INFO`).
- `log_format`: The format for log messages.
- `excluded_sheets`: Sheets in the Excel file to exclude from processing.
- `review_columns`: Columns to review in the Excel file.
- `review_values`: Values to look for in the review columns.
- `admin_roles`: Roles considered as admin roles.
- `mismatch_values`: Values considered as mismatches.
- `login_columns`: Columns that contain login information.
- `database_file`: The SQLite database file to store the results.

## Usage

1. **Install Dependencies**: Ensure you have the necessary Python packages installed. You can use `pip` to install any required packages.

    ```sh
    pip install -r requirements.txt
    ```

2. **Configure the Project**: Copy [example-config.json](http://_vscodecontentref_/6) to [config.json](http://_vscodecontentref_/7) and modify it according to your needs.

    ```sh
    cp example-config.json config.json
    ```

3. **Run the Main Script**: Execute the main script to perform the access review.

    ```sh
    python main.py
    ```

## Project Structure

## Scripts

- [audit_operations.py](http://_vscodecontentref_/8): Contains functions related to auditing operations.
- [db_operations.py](http://_vscodecontentref_/9): Contains functions to interact with the database.
- [file_io_operations.py](http://_vscodecontentref_/10): Contains functions for file input/output operations.
- [main.py](http://_vscodecontentref_/11): The main script to run the access review process.

## Data

Place your Excel files in the [data](http://_vscodecontentref_/12) directory. The script will process these files based on the configuration settings.

## Logging

Logs will be written to the file specified in the `log_file` configuration option. The log level and format can also be configured.

## License

This project is licensed under the MIT License.
