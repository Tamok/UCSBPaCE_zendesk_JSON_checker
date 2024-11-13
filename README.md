# Collate and Analyze JSON Files

This script (`main.py`) is designed to process, collate, and analyze JSON files containing sensitive P2-P3 data, generating summary statistics in an Excel file format. Due to the sensitive nature of the data, only `main.py` should be hosted on GitHub, and users must follow these setup instructions to securely run the script locally.

## Prerequisites

To use this script, you'll need:

1. **Python 3.7+** installed.
2. **Dependencies**: Install the required packages using the following command:

   ```bash
   pip install xlsxwriter
   ```

## Setup Instructions

1. **Clone the Repository**: Clone the repository containing the `main.py` file to your local machine:

   ```bash
   git clone <your-repository-url>
   ```

2. **Create a Secure Data Directory**: Place the sensitive JSON data files in a local directory named `json_files` within the project folder. Do not commit these files to version control, as they contain sensitive information.

3. **Configuration**: Ensure that you have the following files and folder structure:

   ```
   your-project-folder/
   |-- main.py
   |-- json_files/
       |-- file1.json
       |-- file2.json
       |-- ...
   ```

## Running the Script

1. **Execute the Script**: Run the script using Python:

   ```bash
   python main.py
   ```

   - **Log File**: A log file named `collate_json_files.log` will be created in the root directory, recording the detailed processing steps and any errors.
   - **Output Files**: The script will generate the following output files:
     - `combined.json`: A combined JSON file containing data from all JSON files in `json_files`.
     - `combined_analysis.xlsx`: An Excel file containing statistical summaries of the data.

## Data Security Considerations

- **Sensitive Data Warning**: The JSON files and generated output contain P2-P3 level sensitive data. **Do not share these files publicly or commit them to version control**.
- **Local Setup**: Run the script locally in a secure environment. Store the JSON and output files in a location that respects the sensitivity of the data.

## Logging

- The script will log its progress to `collate_json_files.log`.
- In the event of errors, the log file will contain detailed information to help diagnose issues.

## Script Workflow Overview

1. **Setup Logging**: The script sets up both file and console logging.
2. **Data Processing**: It reads each JSON file, processes its entries, and extracts key metrics.
3. **Analysis**: Analyzes the data and writes summary statistics to an Excel file with multiple sheets.
4. **Validation**: The script validates data integrity at several steps to ensure accuracy.

## Troubleshooting

- **Invalid File Paths**: Ensure that all paths to the JSON files are correct and that they are properly placed in the `json_files` directory.
- **Dependency Issues**: Ensure you have installed `xlsxwriter` via `pip install xlsxwriter`.
- **Data Errors**: If the script reports invalid data formats or fails validation, check that the input JSON files are correctly formatted and meet the expected data structure.

## License

This project is licensed under the MIT License. Refer to `LICENSE` for more details.

## Disclaimer

The script processes and generates sensitive data. Ensure all local data storage complies with relevant security protocols and privacy policies.
