# DOCX to PDF Converter

This script automates the process of converting Microsoft Word `.docx` files to `.pdf` format. It utilizes multi-threading to speed up the conversion process.

## Dependencies

- `logging`
- `os`
- `docx2pdf`
- `sys`
- `pythoncom`
- `concurrent.futures`
- `termcolor`

You can install the required libraries using pip:

```bash
pip install docx2pdf termcolor

```

## How to Use

1. Navigate to the folder containing the script and your `.docx` files.
2. Run the script by using the command:
    
    ```bash
    python script_name.py [folder_path]
    
    ```
    
    - If `[folder_path]` is not provided, the script will assume the current directory. It will then prompt you to confirm if you wish to proceed with the conversion of the `.docx` files in the current directory.
    - Successfully converted files will be stored in a sub-folder named "executed" within the given folder.

## Logging

- All the events are logged in a file named `docx2pdf.log`.