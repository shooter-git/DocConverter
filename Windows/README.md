# File Converter Tool for Windows

This tool converts `.xlsx` files to `.csv` and `.docx` files to `.md` (Markdown).

## Requirements
- Windows 10 or later
- Python 3.x installed
- PowerShell enabled for script execution

## Setup Instructions

1. Clone or download the project files.
2. Open **PowerShell** or **Command Prompt** with administrative privileges.
3. Navigate to the project directory.
4. Run the following command to execute the batch script:
   ```powershell
   .\setup.bat
   ```

The script will:
- Create a virtual environment.
- Upgrade `pip` to the latest version using the Python executable from the virtual environment.
- Install the required dependencies from `requirements.txt`.
- Launch the file conversion tool.

## Using the Tool

- The tool allows batch conversion of multiple files or entire directories.
- You can select individual `.xlsx` and `.docx` files or choose a directory that contains both.
- After conversion, all files will be saved in the `Converted_Batch_Documents` folder created in the working directory.
- A single confirmation dialog will appear once all conversions are complete.

## Notes
- This version of the script does **not** extract or handle embedded images in `.docx` files.
- Ensure you have sufficient permissions to run the script.