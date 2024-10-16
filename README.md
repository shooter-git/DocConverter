# DocConverter Tool

This project provides a tool for converting `.xlsx` files to `.csv` and `.docx` files to `.md` (Markdown). The tool supports batch conversion of files and can be run on both Windows and Linux platforms.

## Folder Structure

- `Linux/`: Contains the setup and scripts for running the tool on Linux.
- `Windows/`: Contains the setup and scripts for running the tool on Windows.

Each platform has its own virtual environment (`env/`) and setup scripts.

## Requirements

- **Windows:** Python 3.x installed, PowerShell enabled for script execution.
- **Linux:** Python 3.x installed, `pip` available.

## Installing Python

### For Windows

1. Open **PowerShell** or **Command Prompt**.
2. Use the Windows Package Manager (`winget`) to install Python:
   ```powershell
   winget install Python.Python.3
   ```
3. After installation, verify Python is installed by running:
   ```powershell
   python --version
   ```

Alternatively, you can download Python from the [official Python website](https://www.python.org/downloads/) and install it manually. During the installation, make sure to check the option **Add Python to PATH**.

### For Linux (Ubuntu)

1. Open a terminal and run the following command to install Python:
   ```bash
   sudo apt update
   sudo apt install python3 python3-venv python3-pip
   ```
2. Verify the installation by running:
   ```bash
   python3 --version
   ```

## Setting Up the Tool

### For Windows

1. Navigate to the `Windows` directory:
   ```powershell
   cd Windows
   ```
2. Open **PowerShell** or **Command Prompt** with administrative privileges.
3. Run the setup script:
   ```powershell
   .\setup.bat
   ```
   - This will create a virtual environment, upgrade `pip`, install dependencies, and run the tool.
   
### For Linux

1. Navigate to the `Linux` directory:
   ```bash
   cd Linux
   ```
2. Open a terminal and ensure the script is executable:
   ```bash
   chmod +x setup.sh
   ```
3. Run the setup script:
   ```bash
   ./setup.sh
   ```
   - This will create a virtual environment, upgrade `pip`, install dependencies, and run the tool.

## Using the Tool

Once the tool is set up, you can select individual files or directories for batch conversion. Converted files will be placed in the `Converted_Batch_Documents` folder inside the working directory.

### Features

- Converts `.xlsx` files to `.csv`.
- Converts `.docx` files to `.md`.
- Supports batch processing of files or directories.

## Notes

- This version of the tool does **not** extract or handle embedded images in `.docx` files.
- Ensure you have sufficient permissions to run the setup scripts on your platform.

## License

This project is licensed under the MIT License.