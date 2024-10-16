# File Converter Tool for Linux

This tool converts `.xlsx` files to `.csv` and `.docx` files to `.md` (Markdown).

## Requirements
- Ubuntu or any Linux distribution with Python 3.x
- `pip` installed (the script will upgrade `pip` automatically)

## Setup Instructions

1. Clone or download the project files.
2. Open a terminal and navigate to the project directory.
3. Ensure the script is executable by running:
   ```bash
   chmod +x setup.sh
   ```
4. Run the setup script with the following command:
   ```bash
   ./setup.sh
   ```

The script will:
- Create a virtual environment if it doesn’t exist.
- Upgrade `pip` to the latest version using the virtual environment’s Python executable.
- Install the required dependencies from `requirements.txt`.
- Launch the file conversion tool.

## Using the Tool

- The tool supports batch conversion of multiple `.xlsx` and `.docx` files.
- You can select individual files or entire directories for conversion.
- After conversion, all files will be saved in the `Converted_Batch_Documents` folder created in the working directory.
- A single confirmation dialog will appear once all conversions are complete.

## Notes
- This version of the script does **not** extract or handle embedded images in `.docx` files.
- Ensure Python 3.x is installed and that you have necessary permissions to run the script.