#!/bin/bash

# Create virtual environment if it doesn't exist
if [ ! -d "env" ]; then
    python3 -m venv env
    echo "Virtual environment created."
else
    echo "Virtual environment already exists."
fi

# Activate virtual environment
source env/bin/activate

# Upgrade pip using the virtual environment's Python executable
env/bin/python -m pip install --upgrade pip

# Install necessary packages from requirements.txt
env/bin/python -m pip install -r requirements.txt

# Run the file_converter.py tool
env/bin/python file_converter.py
