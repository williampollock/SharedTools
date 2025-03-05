# SharedTools Repository

This repository provides a collection of tools for working with Excel spreadsheets in Python. In particular, the `get_xls_colors.py` module contains functions to extract cell colors from Excel workbooks using the `openpyxl` library. This README explains how to set up the project on Windows using Visual Studio Code and run a Jupyter Notebook demo that retrieves cell colors from the range D4:D29 in the "Complex Table" sheet of the example spreadsheet (`Table 1 ESUs.xlsx`).

## Prerequisites

- **Git:** Download and install from [git-scm.com](https://git-scm.com/).
- **Python:** Download and install from [python.org](https://www.python.org/).
- **Visual Studio Code:** Download and install from [code.visualstudio.com](https://code.visualstudio.com/).
- **VS Code Extensions:**  
  - Python (provided by Microsoft)  
  - Jupyter (if not already installed)

## Repository Contents

- `get_xls_colors.py` – Python module with functions to retrieve cell colors.
- `Table 1 ESUs.xlsx` – Example Excel spreadsheet (should be located in the repository root).
- `requirements.txt` – Lists the project dependencies.
- `README.md` – This file.

## Setup Instructions

### 1. Clone the Repository

1. Open Visual Studio Code.
2. Open the integrated terminal by selecting View > Terminal from the menu.
3. Run the following command:
   `git clone https://github.com/ohoopes/SharedTools.git`
4. Open the cloned SharedTools folder in VS Code.

### 2. Create and Activate the Virtual Environment via VS Code integrated terminal
1. In the VS Code integrated terminal, navigate to the repository root.
2. Create a virtual environment named shared_tools_env
    `python -m venv shared_tools_env`
3. Activate the virtual environment:
    `shared_tools_env\Scripts\activate`

### 3. Install Dependencies
1. With the virtual environment activated, install the dependencies by running:
    pip install -r requirements.txt

### 4. Configure Visual Studio Code
1. Select the Python Interpreter:
    - Press+Shift+P to open the Command Palette.
    - Type `Python: Select Interpreter` and choose the interpreter from the shared_tools_env virtual environment.
2. Jupyter Support:
    If prompted, install the Jupyter extension to enable notebook support in VS Code.

### 5. Run the Jupyter Notebook Demo
1. open demo_shared_tools.ipynb from SharedTools directory via VS Code file explorer
2. Run the Notebook Cell:
Click the “Run” button at the top of the cell or press Shift + Enter to execute the code. The output will display the list of colors and a dictionary mapping each cell (e.g., ('D', 4)) to its corresponding color.



