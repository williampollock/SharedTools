# SharedTools Repository

This repository provides a collection of helpful python functions that folks can contribute to for training purposes.  Many of the current functions are focused on working with Excel spreadsheets in Python. In particular, the `get_xls_colors.py` module contains functions to extract cell colors from Excel workbooks using the `openpyxl` library. This README explains how to set up the project on Windows using Visual Studio Code and run a Jupyter Notebook demo that retrieves cell colors from the range D4:D29 in the "Complex Table" sheet of the example spreadsheet (`Table 1 ESUs.xlsx`).

## Prerequisites

- **Git:** Download and install from [git-scm.com](https://git-scm.com/). You can also install
  git directly from within VS Code.  THe first time you open the package control tab on the left side of the VS Code window, you should see a big blue button that says: `Download Git for Windows`.  Click that button and follow the installation instructions.
- **Python:** Download and install from [Shanwilpy Installation Guide](https://swi.blob.core.windows.net/shanwilpydocs/html/install.html).
- **Visual Studio Code:** Download and install from [code.visualstudio.com](https://code.visualstudio.com/).
- **VS Code Extensions:**  
  - Python (provided by Microsoft)  
  - Jupyter (if not already installed)
  - Ruff (recommended linter)

## Repository Contents

- `get_xls_colors.py` – Python module with functions to retrieve cell colors.
- `Table 1 ESUs.xlsx` – Example Excel spreadsheet (should be located in the repository root).
- `requirements.txt` – Lists the project dependencies.
- `README.md` – This file.

## Setup Instructions

### 1. Clone the Repository

1. Find a folder on your PC that you want the local repository to reside.  Preferably a central location
   where you keep other python repositories.
3. Open VS Code.  If you set up VS Code properly, then you should be able to right click in
   the folder above and select "Open with Code".
4. Open the integrated terminal in VS Code by selecting View > Terminal from the menu.
5. Run the following command in the integrated terminal:
   
   `git clone https://github.com/ohoopes/SharedTools.git`
7. Open the cloned SharedTools folder in VS Code.

### 2. Create and Activate the Virtual Environment via VS Code integrated terminal
1. In the VS Code integrated terminal, navigate to the repository root.
2. Create a virtual environment named `shared_tools_env`.  It is good practice
   to create seperate virtual environments for each workflow or group of similar
   workflows, mainly because they help keep dependencies separate for different 
   projects. This also prevents conflicts where one project requires a package 
   version that another project is incompatible with. 

   `python -m venv shared_tools_env`
4. Activate the virtual environment:

   `shared_tools_env\Scripts\activate`
6. Set your global username and email so you can contribute to this repository via git.  Open the integrated terminal in VS Code by selecting View > Terminal from the menu.

   `git config --global user.name "Your Name"`

   `git config --global user.email "your.email@example.com"`
   Typically, you use your full name for user.name. Git doesn't enforce a specific format — it simply records what you set as your user name in the commit metadata. If you'd prefer to use your GitHub username, that's fine too.
8. To verify, run:

   `git config --global --list`

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



