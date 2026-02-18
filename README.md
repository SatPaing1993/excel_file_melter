# Commercial Bank Data Consolidation

**Description:**  
Python script to merge and clean multiple commercial bank Excel files.  
Converts messy data into a tidy dataset ready for analysis.  
Capstone project for Advanced Python training.

---

## Features

- Reads multiple Excel files from a folder
- Extracts file name and metadata
- Melts and reshapes data into tidy format
- Splits date into Month and Year
- Cleans and converts data types
- Consolidates all files into a single Excel output

---

## Getting Started

### Prerequisites

- Python 3.8 or higher
- [pandas](https://pandas.pydata.org/)  
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

Install dependencies:

pip install pandas openpyxl

Setup

Clone the repository:

git clone https://github.com/YourUsername/commercial_bank_data_consolidation.git
cd commercial_bank_data_consolidation

Place your Excel files in a folder, for example:

G:\My Drive\Data

Open Final Capstone Adv Python Project.py and set the input folder and output file name:

input_dir = r"G:\My Drive\Data"
output_file = "Final_Commercial_Bank_Consolidated.xlsx"

Run the script:

python "Final Capstone Adv Python Project.py"

Output
Single Excel file: Final_Commercial_Bank_Consolidated.xlsx

Columns: No, Label, Amount, Month, Year, File Name

Repository Structure
├── Final Capstone Adv Python Project.py
├── README.md
├── LICENSE
├── requirements.txt

License
This project is licensed under the MIT License.

Author
Sat Paing Oo

GitHub: SatPaing1993


