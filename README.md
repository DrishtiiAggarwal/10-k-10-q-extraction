# 10-K & 10-Q Financial Data Extraction

This project automates the extraction of financial statements and segment data from SEC filings (10-K and 10-Q) for all 2600 US stocks. The extracted data includes income statements, balance sheets, cash flow statements, revenue segments, and geographical segments, ensuring structured and accurate data retrieval.

## Features

- **Automated Financial Data Extraction**
  - Extracts income statements, balance sheets, and cash flow statements from SEC filings.
  - Retrieves segment data, including revenue and geographical details.

- **Accurate Data Retrieval**
  - Utilizes regex, string matching, and fuzzy search techniques to identify and retrieve relevant financial data sheets.
  - Leverages stack-based methods to handle and format nested financial data.

- **Structured Data Output**
  - Implements a relational data structure to manage hierarchical financial data, ensuring accurate associations between parent and child entries.

- **Scalable and Reliable**
  - Supports data extraction for all 2600 publicly traded US stocks.
  - Designed to handle complex data structures and ensure clean outputs.

## Technologies Used

- **Programming Language:** Python
- **Libraries & Tools:**
  - [Selenium](https://www.selenium.dev/) with Webdriver Manager for web automation.
  - [Pandas](https://pandas.pydata.org/) for data manipulation and analysis.
  - [openpyxl](https://openpyxl.readthedocs.io/) for handling Excel files.
  - [Psycopg2](https://www.psycopg.org/) for PostgreSQL database integration.
  - Regex and fuzzy search techniques for text processing.

- **Database:** PostgreSQL

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/DrishtiiAggarwal/10-k-10-q-extraction.git
   cd 10-k-10-q-extraction
   ```

2. Set up a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Set up the PostgreSQL database and update the connection details in the script.

## Usage

1. Run the script to start extracting financial data:
   ```bash
   python main.py
   ```

2. The script will automatically navigate to the SEC filings, extract the required data, and store it in the PostgreSQL database.

## Output

- Extracted data is stored in a structured format in the PostgreSQL database.
- Optional: Export data to Excel files for further analysis using the openpyxl library.

## Contact

For any questions or support, feel free to reach out:
- **Author:** Drishti Aggarwal
- **GitHub:** [DrishtiiAggarwal](https://github.com/DrishtiiAggarwal)

