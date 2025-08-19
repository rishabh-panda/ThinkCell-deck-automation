# ThinkCell Chart Automation Tool

This repository provides a Python-based workflow to **automatically generate ThinkCell charts from Excel files** and update a **PowerPoint presentation deck** using a predefined template.  

The tool streamlines the process of refreshing business decks with updated data by leveraging **ThinkCell CLI**.

---

## Project Structure

```text
project-root/
│
├── _inputs/               # Input files and template deck
│   ├── _template_deck/    # Contains the single PowerPoint template (.pptx)
│   └── _refresh_files/    # Excel files with updated data for charts
│
├── _intermediate/         # Intermediate generated .ppttc files
│
├── _output/               # Final generated PowerPoint decks
│
├── decks/                 # (Optional) Storage for additional PPT decks
│
├── helpers/               # Helper scripts for data ingestion, JSON, PPTTC writing, ThinkCell CLI
│   ├── data_ingestion.py
│   ├── json_array.py
│   ├── temporary_ppttc.py
│   └── thinkcell_cli.py
│
├── venv/                  # Virtual environment (not committed to git)
│
├── .gitignore             # Git ignore rules
├── main.py                # Entry point script
└── requirements.txt       # Python dependencies
```

---

## How It Works

1. **Locate Template Deck**  
   The script searches `_inputs/_template_deck/` for a single `.pptx` file that serves as the base deck.

2. **Ingest Data Files**  
   Excel files in `_inputs/_refresh_files/` are read and pre-processed using helper functions (e.g., `data_ingestion.py`).

3. **Map Data Files to Charts**  
   Each Excel file is mapped to ThinkCell chart identifiers defined in `main.py`.  

   Example mapping:

   ```python
   file_chart_pairs = [
       ("Sales_YOY.xlsx", "Sales_YOY"),
       ("SpendItem_YOY.xlsx", "SpendItem_YOY"),
       ("Slide_5_Groq_4.xlsx", "Slide_5a"),
       ("Slide_5_Perplexity.xlsx", "Slide_5b"),
       ...
   ]
   ```
4. **Generate ThinkCell JSON Config**
   Using ```json_array.py```, the Excel-to-chart mappings are converted into a ThinkCell JSON configuration.

5. **Write Intermediate PPTTC File**
   With ```temporary_ppttc.py```, the JSON configuration is written to ```_intermediate/combined_charts.ppttc```.

6. **Run ThinkCell CLI**
   Using ```thinkcell_cli.py```, the .ppttc file is applied to the template deck to generate the final updated presentation in ```_output/Updated_PPT.pptx```.
---

## Setup

1. Clone Repository
   Run the following commands to clone the repository and navigate to the project root:
   ```text
   git clone <repo-url>
   cd project-root
   ```

3. Create Virtual Environment
   Create and activate a virtual environment:
   ```text
   python -m venv venv
   source venv/bin/activate    # On Mac/Linux
   venv\Scripts\activate       # On Windows
   ```

5. Install Dependencies
   Install the required Python dependencies:
   ```text
   pip install -r requirements.txt
   ```

7. Install ThinkCell
   Ensure ThinkCell and its CLI integration are installed on your machine.  
   Refer to ThinkCell’s official documentation for enabling CLI support.

---

## Usage

Run the main script to refresh the PowerPoint deck:
```text
python main.py
```

This will:
- Read the Excel files in `_inputs/_refresh_files/`
- Update charts in the template deck located in `_inputs/_template_deck/`
- Generate a new PowerPoint file at `_output/Updated_PPT.pptx`

---

## Notes

- There must be exactly one template deck in `_inputs/_template_deck/`.
- Excel filenames and chart identifiers must match the ThinkCell placeholders in the template.
- Intermediate `.ppttc` files are stored in `_intermediate/` for debugging.
- Additional decks may be stored in the `decks/` folder for reference or archival.

---

## Example Workflow

1. Place your base PowerPoint template in:
   _inputs/_template_deck/template.pptx

2. Drop your refreshed Excel data files in:
   _inputs/_refresh_files/

3. Run:
   ```text
   python main.py
   ```

5. Retrieve your updated deck at:
   _output/Updated_PPT.pptx
