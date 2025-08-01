# 📊 SAP Analyzer

A Python utility for analyzing **SAP BSEG and BKPF data** and comparing it to a **summary balance sheet**.  
It performs:

- Data loading and cleaning (BSEG, BKPF, and summary sheet)
- Merging of SAP tables on `belegnr`
- Validation of missing or unmatched records
- Comparison of **computed balances vs summary balances** per `hauptbuch` (GL account)
- Optional **Excel export** of merged data and comparison results

---

## 📦 Requirements

- Python 3.8+
- [pandas](https://pandas.pydata.org/)
- [numpy](https://numpy.org/)
- [openpyxl](https://pypi.org/project/openpyxl/) *(for Excel export)*

## 📂 Expected Input Files
BSEG.csv – SAP BSEG extract (pipe-separated |)

BKPF.txt – SAP BKPF extract (pipe-separated |)

Susa_BergerUndCo.xlsx – Summary balance sheet (Excel format)

### Key columns used:

belegnr – Document number for joining

hauptbuch – GL account

betrag hauswähr – Amount in company currency

endsaldo – Ending balance from summary

## ▶️ How to Run
- Clone the Repository:\
    git clone https://github.com/jadonblaise/Sap_BSEG_BKPF.git \
    cd Sap_BSEG_BKPF
- Activate virtual environment: \
    source venv/bin/activate -- Linux/ MacOS \
    venv/scripts/activate -- Windows
  
- Place your BSEG, BKPF, and summary files in the same directory as the script.

- Update the file paths in the __main__ section if your file names differ: \
    Analyzer = SAPAnalyzer(
      bseg_path="BSEG.csv",
      bkpf_path="BKPF.txt",
      summary_path="Susa_BergerUndCo.xlsx")

- Install dependencies:

  pip install -r requirements.txt

- Run the script:
    python SAP.py

- The script will:

    Load and clean the datasets
    
    Merge BSEG and BKPF
    
    Validate missing records
    
    Compare computed balances to the summary sheet

- (Optional) Uncomment the export_results line to save output to Excel:\
    Analyzer.export_results("Abstimmung.xlsx")

## 📄 Notes
Ensure betrag hauswähr and endsaldo columns are properly formatted for numeric conversion.

The script automatically normalizes GL account numbers (hauptbuch) to 6-digit strings for accurate matching.

Missing or unmatched records are reported in the console.
