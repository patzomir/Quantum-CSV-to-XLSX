# Quantum-CSV-to-XLSX
Converts Quantum csv files to xlsx ones
Examples are in test/

# Requirements for the csv input
* Each table should begin with a single cell with string "#page"
* Order of titles:
  * Table set title
  * Table title
  * Sub-titles
  * Base-text - should be specified as "Base: "
* Total row (the row on which all percentages are calculated) - should begin with "Total"/"Base"/"Weighted"
* Tables can be put on different sheets using "$$sheet_name$$<sheet name>" in first cell

Dependencies:
* Python 2.7
* xlsxwriter

# Run
`python lib/format.py <input_file> <output_file> <row number for Title after #page row> <0 (single sheet) / 1 (many sheets)>`
