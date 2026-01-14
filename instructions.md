# Instructions

- use `xlsx-template` or `exceljs` to parse the data on node.js server preserving all the styles and templates and formatting
- use [react-spreadsheet](https://iddan.github.io/react-spreadsheet/learn/usage) to visualize data on the React frontend
- `src/sheet_templates` include 4 xlsx files
- actual raw templates should be used from subfolder `raw`
- Dates are formatted as `dd.mm.yyyy`
- all files are Microsoft Excel files so referencing columns and rows would be respective.
- if any column/row is not specified by me – it means it should stay as is if its a static header-related
- listed files above are input/output templates
- output should be a generated in "src/data/output"
- output data should be generated ONLY based on `db.xlsx`. All the other xlsx files are for example purposes.
- expected output is a sinlge `output.xlsx` file with all the sheets generated in a single workbook
- I expect the generated output to be identical to the example output files

Files:

1. `db.xlsx` is the input file.
   - It includes of 10 columns: date, name, document name, document number, document date, from, to, price, category, quantity.
   - Rest of the columns after Column J are to be ignored.
   - Row 1 is the header row, all others are data rows.

2. `output_page.xlsx` is a single page generated based on data from input file. Each file is generated based on a distinct name from input file. Input file should produce multiple such files.
   - Rows 1-12 are headers
   - Row 12 is numerical autoincrement starting from 1
   - It includes two section of columns: sticky static and scrollable dynamic.
     - Sticky and static:
       - Columns A-N are sticky.
       - Rows 1-12 are headers.
       - A15-N15 is the distinct name from input file for this file
       - A8-11 date (distinct per date and document_name from input file)
       - B8-11 name (distinct document_name from input file)
       - C8-11 doc_number
       - D8-11 doc_date
       - E8-11 from
       - F8-11 to
       - G8-11 quantity_in
       - H8-11 quantity_out
       - I-N 8-9 breakdown
         - I9-11 total (sum)
         - J-N 10 breakdown (sum for each column per each category)
           - J-N 11 are 5 separate category columns based on category col from input file (values: roman 1 to 5, I II III IV V)
     - Dynamic and scrollable:
       - Same as I-N 8-9 breakdown but for each distinct "from/to" entity. Each entity should get 6 generated columns with I-N 8-9 line being the name of the entity.

3. `extra.xlsx` is a file containing specific summary data for each distinct name based on each output_page.
   - A1-F13 are static cells
   - A11-12 autoincremental number starting at 1
   - B11-12 name (same distinct name from the previous files)
   - C11-12 just a static column, leave as is
   - D11-12 quantity_in (sum of G column from respective output_page)
   - E11-12 quantity_out (sum of H column from respective output_page)
   - F11-12 quantity_current (I column, last value)

4. `contents.xlsx` is a file containing the navigation. For now only first 3 columns are needed.
   - A1-L3 are headers
   - A is name (same distinct name from the previous files)
   - B is link to the respective output_page

Use files at `src/sheet_templates/raw` as templates to append data to. Styling has to be intact.
