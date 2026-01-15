# Instructions

1. `output_page` is a single page generated based on data from input file. Each page is generated based on a distinct name from input file. Input file should produce multiple such files.
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

2. `extra` is a specific summary data for each distinct name based on each output_page.
   - A1-F13 are static cells
   - A11-12 autoincremental number starting at 1
   - B11-12 name (same distinct name from the previous files)
   - C11-12 just a static column, leave as is
   - D11-12 quantity_in (sum of G column from respective output_page)
   - E11-12 quantity_out (sum of H column from respective output_page)
   - F11-12 quantity_current (I column, last value)

3. `contents` is navigation. For now only first 3 columns are needed.
   - A1-L3 are headers
   - A is name (same distinct name from the previous files)
   - B is link to the respective output_page
