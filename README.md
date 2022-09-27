# xlsxwriter_automated_formatting_functions
## a suite of functions to automate report formatting from pandas dataframes with xlsxwriter

## Overview of Project

### Purpose

The developer is creating a package of functions to automate report formatting in Excel when exporting pandas dataframe data via the xlsxwriter package. These functions will create nice-looking reports with formatting responsive to changes in data size or shape. There are default aesthetic options but some customization is possible with optional arguments in some functions.

### Structure

The code base with the functions themselves can be found in [this](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/formatting_functions_open_source.py) Python script, which features commented documentation. An example of how to use the functions to create a report may be found in [this](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/create_example_report.ipynb) Jupyter Notebook, in which the developer created an example report using mocked-up clinical healthcare data. The example report itself can be found [here](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Reports/Example%20Clinical%20Report.xlsx). The mocked up raw data used to create the report can be found [here](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/tree/main/Data).

## Results

### One Dimensional Data

An example of how a report could be formatted with one dimensional data, with basic columns and no indices is the first tab of the example report.The initial csv data looked like this:

![one dimensional raw data csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/0_1D_before.png)

The  `insert_data`, `set_col_data_type`, `set_col_width` (for the columns with no set data type), `table_bottom_border`, `table_right_border`, `insert_title`, and `format_header` (with `clean_header` option) functions were applied with a `header_offset` of 2 after the data was loaded in and date columns were changed to pandas datetime. The resulting report tab looked like this:

![one dimensional data report table](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/0_1D_after.png)

### Two Dimensional Data

An example of how a report could be formatted with two dimensional data, with a single row index and basic columns is the second tab of the example report. The initial csv data looked like this:

![two dimensional raw data csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/1_2D_before.png)

The `last_col_highlight_header`, `format_index`, `insert_data` with `data_type` argument, `set_column_widths`, `table_bottom_border`, `table_right_border`, and `insert_title` functions were applied to the data after it had been loaded in and had its index set, with a `header_offset` of 2. The resulting report tab looked like this:

![two dimensional data report table](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/1_2D_after.png)

### Three Dimensional Data--Row Multiindex with Columns

An example of how a report could be formatted with three dimensional data, with a row multiindex and basic columns is in the third tab of the example report. The initial csv data looked like this, with two row index colums:

![three dimensional row index raw data 1 csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/2_3D_row_1_before.png)

The `last_col_highlight_header`, `merge_row_index_cells`, `format_row_multiindex`, `insert_row_multiindex_data` with `data_type` argument, `set_column_widths`, `table_bottom_border`, `table_right_border`, and `insert_title` werea applied to the data after if had been loaded in and had its index set, with a `header_offset` of 2. The resulting report tab looked like this:

![three dimensional row index report table 1 csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/2_3D_row_1_after.png)

Even more complex row multiindices may be run through these functions. Another set of data with three row index columns was run through the same functions to create another table on the fourth tab of the example report. The initial csv data looked like this:

![three dimensional row index raw data 2 csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/2_3D_row_2_before.png)

After the same functions were applied to it with the same arguments as in the previous table, the resulting report looked like this:

![three dimensional row index report table 2 csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/2_3D_row_2_after.png)

### Three Dimensional Data--Single Row Index with Column Multiindex

To be continued...
