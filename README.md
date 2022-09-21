# xlsxwriter_automated_formatting_functions
## a suite of functions to automate report formatting from pandas dataframes with xlsxwriter

## Overview of Project

### Purpose

The developer is creating a package of functions to automate report formatting in Excel when exporting pandas dataframe data via the xlsxwriter package. These functions will create nice-looking reports with formatting responsive to changes in data size or shape. There are default aesthetic options but some customization is possible with optional arguments in some functions.

### Structure

The code base with the functions themselves can be found in [this](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/formatting_functions_open_source.py) Python script, which features commented documentation. An example of how to use the functions to create a report may be found in [this](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/create_example_report.ipynb) Jupyter Notebook, in which the developer created an example report using mocked-up clinical healthcare data. The example report itself can be found [here](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Reports/Example%20Clinical%20Report.xlsx).

## Results

### Two Dimensional Data

An example of how a report could be formatted with two dimensional data, with a single row index and basic columns is the first tab of the example report. The initial csv data looked like this:

![two dimensional raw data csv](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/1_2D_before.png)

The `last_col_highlight_header`, `format_index`, `format_single_numeric_data_type_df`, `set_column_widths`, `table_bottom_border`, and `table_right_border` functions were applied to the data after it had been loaded in and had its index set. The resulting report tab looked like this:

![two dimensional data report table](https://github.com/cbeckler/xlsxwriter_automated_formatting_functions/blob/main/Resources/1_2D_after.png)

### Three Dimensional Data--Row Multiindex with Columns

to be continued...
