# SUITE OF FUNCTIONS TO AUTOMATE EXCEL REPORT FORMATTING WITH XLSXWRITER

def last_col_highlight_header(df, wb, sheet, bg_color1, font_color1, bg_color2, font_color2, bg_color3, font_color3):

    # This function will apply formatting to your headers that will automatically apply a different color to your last column to highlight it
    ## This is especially useful for time series: highlighting most recent year etc
    ## Color for index can also be set to be different, or the same as normal header columns
    ### This function should be applied to data that has already been loaded into a worksheet via to_excel() from dataframe
    ### Meant only for dataframes with single row index and columns

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### all color args can be added with keywords (ie, 'red') but hex codes (ex '#FF0000') are better for customization
    ### bg_color1 is the background color for your column headers
    ### font_color1 is the font color for your column headers
    ### bg_color2 is the background color for your LAST column header
    ### font_color2 is the font color for your LAST column header
    ### bg_color3 is the background color for your index header
    ### font_color3 is the font color for your index headers

    # getting column count of the data to use to set upper bound for formatting
    df_column_count = len(df.columns)

    # create format templates
    header_format = wb.add_format({'bold':True,'bg_color':bg_color1,'font_color':font_color1,'align':'center','bottom':True})
    last_col_format = wb.add_format({'bold':True,'bg_color':bg_color2,'font_color':font_color2,'align':'center','bottom':True})

    ## the header_format template is applied in the first row for all columns, which also keeps the value from the df header row
    ## for the last column, the color of the header row will be different, applying last_col_format
    ## the for loop goes over all columns. this prevents the formatting being applied to empty cells
    ### using enumerate and calling values will extract the column value (in this case, column header)
    for col_num, value in enumerate(df.columns.values):
        # because col_num starts at 0 in python, 1 must be added to it so that number of the last column equals the column count
        # the special latest_period formatting will only be applied to the last column
        if col_num + 1 == df_column_count:
            # the first argument of 0 specifies this will be applied to the first row of the excel spreadsheet
            ## col_num + 1 here is so that formatting is applied to the column headers only
            ## because the index row is not counted as a column by the loop
            sheet.write(0, col_num + 1, value, last_col_format)
        else:
            # normal header formatting is applied to all other columns
            sheet.write(0, col_num + 1, value, header_format)

    # the header loop cannot be applied to the index, so formatting is manually applied by overwriting the cell 
    ## also allows adding R border to it only
    index_format = wb.add_format({'bold':True,'bg_color':bg_color3,'font_color':font_color3,'align':'left','bottom':True,'right':True}) 
    # the name of the index is selected and to be put in cell A1
    df_index_name = df.index.name 
    # write to cell A1
    sheet.write('A1', df_index_name, index_format)