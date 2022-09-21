# SUITE OF FUNCTIONS TO AUTOMATE EXCEL REPORT FORMATTING WITH XLSXWRITER

######################## HEADER FORMATTING ##################################

###                 SINGLE ROW INDEX AND COLUMNS DATAFRAMES                 ###

def format_header(df, wb, sheet,  bg_color1='#002387', font_color1='#FFFFFF', bg_color2='#002387', font_color2='#FFFFFF'):

    # This function will apply formatting to your header row    
    ## Color for index can also be set to be different, or the same as normal header columns. It is the same by default
    ### This function should be applied to data that has already been loaded into a worksheet via to_excel()
    ### Meant only for dataframes with single row index and columns 

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    ## OPTIONAL:
    ### all color args can be added with keywords (ie, 'red') but hex codes (ex '#FF0000') are better for customization
    #### bg_color1 is the background color for your column headers
    #### font_color1 is the font color for your column headers
    #### bg_color2 is the background color for your index header
    #### font_color2 is the font color for your index headers

    # create format templates
    header_format = wb.add_format({'bold':True,'bg_color':bg_color1,'font_color':font_color1,'align':'center','bottom':True})

    ## the header_format template is applied in the first row for all columns, which also keeps the value from the df header row
    ## the for loop goes over all columns. this prevents the formatting being applied to empty cells
    ### using enumerate and calling values will extract the column value (in this case, column header)
    for col_num, value in enumerate(df.columns.values):
        # normal header formatting is applied to all header columns
        ## col_num + 1 here is so that formatting is applied to the column headers only
        sheet.write(0, col_num + 1, value, header_format)

    # the header loop cannot be applied to the index, so formatting is manually applied by overwriting the cell 
    ## also allowing me to add R border to it only
    index_format = wb.add_format({'bold':True,'bg_color':bg_color2,'font_color':font_color2,'align':'left','bottom':True,'right':True}) 
    # the name of the index is selected and to be put in cell A1
    df_index_name = df.index.name 
    # instead of writing by row_number, column_number, it is possible to write to the specific cell (A1)
    ## this is only recommended if this cell will be the same every time
    sheet.write('A1', df_index_name, index_format)



def last_col_highlight_header(df, wb, sheet, bg_color1='#002387', font_color1='#FFFFFF', bg_color2='#00A111', font_color2='#FFFFFF', bg_color3='#002387', font_color3='#FFFFFF'):

    # This function will apply formatting to your headers that will automatically apply a different color to your last column to highlight it
    ## This is especially useful for time series: highlighting most recent year etc
    ## Color for index can also be set to be different, or the same as normal header columns. Is the same as normal by default
    ### This function should be applied to data that has already been loaded into a worksheet via to_excel() from dataframe
    ### Meant only for dataframes with single row index and columns

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    ## OPTIONAL:
    ### all color args can be added with keywords (ie, 'red') but hex codes (ex '#FF0000') are better for customization
    #### bg_color1 is the background color for your column headers
    #### font_color1 is the font color for your column headers
    #### bg_color2 is the background color for your LAST column header
    #### font_color2 is the font color for your LAST column header
    #### bg_color3 is the background color for your index header
    #### font_color3 is the font color for your index headers

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
        # the special last_col_format formatting will only be applied to the last column
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


######################## INDEX FORMATTING ##################################

###                 SINGLE ROW INDEX AND COLUMNS DATAFRAMES                 ###

def format_index(df, wb, sheet):

    # This function will apply formatting to your index to bold it and give a right border
    ## This function should be applied to data that has already been loaded into a worksheet via to_excel()
    ## Meant only for dataframes with single row index and columns   

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    # create index format
    index_format = wb.add_format({'bold':True,'right':True})

    ## this iterates through the rows.  this prevents the formatting being applied to empty cells
    ## it applies formatting with the index value for the first column of the report
    ## enumerate is called on the index to get those values
    for row_num, value in enumerate(df.index.values):
        # 1 is added to row num so that we don't start on 0 and overwrite our header!
        # the column is hard-coded to 0 (column A) as this is the only column we want this applied to
        sheet.write(row_num + 1, 0, value, index_format)

    # gets the length of all the values in the index
    index_values = [len(value) for i, value in enumerate(df.index.values)]

    # gets the max of the index values or the name of the index, whichever is greater
    ## + 1 for 'wiggle room'
    max_index_length = max(max(index_values), len(df.index.name)) + 1

    # set index column width
    sheet.set_column('A:A', max_index_length)


######################## DATA FORMATTING ##################################

###                      ANY SHAPE DATAFRAMES                        ###

def format_single_numeric_data_type_df(df, wb, sheet, data_type, col_width=14):

    # This function will apply the specified numeric formatting to all data columns
    ## This function should be applied to data that has already been loaded into a worksheet via to_excel()
    ## Meant only for dataframes that have the same data type for ALL non-index columns, but can have any number of columns and indices
    ### Note: this will set ALL data columns to the same width!

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### data_type is the type of numeric data:
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
        
    ## OPTIONAL:
    ### col_width is the width of the data columns

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
    elif data_type == 'decimal':
        data_format = wb.add_format({'num_format':'#,##0.00'})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
    else:
        raise ValueError(f"{data_type} is not a valid argument for data_type")

    # getting column count of the data to use to set upper bound for formatting
    df_column_count = len(df.columns)
    
    # getting row indices count of the data to use to set lower bound for formatting
    num_row_indices = len(df.index.names)

    ## sets columns B through the last column present in the dataset with the specified data_format and and sets column widths
    sheet.set_column(num_row_indices, df_column_count, col_width, data_format)


###                 ANY NUMBER ROW INDEX AND SINGLE COLUMNS INDEX DATAFRAMES                 ###

def set_column_widths(df, wb, sheet):

    # adapted from a solution found at https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter

    # This function will automatically make all columns wide enough for their full column names to appear without being cut off
    ## Meant for use on data with only one index of columns, but any number of row indices

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    # create a list holding the length of the name of each column
    ## + 1 for 'wiggle room'
    col_name_lengths = [len(name) + 1 for name in df.columns]

    # get the count of how many row indices they are so we can skip those columns in the for loop
    num_row_indices = len(df.index.names)

    # iterating over the df columns:
    for i, width in enumerate(col_name_lengths):
        # apply the matching width to the column
        sheet.set_column(i + num_row_indices, i + num_row_indices, width)  


######################## EDGE BORDER FORMATTING ##################################

###                      ANY SHAPE DATAFRAMES                        ###

def table_bottom_border(df, wb, sheet):

    # This function will apply formatting a bottom border to your table
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    # getting row count of the data to use to set lower bound for formatting
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1
    # get the row count (which doesn't count column header rows)
    data_rows = len(df)
    # add the two together to get total row count
    df_row_count = num_col_indices + data_rows

    # getting count of number of row indices to set range for index formatting
    num_row_indices = len(df.index.names)

    # creating the format for the bottom border (actually top border on the cell below so we don't overwrite data)
    bottom_format = wb.add_format({'top':True})

    # this applies a top border to the cell below the last row fo data for all the columns except the index
    for col_num, value in enumerate(df.columns.values):
        # we are applying a top border to that to fake a bottom border on the table!
        # "" is filling in the cell with nothing, leaving it blank
        # col_num + row_indices will correctly skip the row index columns in the loop
        sheet.write(df_row_count, col_num + num_row_indices, "", bottom_format)
    
    # this applies a top border to the cells below the last row of the index columns since they are excluded from the column for loop
    for i in range(num_row_indices):
        sheet.write(df_row_count, i, "", bottom_format)


def table_right_border(df, wb, sheet):

    # This function will apply formatting a right border to your table
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    # getting the column count

    # getting the count of row index columns
    num_row_indices = len(df.index.names)
    # getting the count of regular columns
    num_cols = len(df.columns)
    # adding them together for total column count
    total_cols = num_row_indices + num_cols

    # getting row count of the data to use to set lower bound for formatting
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1
    # getting count of the data rows (which doesn't count column header rows)
    data_rows = len(df)
    # adding them together to get total rows
    total_rows = num_col_indices + data_rows

    # creating right border format--actually left to next cell over to avoid overwriting data
    right_format = wb.add_format({'left':True})

    # iterating over all our rows in our table:
    for i in range(total_rows):
        # apply the right format to the first column after our table
        sheet.write(i, total_cols, "", right_format)