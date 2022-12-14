# SUITE OF FUNCTIONS TO AUTOMATE EXCEL REPORT FORMATTING WITH XLSXWRITER

######################## HEADER FORMATTING ##################################

###                 ANY NUMBER ROW INDICES AND SINGLE COLUMNS INDEX DATAFRAMES                 ###

from textwrap import wrap


def format_header(df, wb, sheet,  header_bgcolor = '#002387', header_fontcolor = '#FFFFFF', index_bgcolor =  '#002387', \
index_fontcolor = '#FFFFFF', header_offset=0, column_offset=0, clean_header=False, text_wrap=False):

    # This function will apply formatting to your header row    
    ## Index is same color as normal column headers, but this can be changed if desired w/ index_color optional args
    ### Meant only for dataframes with any number of row indices and columns 

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    ## OPTIONAL:
    ## all color args can be added with keywords (ie, 'red') but hex codes (ex '#FF0000') are better for customization
    ### header_bgcolor is the background color for your column headers
    ### header_fontcolor is the font color for your column headers
    ### index_bgcolor is the background color for your index header
    ### index_fontcolor is the font color for your index headers
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### clean_header will give your columns title format names (ex: Birth Date) instead of underscore (birth_date)
    ###     or CamelCase (BirthDate)
    ### text_wrap will wrap the column header labels 
    
    from utility_functions import clean_header_string

    # getting count of number of row indices to set range for index formatting
    
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # create format templates
    if text_wrap == True:
        header_format = wb.add_format({'bold':True,'bg_color':header_bgcolor,'font_color':header_fontcolor,'align':'center',\
        'bottom':True, 'text_wrap':True, 'valign':'vcenter'})
    else:
        header_format = wb.add_format({'bold':True,'bg_color':header_bgcolor,'font_color':header_fontcolor,'align':'center',\
            'bottom':True})

    # optional clean header labels
    
    # if clean_header option is enabled:
    if clean_header == True:
        # create a list of column names
        col_list = [col_name for col_name in df.columns]
        # iterate through col_names and apply clean_header_string function
        fixed_col_names = [clean_header_string(col_name) for col_name in col_list]
        # if there are no row indices:
        if num_row_indices == 0:
            # skip this step
            pass
        # if there is a single row index:
        elif num_row_indices == 1:
            # then set the index name to the cleaned version
            fixed_index_names = [clean_header_string(df.index.name)]
        # if there is a row multiindex:
        else:
            # create a list of cleaned names
            fixed_index_names =  [clean_header_string(name) for name in df.index.names]
    # if clean_header is false:
    elif clean_header == False:
        # have a list of the regular col names
        fixed_col_names = [col_name for col_name in df.columns]
        # if there are no row indices:
        if num_row_indices == 0: 
            # skip this step
            pass
        # if there is a single row index
        elif num_row_indices == 1:
            # list the name of it
            fixed_index_names = [df.index.name]
        else:
            # list the name of all row indices
            fixed_index_names = [name for name in df.index.names]
    else:
        # else raise an error message that an incorrect argument has been given
        raise ValueError(f"{clean_header} is not a valid clean_header option. Valid arguments are True, False.")         

     
    ## the header_format template is applied in the first row for all columns, which also keeps the value from the df header row
    ## the for loop goes over all columns. this prevents the formatting being applied to empty cells
    ### using enumerate and calling values will extract the column value (in this case, column header)
    for col_num, value in enumerate(df.columns.values):
        # normal header formatting is applied to all header columns
        ## col_num + num_row_indices here is so that formatting is applied to the column headers only
        ## fixed_col_names[col_num] will retrieve the correct name based on its position in the list
        sheet.write(header_offset, col_num + num_row_indices + column_offset, fixed_col_names[col_num], header_format)

    # the header loop cannot be applied to the index, so formatting is manually applied by overwriting the cell 
    ## also allowing me to add R border to the rightmost index only

    if text_wrap == True:
        index_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left','bottom':True,\
            'right':True,'text_wrap':True, 'valign':'vcenter'}) 
        # the index headers to the left lack the right border
        index_left_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left',\
            'bottom':True,'text_wrap':True, 'valign':'vcenter'})
    else:
        index_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left','bottom':True,\
            'right':True}) 
        # the index headers to the left lack the right border
        index_left_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left',\
            'bottom':True})

    # iterating over the number of row indices present:
    for col_num in range(num_row_indices):
        # we extract the name of the index
        #index_name = df.index.names[col_num]
        # if the index is the last index in the range:
        if col_num == max(range(num_row_indices)):
            # insert the index name and apply the right border index format
            ## fixed_index_names[col_num] will retrieve the correct name based on its position in the list
            sheet.write(header_offset, col_num + column_offset, fixed_index_names[col_num], index_format)
        else:
            # else insert the index name and apply no right border index format
            sheet.write(header_offset, col_num + column_offset, fixed_index_names[col_num], index_left_format)



def last_col_highlight_header(df, wb, sheet, header_bgcolor = '#002387', header_fontcolor = '#FFFFFF',\
    hilite_bgcolor = '#00A111', hilite_fontcolor = '#FFFFFF', index_bgcolor = '#002387', index_fontcolor = '#FFFFFF', \
    header_offset=0, column_offset=0, clean_header=False, text_wrap=False):

    # This function will apply formatting to your headers that will automatically apply a different color to your last column to highlight it
    ## This is especially useful for time series: highlighting most recent year etc
    ## Index is same color as normal column headers, but this can be changed if desired w/ index_color optional args
    ### Meant only for dataframes with any number row indices and columns  

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    ## OPTIONAL:
    ## certain colors have keywords, but for most precision entering hex codes for colors is best
    ### header_bgcolor is the background color for your column headers
    ### header_fontcolor is the font color for your column headers
    ### hilite_bgcolor is the background color for your LAST column header
    ### hilite_fontcolor is the font color for your LAST column header
    ### index_bgcolor is the background color for your index header
    ### index_fontcolor is the font color for your index headers
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### clean_header will give your columns title format names (ex: Birth Date) instead of underscore (birth_date) 
    ###      or CamelCase (BirthDate)
    ### text_wrap will wrap the column header labels

    from utility_functions import clean_header_string
    
    # getting column count of the data to use to set upper bound for formatting
    ## the len function provides the length of objects--in this case, the list of columns
    df_column_count = len(df.columns)

    # getting count of number of row indices to set range for index formatting
    
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # optional clean header labels

    # if clean_header option is enabled:
    if clean_header == True:
        # create a list of column names
        col_list = [col_name for col_name in df.columns]
        # iterate through col_names and apply clean_header_string function
        fixed_col_names = [clean_header_string(col_name) for col_name in col_list]
        # if there are no row indices:
        if num_row_indices == 0:
            # skip this step
            pass
        # if there is a single row index:
        elif num_row_indices == 1:
            # then set the index name to the cleaned version
            fixed_index_names = [clean_header_string(df.index.name)]
        # if there is a row multiindex:
        else:
            # create a list of cleaned names
            fixed_index_names =  [clean_header_string(name) for name in df.index.names]
    # if clean_header is false:
    elif clean_header == False:
        # have a list of the regular col names
        fixed_col_names = [col_name for col_name in df.columns]
        # if there are no row indices:
        if num_row_indices == 0: 
            # skip this step
            pass
        # if there is a single row index
        elif num_row_indices == 1:
            # list the name of it
            fixed_index_names = [df.index.name]
        else:
            # list the name of all row indices
            fixed_index_names = [name for name in df.index.names]
    else:
        # else raise an error message that an incorrect argument has been given
        raise ValueError(f"{clean_header} is not a valid clean_header option. Valid arguments are True, False.")


    # create format templates

    if text_wrap == True:
        header_format = wb.add_format({'bold':True,'bg_color':header_bgcolor,'font_color':header_fontcolor,'align':'center',\
            'bottom':True,'text_wrap':True,'valign':'vcenter'})
        last_col_format = wb.add_format({'bold':True,'bg_color':hilite_bgcolor,'font_color':hilite_fontcolor,'align':'center',\
            'bottom':True,'text_wrap':True,'valign':'vcenter'})    
    else:
        header_format = wb.add_format({'bold':True,'bg_color':header_bgcolor,'font_color':header_fontcolor,'align':'center',\
            'bottom':True})
        last_col_format = wb.add_format({'bold':True,'bg_color':hilite_bgcolor,'font_color':hilite_fontcolor,'align':'center',\
            'bottom':True})

    ## the header_format template is applied in the first row for all columns, which also keeps the value from the df header row
    ## for the last column, the color of the header row will be different, applying last_col_format
    ## the for loop goes over all columns. this prevents the formatting being applied to empty cells
    ### using enumerate and calling values will extract the column value (in this case, column header)
    for col_num, value in enumerate(df.columns.values):
        # because col_num starts at 0 in python, 1 must be added to it so that number of the last column equals the column count
        # the special latest_period formatting will only be applied to the last column
        if col_num + 1 == df_column_count:
            # the first argument of 0 specifies this will be applied to the first row of the excel spreadsheet
            ## col_num + num_row_indices here is so that formatting is applied to the column headers only
            ## because the index row is not counted as a column by the loop
            ## fixed_col_names[col_num] will retrieve the correct name based on its position in the list
            sheet.write(header_offset, col_num + num_row_indices + column_offset, fixed_col_names[col_num], last_col_format)
        else:
            # normal header formatting is applied to all other columns
            sheet.write(header_offset, col_num + num_row_indices + column_offset, fixed_col_names[col_num], header_format)

    # the header loop cannot be applied to the index, so formatting is manually applied by overwriting the cell 
    ## also allowing me to add R border to the rightmost index only
    
    if text_wrap == True:
        index_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left',\
            'bottom':True,'right':True,'text_wrap':True,'valign':'vcenter'}) 
        # the index headers to the left lack the right border
        index_left_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left',\
            'bottom':True,'text_wrap':True,'valign':'vcenter'})
    else:    
        index_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left',\
            'bottom':True,'right':True}) 
        # the index headers to the left lack the right border
        index_left_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'align':'left',\
            'bottom':True})

    # iterating over the number of row indices present:
    for col_num in range(num_row_indices):
        # we extract the name of the index
        index_name = df.index.names[col_num]
        # if the index is the last index in the range:
        if col_num == max(range(num_row_indices)):
            # insert the index name and apply the right border index format
            ## fixed_index_names[col_num] will retrieve the correct name based on its position in the list
            sheet.write(header_offset, col_num + column_offset, fixed_index_names[col_num], index_format)
        else:
            # else insert the index name and apply no right border index format
            sheet.write(header_offset, col_num + column_offset, fixed_index_names[col_num], index_left_format)

###                 ANY NUMBER ROW INDICES AND TWO LEVEL COLUMN MULITINDEX DATAFRAMES                 ###

def format_header_multiindex(df, wb, sheet,  header1_bgcolor = '#002387', header1_fontcolor = '#FFFFFF', \
    header2_bgcolor = '#137A78', header2_fontcolor = '#FFFFFF', index1_bgcolor =  '#002387', index2_bgcolor = '#137A78', \
    index2_fontcolor = '#FFFFFF', header_offset=0, column_offset=0, clean_header=False, merge_cells=False, text_wrap=False):

     # This function will apply formatting to your header rows    
    ## Index is same color as normal column headers, but this can be changed if desired w/ index_color optional args
    ### Meant only for dataframes with any number of row indices and two header rows (2 level column multiindex) 

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    ## OPTIONAL:
    ## all color args can be added with keywords (ie, 'red') but hex codes (ex '#FF0000') are better for customization
    ### header1_bgcolor is the background color for your column headers for your first row
    ### header1_fontcolor is the font color for your column headers for your first row
    ### header2_bgcolor is the background color for your column headers for your second row
    ### header2_fontcolor is the font color for your column headers for your second row    
    ### index1_bgcolor is the background color for your index headers for your first row
    ### index2_bgcolor is the background color for your index headers for your second row
    ### index2_fontcolor is the font color for your index headers for your second row
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### clean_header will give your columns title format names (ex: Birth Date) instead of underscore (birth_date) or CamelCase (BirthDate)
    ### merge_cells will merge the cells of the first header row
    ####    this MUST be used if you are not using to_excel to import data!
    ### text_wrap will wrap the column header labels for the second header row
    
    from utility_functions import clean_header_string, return_divisible_ints

    # raise an error if the header_offset input is not valid
    if isinstance(header_offset, int) == False:
        raise TypeError(f"{header_offset} is not a valid argument for header_offset. header_offset must be an integer.")
    else:
        pass

    # raise an error if the column_offset input is not valid
    if isinstance(column_offset, int) == False:
        raise TypeError(f"{column_offset} is not a valid argument for column_offset. column_offset must be an integer.")
    else:
        pass

    # raise an error if the merge_cells input is not valid
    if merge_cells == True:
        pass
    elif merge_cells == False:
        pass
    else:
        raise ValueError(f"{merge_cells} is not a valid merge_cells option. Valid arguments are True, False.")

    # raise an error if the text_wrap input is not valid    
    if text_wrap == True:
        pass
    elif text_wrap == False:
        pass
    else:
        raise ValueError(f"{text_wrap} is not a valid text_wrap option. Valid arguments are True, False.")

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1  

    # error if there are not 2 column header rows
    if num_col_indices == 2:
        pass 
    else:
        raise Exception(f"Function is only meant for datasets with two header rows. \
            The number of header rows your data has is {num_col_indices}.")

    # getting count of number of row indices to set range for index formatting
    
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # create format templates
    ## the 'last' format templates apply a right border to the last column of the second header row before the columns start repeating again
    ## and to the last index column before the data columns start

    header1_format = wb.add_format({'bold':True,'bg_color':header1_bgcolor,'font_color':header1_fontcolor,'align':'center',\
        'right':True})
    
    if text_wrap == True:
        header2_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True, 'text_wrap':True,'valign':'vcenter'})
        header2_last_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True, 'right':True, 'text_wrap':True,'valign':'vcenter'})
    else: 
        header2_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True})
        header2_last_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True, 'right':True})
  
    index1_format = wb.add_format({'bg_color':index1_bgcolor})
    index1_last_format = wb.add_format({'bg_color':index1_bgcolor,'right':True})

    if text_wrap == True:
        index2_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'bottom':True,\
            'valign':'vcenter','text_wrap':True})
        index2_last_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'right':True,\
            'bottom':True,'valign':'vcenter','text_wrap':True})
    else:
        index2_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'bottom':True,\
            'valign':'vcenter'})
        index2_last_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'right':True,\
            'bottom':True,'valign':'vcenter'})

    # optional clean header labels
    
    # if clean_header option is enabled:
    if clean_header == True:
        # create a list of column names
        col_list1 = [col_name for col_name in df.columns.get_level_values(0).unique()]
        col_list2 = [col_name for col_name in df.columns.get_level_values(1)]
        # iterate through col_names and apply clean_header_string function
        fixed_col_names1 = [clean_header_string(col_name) for col_name in col_list1]
        fixed_col_names2 = [clean_header_string(col_name) for col_name in col_list2]
        # if there are no row indices:
        if num_row_indices == 0:
            # skip this step
            pass
        # if there is a single row index:
        elif num_row_indices == 1:
            # then set the index name to the cleaned version
            fixed_index_names = [clean_header_string(df.index.name)]
        # if there is a row multiindex:
        else:
            # create a list of cleaned names
            fixed_index_names =  [clean_header_string(name) for name in df.index.names]
    # if clean_header is false:
    elif clean_header == False:
        # have a list of the regular col names
        fixed_col_names1 = [col_name for col_name in df.columns.get_level_values(0)]
        fixed_col_names2 = [col_name for col_name in df.columns.get_level_values(1)]
        # if there are no row indices:
        if num_row_indices == 0: 
            # skip this step
            pass
        # if there is a single row index
        elif num_row_indices == 1:
            # list the name of it
            fixed_index_names = [df.index.name]
        else:
            # list the name of all row indices
            fixed_index_names = [name for name in df.index.names]
    else:
        # else raise an error message that an incorrect argument has been given
        raise ValueError(f"{clean_header} is not a valid clean_header option. Valid arguments are True, False.")         

    # get values to use in formatting

    ## number of header row 1 values
    header1_n = df.columns.levshape[0]
    ## number of header row 2 values
    header2_n = df.columns.levshape[1]
    ## total number of data columns
    total_columns = header1_n * header2_n
    ## number of cells that need to be merged for header one if merge_cells = True
    cells_to_merge = header2_n - 1 
     
    # merge cells for first row of headers
    if merge_cells == True:    
        # get the columns each cell needs to start the merge on by getting divisible numbers between 0 and the number of columns
        ## divided by the number of header row 2 values
        merge_cols = return_divisible_ints(0, total_columns, header2_n)
        # drop the last (and extra) value
        merge_cols.pop()

        # iterating through our list of merge col numbers:
        for col_num in merge_cols:
            # merge the starting column to start column + cells to merge cells together on the first header row
            sheet.merge_range(header_offset, col_num + num_row_indices + column_offset, header_offset, \
                col_num + num_row_indices + column_offset + cells_to_merge,'-')
    elif header_offset != 0:
        # raise an error if the header_offset option is enabled but cells are not being merged
        raise Exception(f"Cells will needs to be merged if header_offset does not equal 0. Current header_offset = {header_offset}. Data cannot be imported with to_excel.")
    elif num_row_indices != 0:
        # else if there is a row index hide the extra row that will contain its label (when importing with to_excel())
        sheet.set_row(2,options={'hidden':True})
        print('Third row of Excel hidden to hide extra row index label when importing with to_excel.\
            If data has not been imported with to_excel, rerun code with merge_cells=True in this function.')
    else:
        # else do nothing
        pass
    
    # formatting first header row

    # iterating though the columns
    for col_num, value in enumerate(df.columns.values):
        # insert col name and apply header1 format
        sheet.write(header_offset, col_num + num_row_indices + column_offset, fixed_col_names1[col_num], header1_format)

    
    # formatting second header row
    ## interating through the columns:
    for col_num, value in enumerate(df.columns.values):
        # if the remainder of the col_num divided by the count of how many values there are = 0:
        ## the last column per header1 category will always have a remainder 0
        if (col_num + 1)%header2_n == 0:
            # apply header2_last_format and insert col name value
            ## header_offset + 1 to not overwrite header1
            sheet.write(header_offset + 1, col_num + num_row_indices + column_offset, fixed_col_names2[col_num], header2_last_format)
        else:
            # else apply regular header2_format
            ## header_offset + 1 to not overwrite header1
            sheet.write(header_offset + 1, col_num + num_row_indices + column_offset, fixed_col_names2[col_num], header2_format)

    # index formatting

    # iterating over the number of row indices present:
    for col_num in range(num_row_indices):
        # if the index is the last index in the range:
        if col_num == max(range(num_row_indices)):
            # insert the index name and apply the right border index format
            ## fixed_index_names[col_num] will retrieve the correct name based on its position in the list
            sheet.write(header_offset, col_num + column_offset, "", index1_last_format)
            ## header_offset + 1 to not overwrite index1
            sheet.write(header_offset + 1, col_num + column_offset, fixed_index_names[col_num], index2_last_format)
        else:
            # else insert the index name and apply no right border index format
            sheet.write(header_offset, col_num + column_offset, "", index1_format)
            ## header_offset + 1 to not overwrite index1
            sheet.write(header_offset + 1, col_num + column_offset, fixed_index_names[col_num], index2_format)


def last_col_highlight_header_multiindex(df, wb, sheet,  header1_bgcolor = '#002387', header1_fontcolor = '#FFFFFF', \
    header1_bghilite = '#00A111', header1_fonthilite = '#FFFFFF', header2_bgcolor = '#137A78', header2_fontcolor = '#FFFFFF',\
    index1_bgcolor =  '#002387', index2_bgcolor = '#137A78', index2_fontcolor = '#FFFFFF', \
    header_offset=0, column_offset=0, clean_header=False, merge_cells=False, text_wrap=False):

     # This function will apply formatting to your header rows and highlight the last cell of your first header row  
    ## Index is same color as normal column headers, but this can be changed if desired w/ index_color optional args
    ### Meant only for dataframes with any number of row indices and two header rows (2 level column multiindex) 

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    
    ## OPTIONAL:
    ## all color args can be added with keywords (ie, 'red') but hex codes (ex '#FF0000') are better for customization
    ### header1_bgcolor is the background color for your column headers for your first row
    ### header1_fontcolor is the font color for your column headers for your first row
    ### header1_bghilite is the background color for your last column header for your first row
    ### header1_fonthilite is the font color for your column last header for your first row
    ### header2_bgcolor is the background color for your column headers for your second row
    ### header2_fontcolor is the font color for your column headers for your second row    
    ### index1_bgcolor is the background color for your index headers for your first row
    ### index2_bgcolor is the background color for your index headers for your second row
    ### index2_fontcolor is the font color for your index headers for your second row
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### clean_header will give your columns title format names (ex: Birth Date) instead of underscore (birth_date) or CamelCase (BirthDate)
    ### merge_cells will merge the cells of the first header row
    ####    this MUST be used if you are not using to_excel to import data!
    ### text_wrap will wrap the column header labels for the second header row
    
    from utility_functions import clean_header_string, return_divisible_ints

    # raise an error if the header_offset input is not valid
    if isinstance(header_offset, int) == False:
        raise TypeError(f"{header_offset} is not a valid argument for header_offset. header_offset must be an integer.")
    else:
        pass

    # raise an error if the column_offset input is not valid
    if isinstance(column_offset, int) == False:
        raise TypeError(f"{column_offset} is not a valid argument for column_offset. column_offset must be an integer.")
    else:
        pass

    # raise an error if the merge_cells input is not valid
    if merge_cells == True:
        pass
    elif merge_cells == False:
        pass
    else:
        raise ValueError(f"{merge_cells} is not a valid merge_cells option. Valid arguments are True, False.")

    # raise an error if the text_wrap input is not valid    
    if text_wrap == True:
        pass
    elif text_wrap == False:
        pass
    else:
        raise ValueError(f"{text_wrap} is not a valid text_wrap option. Valid arguments are True, False.")

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1  

    # error if there are not 2 column header rows
    if num_col_indices == 2:
        pass 
    else:
        raise Exception(f"Function is only meant for datasets with two header rows. \
            The number of header rows your data has is {num_col_indices}.")

    # getting count of number of row indices to set range for index formatting
    
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # create format templates
    ## the 'last' format templates apply a right border to the last column of the second header row before the columns 
    ## start repeating again
    ## and to the last index column before the data columns start

    header1_format = wb.add_format({'bold':True,'bg_color':header1_bgcolor,'font_color':header1_fontcolor,'align':'center',\
        'right':True})
    header1_last_format = wb.add_format({'bold':True,'bg_color':header1_bghilite,'font_color':header1_fonthilite,'align':'center',\
        'right':True})
    
    if text_wrap == True:
        header2_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True, 'text_wrap':True,'valign':'vcenter'})
        header2_last_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True, 'right':True, 'text_wrap':True,'valign':'vcenter'})
    else: 
        header2_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True})
        header2_last_format = wb.add_format({'bold':True,'bg_color':header2_bgcolor,'font_color':header2_fontcolor,'align':'center',\
            'bottom':True, 'right':True})
  
    index1_format = wb.add_format({'bg_color':index1_bgcolor})
    index1_last_format = wb.add_format({'bg_color':index1_bgcolor,'right':True})
    
    if text_wrap==True:
        index2_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'bottom':True,\
            'valign':'vcenter','text_wrap':True})
        index2_last_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'right':True,\
            'bottom':True,'valign':'vcenter','text_wrap':True})
    else:
        index2_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'bottom':True,\
            'valign':'vcenter'})
        index2_last_format = wb.add_format({'bold':True,'bg_color':index2_bgcolor,'font_color':index2_fontcolor,'right':True,\
            'bottom':True,'valign':'vcenter'})

    # optional clean header labels
    
    # if clean_header option is enabled:
    if clean_header == True:
        # create a list of column names
        col_list1 = [col_name for col_name in df.columns.get_level_values(0).unique()]
        col_list2 = [col_name for col_name in df.columns.get_level_values(1)]
        # iterate through col_names and apply clean_header_string function
        fixed_col_names1 = [clean_header_string(col_name) for col_name in col_list1]
        fixed_col_names2 = [clean_header_string(col_name) for col_name in col_list2]
        # if there are no row indices:
        if num_row_indices == 0:
            # skip this step
            pass
        # if there is a single row index:
        elif num_row_indices == 1:
            # then set the index name to the cleaned version
            fixed_index_names = [clean_header_string(df.index.name)]
        # if there is a row multiindex:
        else:
            # create a list of cleaned names
            fixed_index_names =  [clean_header_string(name) for name in df.index.names]
    # if clean_header is false:
    elif clean_header == False:
        # have a list of the regular col names
        fixed_col_names1 = [col_name for col_name in df.columns.get_level_values(0)]
        fixed_col_names2 = [col_name for col_name in df.columns.get_level_values(1)]
        # if there are no row indices:
        if num_row_indices == 0: 
            # skip this step
            pass
        # if there is a single row index
        elif num_row_indices == 1:
            # list the name of it
            fixed_index_names = [df.index.name]
        else:
            # list the name of all row indices
            fixed_index_names = [name for name in df.index.names]
    else:
        # else raise an error message that an incorrect argument has been given
        raise ValueError(f"{clean_header} is not a valid clean_header option. Valid arguments are True, False.")         

    # get values to use in formatting

    ## number of header row 1 values
    header1_n = df.columns.levshape[0]
    ## number of header row 2 values
    header2_n = df.columns.levshape[1]
    ## total number of data columns
    total_columns = header1_n * header2_n
    ## number of cells that need to be merged for header one if merge_cells = True
    cells_to_merge = header2_n - 1 
    ## get the number of the column that is the first column of the last level
    ### we will need this to reference the latest time period
    first_last_level = total_columns - (header2_n - 1)
     
    # merge cells for first row of headers
    if merge_cells == True:    
        # get the columns each cell needs to start the merge on by getting divisible numbers between 0 and the number of columns
        ## divided by the number of header row 2 values
        merge_cols = return_divisible_ints(0, total_columns, header2_n)
        # drop the last (and extra) value
        merge_cols.pop()

        # iterating through our list of merge col numbers:
        for col_num in merge_cols:
            # merge the starting column to start column + cells to merge cells together on the first header row
            sheet.merge_range(header_offset, col_num + num_row_indices + column_offset, header_offset, \
                col_num + num_row_indices + column_offset + cells_to_merge,'-')
    elif header_offset != 0:
        # raise an error if the header_offset option is enabled but cells are not being merged
        raise Exception(f"Cells will needs to be merged if header_offset does not equal 0. \
            Current header_offset = {header_offset}. Data cannot be imported with to_excel.")
    elif num_row_indices != 0:
        # else if there is a row index hide the extra row that will contain its label (when importing with to_excel())
        sheet.set_row(2,options={'hidden':True})
        print('Third row of Excel hidden to hide extra row index label when importing with to_excel.\
            If data has not been imported with to_excel, rerun code with merge_cells=True in this function.')
    else:
        # else do nothing
        pass
    
    # formatting first header row

    # iterating though the columns
    for col_num, value in enumerate(df.columns.values):
        if col_num + 1 == first_last_level:
            sheet.write(header_offset, col_num + num_row_indices + column_offset, fixed_col_names1[col_num], header1_last_format)
        else:
            # insert col name and apply header1 format
            sheet.write(header_offset, col_num + num_row_indices + column_offset, fixed_col_names1[col_num], header1_format)

    
    # formatting second header row
    ## interating through the columns:
    for col_num, value in enumerate(df.columns.values):
        # if the remainder of the col_num divided by the count of how many values there are = 0:
        ## the last column per header1 category will always have a remainder 0
        if (col_num + 1)%header2_n == 0:
            # apply header2_last_format and insert col name value
            ## header_offset + 1 to not overwrite header1
            sheet.write(header_offset + 1, col_num + num_row_indices + column_offset, fixed_col_names2[col_num], header2_last_format)
        else:
            # else apply regular header2_format
            ## header_offset + 1 to not overwrite header1
            sheet.write(header_offset + 1, col_num + num_row_indices + column_offset, fixed_col_names2[col_num], header2_format)

    # index formatting

    # iterating over the number of row indices present:
    for col_num in range(num_row_indices):
        # if the index is the last index in the range:
        if col_num == max(range(num_row_indices)):
            # insert the index name and apply the right border index format
            ## fixed_index_names[col_num] will retrieve the correct name based on its position in the list
            sheet.write(header_offset, col_num + column_offset, "", index1_last_format)
            ## header_offset + 1 to not overwrite index1
            sheet.write(header_offset + 1, col_num + column_offset, fixed_index_names[col_num], index2_last_format)
        else:
            # else insert the index name and apply no right border index format
            sheet.write(header_offset, col_num + column_offset, "", index1_format)
            ## header_offset + 1 to not overwrite index1
            sheet.write(header_offset + 1, col_num + column_offset, fixed_index_names[col_num], index2_format)


######################## INDEX FORMATTING ##################################


###                 SINGLE ROW INDEX AND ANY NUMBER OF COLUMN LEVELS DATAFRAMES                 ###

def format_index(df, wb, sheet, header_offset=0, column_offset=0, set_width=True, text_wrap=False, wrap_rows=2):

    # This function will apply formatting to your index to bold it and give a right border
    ## Meant only for dataframes with single row index and any number of column levels

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### set_width automatically sets your index widthl when True. defaults to True
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True

    from math import ceil

    # if there is no index set raise error
    if None in df.index.names:
        raise Exception("No index set for dataframe.")
    else:
        pass

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1   

    # create index format
    index_format = wb.add_format({'bold':True,'right':True})

    ## this iterates through the rows.  this prevents the formatting being applied to empty cells
    ## it applies formatting with the index value for the first column of the report
    ## enumerate is called on the index to get those values
    for row_num, value in enumerate(df.index.values):
        # 1 is added to row num so that we don't start on 0 and overwrite our header!
        # the column is hard-coded to 0 (column A) as this is the only column we want this applied to
        sheet.write(row_num + num_col_indices + header_offset, column_offset, value, index_format)

    
    if set_width == True:
        # gets the length of all the values in the index
        index_values = [len(value) for row_num, value in enumerate(df.index.values)]

        # index name length
        if text_wrap == True and " " in df.index.name:
            name_length = ceil(len(df.index.name)/wrap_rows)
        else: 
            name_length = len(df.index.name)

        # gets the max of the index values or the name of the index, whichever is greater
        ## + 1 for 'wiggle room'
        max_index_length = max(max(index_values), name_length) + 1

        # set index column width
        sheet.set_column(column_offset, column_offset, max_index_length)
    else:
        pass


def highlight_last_index(df, wb, sheet, index_bgcolor='#002387', index_fontcolor='FFFFFF', \
    hilite_bgcolor='#00A111', hilite_fontcolor='FFFFFF', header_offset=0, column_offset=0, set_width=True, \
    text_wrap=False, wrap_rows=2):

    # This function will apply formatting to your index to bold it and give a right border and bottom borders
    ## It will fill one color for all your index row backgrounds and a different color for your last index row value as a highlight
    ## Meant only for dataframes with single row index and any number of column levels

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### index_bgcolor is the background color of your index rows
    ### index_fontcolor is the font color of your index rows
    ### hilite_bgcolor is the background color of your last index row
    ### hilite_fontcolor is the font color of your last index row
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### set_width automatically sets your index width when True. defaults to True
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True

    from math import ceil

    # if there is no index set raise error
    if None in df.index.names:
        raise Exception("No index set for dataframe.")
    else:
        pass

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1   
    
    # index formats
    index_format = wb.add_format({'bold':True,'bg_color':index_bgcolor,'font_color':index_fontcolor,'right':True,\
        'bottom':True})
    last_index_format = wb.add_format({'bold':True,'bg_color':hilite_bgcolor,'font_color':hilite_fontcolor,'right':True,\
        'bottom':True})

    # getting basic parameters to use in functions
    ## count of index values
    index_value_count = df.index.get_level_values(0).nunique()

    # formatting row index (periods)

    # iterating over row index values:
    for row_num, value in enumerate(df.index.values):
        # if the remainder of row_num divided by count of index values is 0 (aka it is the last row)
        ## row_num + 1 here since python starts from 0
        if (row_num+1)%index_value_count == 0:
            # apply latest period row format
            ## row_num + 1 here is *different* from the row_num + 1 above--it's so we don't overwrite our 1 header row
            ### if we had 2 header rows it would be row_num + 2, etc
            sheet.write(row_num + num_col_indices + header_offset, column_offset, value, last_index_format)
        else:
            # else apply period row format
            sheet.write(row_num + num_col_indices + header_offset, column_offset, value, index_format)  

    if set_width == True:
        # gets the length of all the values in the index
        index_values = [len(value) for row_num, value in enumerate(df.index.values)]

        # get length of index name
        if text_wrap == True and " " in df.index.name:
            name_width = ceil(len(df.index.name)/wrap_rows)
        else:
            name_width = len(df.index.name)

        # gets the max of the index values or the name of the index, whichever is greater
        ## + 1 for 'wiggle room'
        max_index_length = max(max(index_values), name_width) + 1

        # set index column width
        sheet.set_column(column_offset, column_offset, max_index_length) 
    else:
        pass
    

###                ROW MULTIINDEX AND ANY NUMBER OF COLUMN LEVELS DATAFRAMES                 ###

def merge_row_index_cells(df, wb, sheet, header_offset=0, column_offset=0):

    from utility_functions import return_divisible_ints

    # This function will merge the cells in your index columns that are from the same category
    ## Can be used on any row multiindex dataframe
    ### NOTE: will break if not all index categories are present in each index!
    ### NOTE: will also break if row indices are not arranged in least categories to most categories order (which is pandas standard)
    ### This function will need to be used if creating a row with row multiindex data and not using to_excel() to import

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0    

    #getting count of row_indices

    # if there is no index set raise error
    if None in df.index.names:
        raise Exception("No index set for dataframe.")
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # exit function with error if it is not a multiindex
    if num_row_indices == 1:
        raise Exception("Function is not meant for single row index datasets.")
    else:
        pass

    # determining how many rows are in the major (leftmost) index by dividing the total row count by index[0] unique values
    rows_per_major_index = int(len(df)/len(df.index.unique(0)))

    # getting the count of categories per index
    
    # create an empty list to hold the values:
    cat_counts = []

    # iterating over our row indices:
    for col_num in range(num_row_indices):
        # get count of unique values
        cat_count = len(df.index.unique(col_num))
        # append them to list
        cat_counts.append(cat_count)

    # get the count of rows per category for each index

    # create a empty list to hold the values
    cat_row_counts = []

    # iterating through the number of row indices we have:
    for col_num in range(num_row_indices):
        # get the category count of the index
        cat_count = len(df.index.unique(col_num))
        # if it is the major index[0]:
        if col_num == 0:
            # rows_per_cat is the rows_per_major_index
            rows_per_cat = rows_per_major_index
        else:
            # else rows_per_cat of last index divided by the category count of current index
            rows_per_cat = int(cat_row_counts[col_num-1]/cat_counts[col_num])
        # append rows_per_cat value to list
        cat_row_counts.append(rows_per_cat)
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1

    # get the number of rows in the data
    data_rows = len(df)

    # need a list of rows to start cell merges on
    
    # iterating over our numbers of cells to merge per index:
    for col_num, merge_n in enumerate(cat_row_counts):        
        #skip if there are no cells to merge
        if merge_n == 0 or merge_n == 1:
            pass
        else:
            # create a list using return_divisible_ints with 0 as the start_num and our count of data rows as the end_num of range, 
            # divided by merge_n
            divisible_rows = [i for i in return_divisible_ints(0, data_rows, merge_n)]
            # will return 1 too many values--drop the last one
            divisible_rows.pop()
            for row_num in divisible_rows:
                # merge cells     starting cell is row_num + num_col_indices + header_offset
                sheet.merge_range(row_num + num_col_indices + header_offset,
                          # the index column
                          col_num + column_offset,
                          # row_num + header rows + amount of cells to merge - 1 for our ending cell
                          ## -1 because the row_num cell is already accounted for
                          row_num + num_col_indices + header_offset + merge_n - 1,
                          # the index column
                          col_num + column_offset,
                          # message to fill in which will warn user if they forget to import index labels in subsequent steps
                          'Forgot to Import Data!')


def format_row_multiindex(df, wb, sheet, header_offset=0, column_offset=0, set_width=True, text_wrap=False, wrap_rows=2):

    # This function will apply formatting to your index to bold it and give a right border
    ## Meant only for dataframes with row mulitiindex and and number of columns levels
    ### NOTE: will break if not all index categories are present in each index!
    ### NOTE: will also break if row indices are not arranged in least categories to most categories order (which is pandas standard)
    ### if you are not importing with to_excel(), the merge_row_index_cells() function must be applied first

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### set_width automatically sets your index width when True. defaults to True    
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True

    from math import ceil

    #getting count of row_indices
    # if there is no index set raise error
    if None in df.index.names:
        raise Exception("No index set for dataframe.")
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # getting rows per category in first index
    rows_per_major_index = int(len(df)/len(df.index.unique(0)))

    # exit function with error if it is not a multiindex
    if num_row_indices == 1:
        raise Exception("Function is not meant for single row index datasets.")
    else:
        pass

    # getting the count of categories per index:
    
    # create an empty list to hold the values:
    cat_counts = []

    # iterating over our row indices:
    for col_num in range(num_row_indices):
        # get count of unique values
        cat_count = len(df.index.unique(col_num))
        # append them to list
        cat_counts.append(cat_count)
        
    # create a empty list to hold the values
    cat_row_counts = []

    # iterating through the number of row indices we have:
    for col_num in range(num_row_indices):
        # get the category count of the index
        cat_count = len(df.index.unique(col_num))
        # if it is the major index[0]:
        if col_num == 0:
            # rows_per_cat is the rows_per_major_index
            rows_per_cat = rows_per_major_index
        else:
            # else rows_per_cat of last index divided by the category count of current index
            rows_per_cat = int(cat_row_counts[col_num-1]/cat_counts[col_num])
        # append rows_per_cat value to list
        cat_row_counts.append(rows_per_cat)

   
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1  

     
    # creating formats
    index_format = wb.add_format({'bold':True,'valign':'vcenter'})
    index_bottom_row_format = wb.add_format({'bold':True,'valign':'vcenter','bottom':True})
    last_index_format = wb.add_format({'bold':True,'valign':'vcenter','right':True})
    last_index_bottom_format = wb.add_format({'bold':True,'valign':'vcenter','bottom':True,'right':True})
    
    # iterating over our indices:
    for col_num in range(num_row_indices):
        # if it is the first (major) index:
        if col_num == 0:
            # iterating over the values in the index:
            for row_num, value in enumerate(df.index.get_level_values(col_num)):
                # insert index value and apply bottom border index format to all cells
                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value, index_bottom_row_format)
        # if it is the last (one row per category) index:
        elif col_num == max(range(num_row_indices)):
            # raise an error if there is more than one row per each value
            if cat_row_counts[col_num] != 1:
                raise Exception('Your final index has more than one row per each value.')
            else:
            # iterating over the values in the index:
                for row_num, value in enumerate(df.index.get_level_values(col_num)):
                    # if it is the last row before a new major index category:
                    if (row_num+1)%rows_per_major_index==0:
                        # apply the last index bottom format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value, last_index_bottom_format)
                    else:
                         # apply last index format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value, last_index_format) 
        else:        
            # for all other indices iterate over index values:
            for row_num, value in enumerate(df.index.get_level_values(col_num)):
                # if it is the last row in the index category:
                ## as determined by if the row number is divisible by the number of categories times the rows per category
                if (row_num+1)%(cat_counts[col_num]*cat_row_counts[col_num])==0:
                    # insert index value and apply bottom border index formatting
                    sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value, index_bottom_row_format)
                else:
                    # else insert index value and apply no border index formating
                     sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value, index_format)



    # set index column widths 
    if set_width == True:
        # create empty list to hold max_index_lengths
        max_index_lengths = []

        # iterating over row indices:
        for col_num in range(num_row_indices):
            # store the length of all index values in a list
            index_values = [len(value) for i, value in enumerate(df.index.get_level_values(col_num))]
            # get index name length
            if text_wrap == True and " " in df.index.names[col_num]:
                name_width = ceil(len(df.index.names[col_num])/wrap_rows)
            else:
                name_width = len(df.index.names[col_num])
            # get the max width of the longest value or title, whichever is longer
            ## + 1 for 'wiggle room'
            max_index_length = max(max(index_values), name_width) + 1
            # add that to the max_index_lengths list
            max_index_lengths.append(max_index_length)

        # iterating over row indices again:
        for col_num in range(num_row_indices):
            # set width to matching max index length
            sheet.set_column(col_num + column_offset, col_num + column_offset, max_index_lengths[col_num])
    else: 
        pass

    


######################## DATA FORMATTING ##################################


###                ANY NUMBER ROW INDICES AND SINGLE COLUMNS INDEX DATAFRAMES                 ###


def set_col_width(df, wb, sheet, col_name, method='headers', column_offset=0, text_wrap=False, wrap_rows=2):

    # adapted from a solution from dfresh22 found at 
    # https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter

    # This function will automatically make specified column wide enough for their full column names to appear without being cut off
    ## Can be used for width based on data or data and header though
    ## Meant for use on data with only one index of columns, but any number of row indices

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### method is how the width is set:
    #       'header' sets width based on the length of column names. is default
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True

    # error if entered col_name not in dataframe

    from math import ceil

    # create list of all col_names
    col_name_list = [col_name for col_name in df.columns]

    if col_name not in col_name_list:
        raise ValueError(f"{col_name} not in dataframe. Columns in data are: {col_name_list}")
    else:
        pass

    # list of valid method args
    valid_methods = ['headers', 'data', 'all']

    # error if valid method arg not used
    if method not in valid_methods:
        raise ValueError(f"{method} is not a valid method option. Valid methods are: {valid_methods}")
    else:
        pass

    # create an object holding the length of the name of the column
    ## + 1 for 'wiggle room'
    if text_wrap == True and " " in df[col_name].name:
        col_name_length = ceil(len(df[col_name].name)/wrap_rows) + 1
    else:
        col_name_length = len(df[col_name].name) + 1

    # getting length of longest data point

    # list of all column values
    values = df[col_name].tolist()
    # create empty list to store the lengths
    value_lengths = []
    # iterating over the values list:
    for row_num, value in enumerate(values):
            # get the length in characters of each value
            length = len(str(value))
            # add it to the value_lengths list
            value_lengths.append(length)
            if row_num + 1 == len(values):
                # get the max width value
                ## + 1 for 'wiggle room'
                max_data_width = max(value_lengths) + 1

    # get max of headers and data
    max_all_lengths = max(col_name_length, max_data_width)

    if method == 'headers':
        col_width = col_name_length
    elif method == 'data':
        col_width = max_data_width
    elif method == 'all':
        col_width = max_all_lengths

    # get the count of how many row indices they are so we can skip those columns in the for loop
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    for col_num, df_col_name in enumerate(df.columns):
        # if the specified column name matches 
        if df_col_name == col_name:    
            sheet.set_column(col_num + num_row_indices + column_offset, col_num + num_row_indices + column_offset, col_width)  
        else:
            pass



###                 SINGLE ROW INDEX AND SINGLE COLUMNS INDEX DATAFRAMES                 ###


def insert_data(df, wb, sheet, header_offset=0, column_offset=0, data_type=None, replace_nulls=True, null_value='-', \
    null_align='center'):
    
    # This function will insert your data in desired cells with a header_offset
    ## Can be used on any dataframe
    

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### data_type is the type of numeric data:
    #### this arg should only be used if all your data is the same data type!
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    import pandas as pd

    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
    elif data_type == 'decimal_2':
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
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
    elif data_type == 'text':
        raise Exception('Data types are text by default! Function not needed.')
    elif data_type == None:
        pass
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # getting the column count

    # getting the count of row index columns
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
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

    # iterating over data columns excluding row index columns:
    for col_num in range(num_row_indices, total_cols):
        # iterating over rows containing data:
        for row_num, value in enumerate(df.values):
            # no data_type is specified:
            if data_type == None:
                # if data is null
                if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                    # insert null value and apply null formatting
                    sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_format)
                else:
                    # insert the data into the cell matching the postion in the datatframe
                    ## value[] has num_row_indices subtracted from it for indexing since that was added to the col_num in range()
                    sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value[col_num-num_row_indices])
            else:
                # if data is null
                if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                    # insert null value and apply null formatting
                    sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_format)
                else:
                    # insert the data into the cell and apply specified formatting
                    sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value[col_num-num_row_indices],\
                        data_format)


def format_single_data_type_df(df, wb, sheet, data_type, col_width=14, col_width_method=None, column_offset=0, \
    text_wrap=False, wrap_rows=2):

    # This function will apply the specified numeric formatting to all data columns
    ## Meant only for dataframes that have the same data type for ALL non-index columns, but can have any number of columns and indices
    ### Note: this will set ALL data columns to the same width!

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### data_type is the type of data:
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
        
    ## OPTIONAL:
    ### col_width is the width of the data columns. defaults to 14
    ### coL_width_method is how the width is set. defaults to None, which itself defaults to the default col_width_num (14):
    #       'header' sets width based on the length of column names
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True

    import numpy as np
    from math import ceil

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
    elif data_type == 'decimal_2':
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
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
    elif data_type == 'text':
        raise Exception('Data types are text by default! Function not needed.')
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")


    # create list of all valid methods
    valid_methods = ['headers', 'data', 'all']

    # error if col_width_method not valid
    if col_width_method == None:
        pass
    elif col_width_method not in valid_methods:
        raise ValueError(f"{col_width_method} is not a valid method option, Valid methods are None or: {valid_methods}")
    else:
        pass

    # create a list holding the length of the name of each column
    ## + 1 for 'wiggle room'
    if text_wrap == True:
        col_names = [name for name in df.columns]
        col_name_lengths = []
        for name in col_names:
            if " " in name:
                length = ceil(len(name)/wrap_rows) + 1
            else:
                length = len(name) + 1
            col_name_lengths.append(length)
    else:
        col_name_lengths = [len(name) + 1 for name in df.columns]

    # create a list holding the max length of the data in each column
    max_data_lengths = []
        
    # iterating over the data columns:    
    for col in list(df):
        # store their values in a list
        values = df[col].tolist()
        # create an empty list to store the lengths
        value_lengths = []
        # iterating over the values list:
        for row_num, value in enumerate(values):
            # get the length in characters of each value
            length = len(str(value))
            # add it to the value_lengths list
            value_lengths.append(length)
            # if it is the final iteration over the values with the completed value_lengths list for the column:
            ## + 1 since python numbering starts at 0
            if row_num + 1 == len(values):
                # get the max width value
                ## + 1 for 'wiggle room'
                max_data_width = max(value_lengths) + 1
                # append it to our data lengths list
                max_data_lengths.append(max_data_width)

    # create a list for the max of data and column width, whichever is greater
    max_all_lengths = np.maximum(col_name_lengths, max_data_lengths)

    col_width_num_list = [col_width for col_num in df.columns]

    if col_width_method == 'headers':
        width_list = col_name_lengths
    elif col_width_method == 'data':
        width_list = max_data_lengths
    elif col_width_method == 'all':
        width_list = max_all_lengths
    else:
        width_list = col_width_num_list

    # getting column count of the data to use to set upper bound for formatting
    df_column_count = len(df.columns)
    
    # getting row indices count of the data to use to set lower bound for formatting
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    ## sets columns B through the last column present in the dataset with the specified data_format and and sets column widths
    for col_num, width in enumerate(width_list):    
        sheet.set_column(col_num + num_row_indices + column_offset, col_num + df_column_count + column_offset, width, data_format)


def set_col_data_type(df, wb, sheet, col_name, data_type, col_width_method=None, col_width_num=14, column_offset=0, \
    text_wrap=False, wrap_rows=2):

    # This function will apply the specified formatting to the specified column
    ## Can work on dataframes with single row index and single row of column headers
    ### Note: date formatting will only apply correctly to datetime columns

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### col_name is the name of your column
    ### data_type is the type of data:
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    

        
    ## OPTIONAL:
    ### col_width_num is the width of the data columns. defaults to 14
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### coL_width_method is how the width is set. defaults to None, which itself defaults to the default col_width_num (14):
    #       'header' sets width based on the length of column names
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True
    
    from math import ceil

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
    elif data_type == 'decimal_2':
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
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
    elif data_type == 'text':
        raise Exception('Data types are text by default! Function not needed.')
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # error if entered col_name not in dataframe

    # create list of all col_names
    col_name_list = [col_name for col_name in df.columns]

    if col_name not in col_name_list:
        raise ValueError(f"{col_name} not in dataframe. Columns in data are: {col_name_list}")
    else:
        pass

    
    # create list of all valid methods
    valid_methods = ['headers', 'data', 'all']

    # error if col_width_method not valid
    if col_width_method == None:
        pass
    elif col_width_method not in valid_methods:
        raise ValueError(f"{col_width_method} is not a valid method option, Valid methods are None or: {valid_methods}")
    else:
        pass

    # setting col_width

    # create an object holding the length of the name of the column
    ## + 1 for 'wiggle room'
    if text_wrap == True and " " in df[col_name].name:
        col_name_length = ceil(len(df[col_name])/wrap_rows) + 1
    else:
        col_name_length = len(df[col_name].name) + 1

    # getting length of longest data point

    # list of all column values
    values = df[col_name].tolist()
    # create empty list to store the lengths
    value_lengths = []
    # iterating over the values list:
    for row_num, value in enumerate(values):
            # get the length in characters of each value
            length = len(str(value))
            # add it to the value_lengths list
            value_lengths.append(length)
            if row_num + 1 == len(values):
                # get the max width value
                ## + 1 for 'wiggle room'
                max_data_width = max(value_lengths) + 1

    # get max of headers and data
    max_all_lengths = max(col_name_length, max_data_width)

    if col_width_method == 'headers':
        col_width = col_name_length
    elif col_width_method == 'data':
        col_width = max_data_width
    elif col_width_method == 'all':
        col_width = max_all_lengths
    else:
        col_width = col_width_num

    # getting row indices count of the data to use to set lower bound for formatting
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)
        
    # iterate through columns until we get to the selected column:
    for col_num, df_col_name in enumerate(df.columns):
        # if the specified column name matches 
        if df_col_name == col_name:
            sheet.set_column(col_num + num_row_indices + column_offset, col_num + num_row_indices + column_offset, col_width,\
                data_format)
        else:
            pass


###                 ANY NUMBER ROW INDEX AND SINGLE COLUMNS INDEX DATAFRAMES                 ###


def set_column_widths(df, wb, sheet, column_offset=0, method='headers', text_wrap=False, wrap_rows=2):

    import numpy as np

    # adapted from a solution from dfresh22 found at 
    # https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter

    # This function will automatically make all columns wide enough for their full column names to appear without being cut off
    ## Can be used for width based on data or data and header though
    ## Meant for use on data with only one index of columns, but any number of row indices

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### method is how the width is set:
    #       'header' sets width based on the length of column names. is default
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### text_wrap specifies if you index headers were wrapped when applying header formatting. default is False
    ### wrap_rows is how many rows wide the wrapped header text should be. default is 2
    ####        should be used if text_wrap is True
    
    from math import ceil

    # list of valid method args
    valid_methods = ['headers', 'data', 'all']

    # error if valid method arg not used
    if method not in valid_methods:
        raise ValueError(f"{method} is not a valid method option. Valid methods are: {valid_methods}")
    else:
        pass
    

    # create a list holding the length of the name of each column
    ## + 1 for 'wiggle room'
    if text_wrap == True:
        col_name_lengths = []
        for name in df.columns:
            if " " in name:
                length = ceil(len(name)/wrap_rows) + 1
            else:
                length = len(name) + 1
            col_name_lengths.append(length)
    else:
        col_name_lengths = [len(name) + 1 for name in df.columns]

    # create a list holding the max length of the data in each column
    max_data_lengths = []
        
    # iterating over the data columns:    
    for col in list(df):
        # store their values in a list
        values = df[col].tolist()
        # create an empty list to store the lengths
        value_lengths = []
        # iterating over the values list:
        for row_num, value in enumerate(values):
            # get the length in characters of each value
            length = len(str(value))
            # add it to the value_lengths list
            value_lengths.append(length)
            # if it is the final iteration over the values with the completed value_lengths list for the column:
            ## + 1 since python numbering starts at 0
            if row_num + 1 == len(values):
                # get the max width value
                ## + 1 for 'wiggle room'
                max_data_width = max(value_lengths) + 1
                # append it to our data lengths list
                max_data_lengths.append(max_data_width)   

    # create a list for the max of data and column width, whichever is greater
    max_all_lengths = np.maximum(col_name_lengths, max_data_lengths)
    
    # get the count of how many row indices they are so we can skip those columns in the for loop
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # choosing list to use based on method:
    if method == 'headers':
        width_list = col_name_lengths
    elif method == 'data':
        width_list = max_data_lengths
    elif method ==  'all':
        width_list = max_all_lengths 

    # iterating over the df columns:
    for col_num, width in enumerate(width_list):
        # apply the matching width to the column
        sheet.set_column(col_num + num_row_indices + column_offset, col_num + num_row_indices + column_offset, width)  


###                 ROW MULTIINDEX AND SINGLE COLUMNS INDEX DATAFRAMES                 ###

def insert_row_multiindex_data(df, wb, sheet, header_offset=0, column_offset=0, data_type=None, replace_nulls=True, null_value='-', \
    null_align='center'):

    # This function will insert your data in desired cells and underline the last row per major index category
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### data_type is the type of numeric data:
    #### this arg should only be used if all your data is the same data type!
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    import pandas as pd

    #getting count of row_indices
    # if there is no index raise error
    if None in df.index.names:
        raise Exception("No index set on dataframe.")
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # exit function with error if it is not a multiindex
    if num_row_indices == 1:
        raise Exception("Function is not meant for single row index datasets.")
    else:
        pass

    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'right'})

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0','bottom':True})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.0','bottom':True})
    elif data_type == 'decimal_2':
        data_format = wb.add_format({'num_format':'#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.00','bottom':True})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0','bottom':True})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0.00','bottom':True})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
        data_bottom_format = wb.add_format({'num_format':'0%','bottom':True})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
        data_bottom_format = wb.add_format({'num_format':'0.0%','bottom':True})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
        data_bottom_format = wb.add_format({'num_format':'0.00%','bottom':True})
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd','bottom':True})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy','bottom':True})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','bottom':True})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','bottom':True})
    elif data_type == 'text':
        print("Data types are text by default. No error, continuing function.")
        data_bottom_format = wb.add_format({'bottom':True})
    elif data_type == None:
        data_bottom_format = wb.add_format({'bottom':True})
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # determining how many rows are in the major (leftmost) index by dividing the total row count by index[0] unique values
    rows_per_major_index = int(len(df)/len(df.index.unique(0)))

    # getting the column count

    # getting the count of regular columns
    num_cols = len(df.columns)
    # adding num_cols and number of row indices together for total column count
    total_cols = num_row_indices + num_cols

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1
    

    # iterating over data columns excluding row index columns:
    for col_num in range(num_row_indices, total_cols):
        # iterating over rows containing data:
        for row_num, value in enumerate(df.values):
            # if no data type is assigned:
            if data_type == None or data_type == 'text':
                # for the last row per first index category:
                if (row_num + 1)%rows_per_major_index==0:
                    # we check if the value of the cell is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply bottom border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_bottom_format)
                    else:
                        # insert data with a bottom border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value[col_num-num_row_indices],\
                            data_bottom_format)
                else:
                    # we check if the value of the cell is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply null format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_format)
                    else:
                        # insert data with no formatting
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value[col_num-num_row_indices])
            else:
                # else for the last row per first index category:
                if (row_num + 1)%rows_per_major_index==0:
                    # we check if the value of the cell is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply bottom border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_bottom_format)    
                    else:
                        # insert the data and apply the specified formatting with bottom border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                            value[col_num-num_row_indices], data_bottom_format)
                else:
                    # we check if the value of the cell is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply null format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_format)
                    else:
                        # insert the data and apply specified formatting (no border)
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                            value[col_num-num_row_indices], data_format)


def set_row_multiindex_col_dtype(df, wb, sheet, col_name, data_type, column_offset=0, header_offset=0, replace_nulls=True, \
    null_value='-', null_align='center'):

    # This function will apply the specified formatting to the specified column
    ## Can work on dataframes with a row multiindex and any number of column header levels
    ### Note: date formatting will only apply correctly to datetime columns

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### col_name is the name of your column
    ### data_type is the type of data:
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    
            
    ## OPTIONAL:
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    #getting count of row_indices
    # if there is no index raise error
    if None in df.index.names:
        raise Exception("No index set on dataframe.")
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # exit function with error if it is not a multiindex
    if num_row_indices == 1:
        raise Exception("Function is not meant for single row index datasets.")
    else:
        pass

    import pandas as pd

    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'right'})

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0','bottom':True})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.0','bottom':True})
    elif data_type == 'decimal_2':
        data_format = wb.add_format({'num_format':'#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.00','bottom':True})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0','bottom':True})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0.00','bottom':True})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
        data_bottom_format = wb.add_format({'num_format':'0%','bottom':True})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
        data_bottom_format = wb.add_format({'num_format':'0.0%','bottom':True})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
        data_bottom_format = wb.add_format({'num_format':'0.00%','bottom':True})
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd','bottom':True})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy','bottom':True})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','bottom':True})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','bottom':True})
    elif data_type == 'text':
        data_bottom_format = wb.add_format({'bottom':True})
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # error if entered col_name not in dataframe

    # create list of all col_names
    col_name_list = [col_name for col_name in df.columns]

    if col_name not in col_name_list:
        raise ValueError(f"{col_name} not in dataframe. Columns in data are: {col_name_list}")
    else:
        pass

    # determining how many rows are in the major (leftmost) index by dividing the total row count by index[0] unique values
    rows_per_major_index = int(len(df)/len(df.index.unique(0)))

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1

    # iterate through columns until we get to the selected column:
    for col_num, df_col_name in enumerate(df.columns):
        # if the specified column name matches 
        if df_col_name == col_name:
            # iterating over rows containing data:
            for row_num, value in enumerate(df.values):
                if data_type == 'text':
                    # for the last row per first index category:
                    if (row_num + 1)%rows_per_major_index==0:
                        # if the value is null
                        if pd.isna(value[col_num]) and replace_nulls == True:
                            # insert null value and apply bottom formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                null_value, null_bottom_format)
                        else:
                            # insert data with a bottom border
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                value[col_num], data_bottom_format)
                    else:
                        # if the value is null
                        if pd.isna(value[col_num]) and replace_nulls == True:
                            # insert null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                null_value, null_format)
                        else:
                            # insert data with no formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                value[col_num])
                else:
                    # else for the last row per first index category:
                    if (row_num + 1)%rows_per_major_index==0:
                        # if the value is null
                        if pd.isna(value[col_num]) and replace_nulls == True:
                            # insert null value and apply bottom formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                null_value, null_bottom_format)
                        else:
                            # insert the data and apply the specified formatting with bottom border
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                value[col_num], data_bottom_format)
                    else:
                        # if the value is null
                        if pd.isna(value[col_num]) and replace_nulls == True:
                            # insert null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                null_value, null_format)
                        else:
                            # insert the data and apply specified formatting (no border)
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                value[col_num], data_format)
        else:
            pass   


###                 ANY NUMBER ROW INDICES AND TWO LEVEL COLUMN MULITINDEX DATAFRAMES                 ###


def set_multiindex_column_widths(df, wb, sheet, column_offset=0, method='headers', text_wrap=False, wrap_rows=2):

    import numpy as np
    from math import ceil
    # adapted from a solution from dfresh22 found at 
    # https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter

    # This function will automatically make all columns wide enough for their full column names to appear without being cut off
    ## Can be used for width based on data or data and header though
    ## Meant for use on data with two header rows (a two level column index), but any number of row indices

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### method is how the width is set:
    #       'header' sets width based on the length of column names. is default
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### text_wrap will wrap text in headers when True
    ### wrap_rows is how many rows wide the wrapped text should be. default is 2

    # check to make sure column_offset input is valid
    ## if column_offset is not an integer, raise an error
    if isinstance(column_offset, int) == False:
        raise TypeError(f"{column_offset} is not a valid argument for column_offset. column_offset must be an integer.")
    else:
        pass

    # check to make sure wrap_rows input is valid
    ## if wrap_rows is not an integer, raise an error
    if isinstance(wrap_rows, int) == False:
        raise TypeError(f"{wrap_rows} is not a valid argument for wrap_rows. wrap_rows must be an integer.")
    else:
        pass

    # list of valid method args
    valid_methods = ['headers', 'data', 'all']

    # error if valid method arg not used
    if method not in valid_methods:
        raise ValueError(f"{method} is not a valid method option. Valid methods are: {valid_methods}")
    else:
        pass
    

    # create a list holding the length of the name of each column
    ## + 1 for 'wiggle room'
    #col_name_lengths = [len(name) + 1 for name in df.columns.get_level_values(1)]

    # create a list holding the max length of the data in each column
    max_data_lengths = []
        
    # iterating over the data columns:    
    for col in list(df):
        # store their values in a list
        values = df[col].tolist()
        # create an empty list to store the lengths
        value_lengths = []
        # iterating over the values list:
        for row_num, value in enumerate(values):
            # get the length in characters of each value
            length = len(str(value))
            # add it to the value_lengths list
            value_lengths.append(length)
            # if it is the final iteration over the values with the completed value_lengths list for the column:
            ## + 1 since python numbering starts at 0
            if row_num + 1 == len(values):
                # get the max width value
                ## + 1 for 'wiggle room'
                max_data_width = max(value_lengths) + 1
                # append it to our data lengths list
                max_data_lengths.append(max_data_width)   

        
    # get the count of how many row indices they are so we can skip those columns in the for loop
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # adjust col widths for text wrapping of headers
    if text_wrap == True:
        # if there is text wrapping, the width of columns is their previous col_name length divided by the number of rows for wrapping 
        ## rounded up so that it is an integer value to prevent errors
        col_name_lengths = []
        for name in df.columns.get_level_values(1):
            if " " in name:
                length = ceil(len(name)/wrap_rows) + 1
            else:
                length = len(name) + 1
            col_name_lengths.append(length)
    elif text_wrap==False:
        col_name_lengths = [len(name) + 1 for name in df.columns.get_level_values(1)]
    else:
        raise ValueError(f"{text_wrap} is not not a valid text_wrap argument. text_wrap must be True or False.")

    # create a list for the max of data and column width, whichever is greater
    max_all_lengths = np.maximum(col_name_lengths, max_data_lengths)

    # choosing list to use based on method:
    if method == 'headers':
        width_list = col_name_lengths
    elif method == 'data':
        width_list = max_data_lengths
    elif method ==  'all':
        width_list = max_all_lengths 

    # iterating over the df columns:
    for col_num, width in enumerate(width_list):
        # apply the matching width to the column
        sheet.set_column(col_num + num_row_indices + column_offset, col_num + num_row_indices + column_offset, width)


###                 NO OR SINGLE ROW INDEX AND TWO LEVEL COLUMN MULITINDEX DATAFRAMES                 ###


def insert_col_multiindex_data(df, wb, sheet, header_offset=0, column_offset=0, data_type=None, replace_nulls=True, null_value='-', \
    null_align='center'):

    # This function will insert your data in desired cells and apply a right border to the last column of each first level category
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### data_type is the type of numeric data:
    #### this arg should only be used if all your data is the same data type!
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    import pandas as pd

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        raise Exception('Data does not have a columns multiindex')

    # error if there are not 2 column header rows
    if num_col_indices == 2:
        pass 
    else:
        raise Exception(f"Function is only meant for datasets with two header rows. \
            The number of header rows your data has is {num_col_indices}.")

    #getting count of row_indices
    # if there is no index set to 0
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # error if it is a row multiindex
    if num_row_indices > 1:
        raise Exception(f"Function is only meant with datasets with a single row index or no row index. \
            The number of row indices your data has is {num_row_indices}.")

    
    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
        null_right_format = wb.add_format({'right':True, 'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
        null_right_format = wb.add_format({'right':True, 'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})
        null_right_format = wb.add_format({'right':True, 'align':'right'})


    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
        data_right_format = wb.add_format({'num_format':'#,##0','right':True})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
        data_right_format = wb.add_format({'num_format':'#,##0.0','right':True})
    elif data_type == 'decimal_2':
        data_format = wb.add_format({'num_format':'#,##0.00'})
        data_right_format = wb.add_format({'num_format':'#,##0.00','right':True})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
        data_right_format = wb.add_format({'num_format':'$#,##0','right':True})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
        data_right_format = wb.add_format({'num_format':'$#,##0.00','right':True})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
        data_right_format = wb.add_format({'num_format':'0%','right':True})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
        data_right_format = wb.add_format({'num_format':'0.0%','right':True})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
        data_right_format = wb.add_format({'num_format':'0.00%','right':True})
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd','right':True})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy','right':True})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','right':True})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','right':True})
    elif data_type == 'text':
        print("Data types are text by default. No error, continuing function.")
        data_right_format = wb.add_format({'right':True})
    elif data_type == None:
        data_right_format = wb.add_format({'right':True})
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    ## number of header row 2 values which are what we need to loop over
    header2_n = df.columns.levshape[1]

    # getting the column count

    # getting the count of regular columns
    num_cols = len(df.columns)
    # adding num_cols and number of row indices together for total column count
    total_cols = num_row_indices + num_cols
        

    # iterating over data columns excluding row index columns:
    for col_num in range(num_row_indices, total_cols):
        # iterating over rows containing data:
        for row_num, value in enumerate(df.values):
            # if no data type is assigned:
            if data_type == None or data_type == 'text':
                # for the last row per first index category:
                if (col_num)%header2_n==0:
                    # if the value is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply right border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_right_format)
                    else:
                        # insert data with a bottom border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value[col_num-num_row_indices],\
                            data_right_format)
                else:
                    # if the value is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply null format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_format)
                    else:
                        # insert data with no formatting
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, value[col_num-num_row_indices])
            else:
                # else for the last row per first index category:
                if (col_num)%header2_n==0:
                    # if the value is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply right border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_right_format)
                    else:
                        # insert the data and apply the specified formatting with bottom border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                            value[col_num-num_row_indices], data_right_format)
                else:
                    # if the value is null
                    if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                        # insert null value and apply null format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value,\
                            null_format)
                    else:
                        # insert the data and apply specified formatting (no border)
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                            value[col_num-num_row_indices], data_format)
                    

def set_col_multiindex_dtype(df, wb, sheet, col_name, data_type, column_offset=0, header_offset=0, replace_nulls=True, \
    null_value='-', null_align='center'):

    # This function will apply the specified formatting to the all columns in column level 2 matching the column name
    ## Can work on dataframes with any number of row indices and two column header levels
    ### Note: date formatting will only apply correctly to datetime columns

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### col_name is the name of your column
    ### data_type is the type of data:
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)    

        
    ## OPTIONAL:
    ### col_width_num is the width of the data columns. defaults to 14
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### coL_width_method is how the width is set. defaults to None, which itself defaults to the default col_width_num (14):
    #       'header' sets width based on the length of column names
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    import pandas as pd

    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
        null_right_format = wb.add_format({'right':True, 'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
        null_right_format = wb.add_format({'right':True, 'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})
        null_right_format = wb.add_format({'right':True, 'align':'right'})

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']

     # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
        data_right_format = wb.add_format({'num_format':'#,##0','right':True})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
        data_right_format = wb.add_format({'num_format':'#,##0.0','right':True})
    elif data_type == 'decimal_2':
        data_format = wb.add_format({'num_format':'#,##0.00'})
        data_right_format = wb.add_format({'num_format':'#,##0.00','right':True})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
        data_right_format = wb.add_format({'num_format':'$#,##0','right':True})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
        data_right_format = wb.add_format({'num_format':'$#,##0.00','right':True})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
        data_right_format = wb.add_format({'num_format':'0%','right':True})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
        data_right_format = wb.add_format({'num_format':'0.0%','right':True})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
        data_right_format = wb.add_format({'num_format':'0.00%','right':True})
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd','right':True})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy','right':True})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','right':True})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','right':True})
    elif data_type == 'text':
        data_right_format = wb.add_format({'right':True})
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # error if entered col_name not in dataframe

    # create list of all col_names
    col_name_list = [col_name for col_name in df.columns.get_level_values(1)]

    if col_name not in col_name_list:
        raise ValueError(f"{col_name} not in dataframe. Columns in data are: {col_name_list}")
    else:
        pass

    # getting row indices count of the data to use to set lower bound for formatting
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # error if it is a row multiindex
    if num_row_indices > 1:
        raise Exception(f"Function is only meant with datasets with a single row index or no row index. \
            The number of row indices your data has is {num_row_indices}.")

    # get number of unique columns in header row two
    cols_per_major_col = len(df.columns.get_level_values(1).unique())

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1

    # raise error if there are not the correct amount of header rows
    if num_col_indices != 2:
        raise Exception(f"Function is only meant for tables with two headers rows. {df} has {num_col_indices} header rows.")

    # format data

    # iterating through the columns:
    for col_num, df_col_name in enumerate(df.columns):
        # if the full (multilevel) column name contains the specificed column name:
        if col_name in df_col_name:
            # if the column is the last column per set in the second header column groupings
            if (col_num+1)%cols_per_major_col==0:
                # iterate through rows containing data
                for row_num, value in enumerate(df.values):
                    # if the value is null
                    if pd.isna(value[col_num]) and replace_nulls == True:
                        # insert null value and apply right border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                           null_value, null_right_format)
                    else:
                        # insert data and formatting with right border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                            value[col_num], data_right_format)
            else:
                # for all other columns iterate through rows containing data
                for row_num, value in enumerate(df.values):
                    # if the value is null
                    if pd.isna(value[col_num]) and replace_nulls == True:
                        # insert null value and apply null format
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                           null_value, null_format)
                    else:
                        # insert data and formatting with no border
                        sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                            value[col_num], data_format)
        else:
            # skip all columns that do not match specified column name
            pass  


###                 ROW MULTIINDEX AND TWO LEVEL COLUMN MULITINDEX DATAFRAMES                 ###


def insert_4d_data(df, wb, sheet, header_offset=0, column_offset=0, data_type=None, replace_nulls=True, null_value='-', \
    null_align='center'):

    # This function will insert your data in desired cells and apply a right border to the last column of each first level category
    # and a bottom border to the last row of each major (first) row category
    ## used on data with a two row column multiindex and any number of row multiindices

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### data_type is the type of numeric data:
    #### this arg should only be used if all your data is the same data type!
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    import pandas as pd
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        raise Exception('Data does not have a columns multiindex')

    # error if there are not 2 column header rows
    if num_col_indices == 2:
        pass 
    else:
        raise Exception(f"Function is only meant for datasets with two header rows. \
            The number of header rows your data has is {num_col_indices}.")

    #getting count of row_indices
    # if there is no index raise error
    if None in df.index.names:
        raise Exception("No index set on dataframe.")
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # exit function with error if it is not a multiindex
    if num_row_indices == 1:
        raise Exception("Function is not meant for single row index datasets.")
    else:
        pass

    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'center'})
        null_right_format = wb.add_format({'right':True, 'align':'center'})
        null_corner_format = wb.add_format({'right':True,'bottom':True, 'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'left'})
        null_right_format = wb.add_format({'right':True, 'align':'left'})
        null_corner_format = wb.add_format({'right':True,'bottom':True, 'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'right'})
        null_right_format = wb.add_format({'right':True, 'align':'right'})
        null_corner_format = wb.add_format({'right':True,'bottom':True, 'align':'right'})

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']    

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0','bottom':True})
        data_right_format = wb.add_format({'num_format':'#,##0','right':True})
        data_corner_format = wb.add_format({'num_format':'#,##0','right':True,'bottom':True})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.0','bottom':True})
        data_right_format = wb.add_format({'num_format':'#,##0.0','right':True})
        data_corner_format = wb.add_format({'num_format':'#,##0.0','right':True,'bottom':True})
    elif data_type == 'decimal_2':
        data_format = wb.add_format({'num_format':'#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.00','bottom':True})
        data_right_format = wb.add_format({'num_format':'#,##0.00','right':True})
        data_right_format = wb.add_format({'num_format':'#,##0.00','right':True,'bottom':True})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0','bottom':True})
        data_right_format = wb.add_format({'num_format':'$#,##0','right':True})
        data_corner_format = wb.add_format({'num_format':'$#,##0','right':True,'bottom':True})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0.00','bottom':True})
        data_right_format = wb.add_format({'num_format':'$#,##0.00','right':True})
        data_corner_format = wb.add_format({'num_format':'$#,##0.00','right':True,'bottom':True})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
        data_bottom_format = wb.add_format({'num_format':'0%','bottom':True})
        data_right_format = wb.add_format({'num_format':'0%','right':True})
        data_corner_format = wb.add_format({'num_format':'0%','right':True,'bottom':True})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
        data_bottom_format = wb.add_format({'num_format':'0.0%','bottom':True})
        data_right_format = wb.add_format({'num_format':'0.0%','right':True})
        data_corner_format = wb.add_format({'num_format':'0.0%','right':True,'bottom':True})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
        data_bottom_format = wb.add_format({'num_format':'0.00%','bottom':True})
        data_right_format = wb.add_format({'num_format':'0.00%','right':True})
        data_corner_format = wb.add_format({'num_format':'0.00%','right':True,'bottom':True})
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd','bottom':True})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd','right':True})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd','right':True,'bottom':True})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy','bottom':True})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy','right':True})
        data_corner_format = wb.add_format({'num_format':'m/d/yyyy','right':True,'bottom':True})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','bottom':True})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','right':True})
        data_corner_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','right':True,'bottom':True})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','bottom':True})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','right':True})
        data_corner_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','right':True,'bottom':True})
    elif data_type == 'text':
        print("Data types are text by default. No error, continuing function.")
        data_bottom_format = wb.add_format({'bottom':True})
        data_right_format = wb.add_format({'right':True})
        data_corner_format = wb.add_format({'right':True,'bottom':True})
    elif data_type == None:
        data_bottom_format = wb.add_format({'bottom':True})
        data_right_format = wb.add_format({'right':True})
        data_corner_format = wb.add_format({'right':True,'bottom':True})
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # determining how many rows are in the major (leftmost) index by dividing the total row count by index[0] unique values
    rows_per_major_index = int(len(df)/len(df.index.unique(0)))

    # number of header row 2 values which are what we need to loop over
    header2_n = df.columns.levshape[1]

    # getting the column count

    # getting the count of regular columns
    num_cols = len(df.columns)
    # adding num_cols and number of row indices together for total column count
    total_cols = num_row_indices + num_cols

    # value to subtract to correctly apply right border
    col_subtract = num_row_indices - 1

    # iterating over data columns excluding row index columns:
    for col_num in range(num_row_indices, total_cols):
        # iterating over rows containing data:
        for row_num, value in enumerate(df.values):
            # if no data type is assigned:
            if data_type == None or data_type == 'text':
                # for the last row per first index category:
                if (row_num + 1)%rows_per_major_index==0:
                    # for the last row per first index category:
                    if (col_num-col_subtract)%header2_n==0:
                        # we check if the value of the cell is null
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_corner_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_corner_format)
                    #else if the cell is on the last row but not the last column:
                    else:
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_bottom_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_bottom_format)
                #else if the row is NOT the last month row
                else:
                   # for the last row per first index category:
                    if (col_num-col_subtract)%header2_n==0:
                        # we check if the value of the cell is null
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_right_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_right_format)
                    # for all other cells
                    else:
                        # we check if the value of the cell is null
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_format)
                        else:
                            # otherwise insert data
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices])
            # else if data_type is not None or text:
            else:
                if (row_num + 1)%rows_per_major_index==0:
                    # for the last row per first index category:
                    if (col_num-col_subtract)%header2_n==0:
                        # we check if the value of the cell is null
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_corner_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_corner_format)
                    #else if the cell is on the last row but not the last column:
                    else:
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_bottom_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_bottom_format)
                #else if the row is NOT the last month row
                else:
                   # for the last row per first index category:
                    if (col_num-col_subtract)%header2_n==0:
                        # we check if the value of the cell is null
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_right_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_right_format)
                    # for all other cells
                    else:
                        # we check if the value of the cell is null
                        if pd.isna(value[col_num-num_row_indices]) and replace_nulls == True:
                            # replace value with null value and apply null formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, null_value, null_format)
                        else:
                            # otherwise insert data and apply formatting
                            sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset, \
                                value[col_num-num_row_indices], data_format)
                                

def set_4d_multiindex_dtype(df, wb, sheet, col_name, data_type, column_offset=0, header_offset=0, replace_nulls=True, null_value='-', \
    null_align='center'):

    # This function will insert your data in desired cells and apply a right border to the last column of each first level category
    # and a bottom border to the last row of each major (first) row category for the specified column
    ## used on data with a two row column multiindex and any number of row multiindices

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ### data_type is the type of numeric data:
    #### this arg should only be used if all your data is the same data type!
    #       'numeric' = comma-separated integer (ex 1,200)
    #       'decimal_1' = comma-separated decimal to tenths (ex 1,200.0)
    #       'decimal_2' = comma-separated decimal to hundredths (ex 1,200.00)
    #       'dollar' = comma-separated whole number currency (USD) (ex $1,200)
    #       'dollar_cents' = comma-separated decimal currency (USD) to hundredths (ex $1,200.00)
    #       'percent' = integer percentage (ex 20%)
    #       'percent_1' = decimal percentage to tenths (ex 20.0%)
    #       'percent_2' = decimal percentage to hundredths (ex 20.00%)
    #       'date' = sql-friendly date (ex 1992-08-14)
    #       'date_alt' = human-friendly date (ex 8/14/1992)
    #       'datetime' = sql-friendly datetime (ex 1992-08-14 17:26:00)
    #       'datetime_alt' = human-friendly datetime (ex 8/14/1992 5:22 PM)
    ### replace_nulls will replace null values with the specified replacement. defaults to true
    ### null_value is what replaces nulls. defaults to '-'
    ### null_align is the horizontal alignment for null values. defaults to center

    import pandas as pd
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        raise Exception('Data does not have a columns multiindex')

    # error if there are not 2 column header rows
    if num_col_indices == 2:
        pass 
    else:
        raise Exception(f"Function is only meant for datasets with two header rows. \
            The number of header rows your data has is {num_col_indices}.")

    #getting count of row_indices
    # if there is no index raise error
    if None in df.index.names:
        raise Exception("No index set on dataframe.")
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # exit function with error if it is not a multiindex
    if num_row_indices == 1:
        raise Exception("Function is not meant for single row index datasets.")
    else:
        pass

    # check for valid alignment input
    valid_align = ['center','left','right']

    if null_align in valid_align:
        pass
    else:
        raise ValueError(f"{null_align} is not a valid alignment option. Valid options are {valid_align}")
        
    # create list of all col_names
    col_name_list = [col_name for col_name in df.columns.get_level_values(1)]

    if col_name not in col_name_list:
        raise ValueError(f"{col_name} not in dataframe. Columns in data are: {col_name_list}")
    else:
        pass

    if null_align == 'center':
        null_format = wb.add_format({'align':'center'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'center'})
        null_right_format = wb.add_format({'right':True, 'align':'center'})
        null_corner_format = wb.add_format({'right':True,'bottom':True, 'align':'center'})
    elif null_align == 'left':
        null_format = wb.add_format({'align':'left'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'left'})
        null_right_format = wb.add_format({'right':True, 'align':'left'})
        null_corner_format = wb.add_format({'right':True,'bottom':True, 'align':'left'})
    elif null_align == 'right':
        null_format = wb.add_format({'align':'right'})
        null_bottom_format = wb.add_format({'bottom':True, 'align':'right'})
        null_right_format = wb.add_format({'right':True, 'align':'right'})
        null_corner_format = wb.add_format({'right':True,'bottom':True, 'align':'right'})

    # list of valid dtype args
    valid_dtypes = ['numeric','decimal_1','decimal_2','dollar','dollar_cents','percent','percent_1','percent_2','date',\
        'date_alt','datetime','datetime_alt']    

    # this if statement sets the formatting based off the data_type argument
    ## it will raise an error to tell the user if they have entered an invalid data_type argument
    if data_type == 'numeric':
        data_format = wb.add_format({'num_format':'#,##0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0','bottom':True})
        data_right_format = wb.add_format({'num_format':'#,##0','right':True})
        data_corner_format = wb.add_format({'num_format':'#,##0','right':True,'bottom':True})
    elif data_type == 'decimal_1':
        data_format = wb.add_format({'num_format':'#,##0.0'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.0','bottom':True})
        data_right_format = wb.add_format({'num_format':'#,##0.0','right':True})
        data_corner_format = wb.add_format({'num_format':'#,##0.0','right':True,'bottom':True})
    elif data_type == 'decimal_2':
        data_format = wb.add_format({'num_format':'#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'#,##0.00','bottom':True})
        data_right_format = wb.add_format({'num_format':'#,##0.00','right':True})
        data_right_format = wb.add_format({'num_format':'#,##0.00','right':True,'bottom':True})
    elif data_type == 'dollar':
        data_format = wb.add_format({'num_format':'$#,##0'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0','bottom':True})
        data_right_format = wb.add_format({'num_format':'$#,##0','right':True})
        data_corner_format = wb.add_format({'num_format':'$#,##0','right':True,'bottom':True})
    elif data_type == 'dollar_cents':
        data_format = wb.add_format({'num_format':'$#,##0.00'})
        data_bottom_format = wb.add_format({'num_format':'$#,##0.00','bottom':True})
        data_right_format = wb.add_format({'num_format':'$#,##0.00','right':True})
        data_corner_format = wb.add_format({'num_format':'$#,##0.00','right':True,'bottom':True})
    elif data_type == 'percent':
        data_format = wb.add_format({'num_format':'0%'})
        data_bottom_format = wb.add_format({'num_format':'0%','bottom':True})
        data_right_format = wb.add_format({'num_format':'0%','right':True})
        data_corner_format = wb.add_format({'num_format':'0%','right':True,'bottom':True})
    elif data_type == 'percent_1':
        data_format = wb.add_format({'num_format':'0.0%'})
        data_bottom_format = wb.add_format({'num_format':'0.0%','bottom':True})
        data_right_format = wb.add_format({'num_format':'0.0%','right':True})
        data_corner_format = wb.add_format({'num_format':'0.0%','right':True,'bottom':True})
    elif data_type == 'percent_2':
        data_format = wb.add_format({'num_format':'0.00%'})
        data_bottom_format = wb.add_format({'num_format':'0.00%','bottom':True})
        data_right_format = wb.add_format({'num_format':'0.00%','right':True})
        data_corner_format = wb.add_format({'num_format':'0.00%','right':True,'bottom':True})
    elif data_type == 'date':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd','bottom':True})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd','right':True})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd','right':True,'bottom':True})
    elif data_type == 'date_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy','bottom':True})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy','right':True})
        data_corner_format = wb.add_format({'num_format':'m/d/yyyy','right':True,'bottom':True})
    elif data_type == 'datetime':
        data_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm'})
        data_bottom_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','bottom':True})
        data_right_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','right':True})
        data_corner_format = wb.add_format({'num_format':'yyyy-mm-dd h:mm','right':True,'bottom':True})
    elif data_type == 'datetime_alt':
        data_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM'})
        data_bottom_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','bottom':True})
        data_right_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','right':True})
        data_corner_format = wb.add_format({'num_format':'m/d/yyyy h:mm AM/PM','right':True,'bottom':True})
    elif data_type == 'text':
        print("Data types are text by default. No error, continuing function.")
        data_bottom_format = wb.add_format({'bottom':True})
        data_right_format = wb.add_format({'right':True})
        data_corner_format = wb.add_format({'right':True,'bottom':True})
    elif data_type == None:
        data_bottom_format = wb.add_format({'bottom':True})
        data_right_format = wb.add_format({'right':True})
        data_corner_format = wb.add_format({'right':True,'bottom':True})
    else:
        raise ValueError(f"{data_type} is not a valid data_format option. Valid options are: {valid_dtypes}")

    # determining how many rows are in the major (leftmost) index by dividing the total row count by index[0] unique values
    rows_per_major_index = int(len(df)/len(df.index.unique(0)))

    # number of header row 2 values which are what we need to loop over
    header2_n = df.columns.levshape[1]

    # getting the column count

    # getting the count of regular columns
    num_cols = len(df.columns)
    # adding num_cols and number of row indices together for total column count
    total_cols = num_row_indices + num_cols

    # value to subtract to correctly apply right border
    col_subtract = num_row_indices - 1

    # iterating over data columns excluding row index columns:
    for col_num, df_col_name in enumerate(df.columns):
        # if the full (multilevel) column name contains the specificed column name:
        if col_name in df_col_name:
            # iterating over rows containing data:
            for row_num, value in enumerate(df.values):
                # if no data type is assigned:
                if data_type == None or data_type == 'text':
                    # for the last row per first index category:
                    if (row_num + 1)%rows_per_major_index==0:
                        # for the last row per first index category:
                        if (col_num+1)%header2_n==0:
                            # we check if the value of the cell is null
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    null_value, null_corner_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_corner_format)
                        #else if the cell is on the last row but not the last column:
                        else:
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices,\
                                     null_value, null_bottom_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_bottom_format)
                    #else if the row is NOT the last month row
                    else:
                    # for the last row per first index category:
                        if (col_num+1)%header2_n==0:
                            # we check if the value of the cell is null
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    null_value, null_right_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_right_format)
                        # for all other cells
                        else:
                            # we check if the value of the cell is null
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    null_value, null_format)
                            else:
                                # otherwise insert data
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num])
                # else if data_type is not None or text:
                else:
                    if (row_num + 1)%rows_per_major_index==0:
                        # for the last row per first index category:
                        if (col_num+1)%header2_n==0:
                            # we check if the value of the cell is null
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices \
                                    , null_value, null_corner_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_corner_format)
                        #else if the cell is on the last row but not the last column:
                        else:
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    null_value, null_bottom_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_bottom_format)
                    #else if the row is NOT the last month row
                    else:
                    # for the last row per first index category:
                        if (col_num+1)%header2_n==0:
                            # we check if the value of the cell is null
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    null_value, null_right_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_right_format)
                        # for all other cells
                        else:
                            # we check if the value of the cell is null
                            if pd.isna(value[col_num]) and replace_nulls == True:
                                # replace value with null value and apply null formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    null_value, null_format)
                            else:
                                # otherwise insert data and apply formatting
                                sheet.write(row_num + num_col_indices + header_offset, col_num + column_offset + num_row_indices, \
                                    value[col_num], data_format) 
        else:
            # skip all columns that do not match specified column name
            pass    



######################## EDGE BORDER FORMATTING ##################################

###                      ANY SHAPE DATAFRAMES                        ###

def table_bottom_border(df, wb, sheet, header_offset=0, column_offset=0):

    # This function will apply formatting a bottom border to your table
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    
    # getting row count of the data to use to set lower bound for formatting
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1
    # get the row count (which doesn't count column rows)
    data_rows = len(df)
    # add the two together to get total row count
    df_row_count = num_col_indices + data_rows + header_offset

    # getting count of number of row indices to set range for index formatting
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        num_row_indices = len(df.index.names)

    # creating the format for the bottom border (actually top border on the cell below so we don't overwrite data)
    bottom_format = wb.add_format({'top':True})

    # this applies a top border to the cell below the last row fo data for all the columns except the index
    for col_num, value in enumerate(df.columns.values):
        # we are applying a top border to that to fake a bottom border on the table!
        # "" is filling in the cell with nothing, leaving it blank
        # col_num + row_indices will correctly skip the row index columns in the loop
        sheet.write(df_row_count, col_num + num_row_indices + column_offset, "", bottom_format)
    
    # this applies a top border to the cells below the last row of the index columns since they are excluded from the column for loop
    for col_num in range(num_row_indices):
        sheet.write(df_row_count, col_num + column_offset, "", bottom_format)


def table_right_border(df, wb, sheet, header_offset=0, column_offset=0):

    # This function will apply formatting a right border to your table
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0

    # getting the column count

    # getting the count of row index columns
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df.index.names:
        num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
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
    # getting count of the data rows
    data_rows = len(df)
    # adding them together to get total rows
    total_rows = num_col_indices + data_rows + header_offset

    # creating right border format--actually left to next cell over to avoid overwriting data
    right_format = wb.add_format({'left':True})

    # iterating over all our rows in our table:
    for row_num in range(header_offset, total_rows):
        # apply the right format to the first column after our table
        sheet.write(row_num, total_cols + column_offset, "", right_format)


def table_left_border(df, wb, sheet, column_offset, header_offset=0):

    # This function will apply formatting a left border to your table
    ## Can be used on any dataframe

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### header_offset is the number of rows to skip if you want blank rows on top for title etc. defaults to 0
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0

   # getting row count of the data to use to set lower bound for formatting

   # raise exception if attempting to apply to table that starts in column A
    if column_offset == 0:
        raise Exception("Left border cannot be applied to tables that start in column A.")
    else:
        pass
    
    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1
    # getting count of the data rows
    data_rows = len(df)
    # adding them together to get total rows
    total_rows = num_col_indices + data_rows + header_offset

    # creating left border format--actually right to next cell over to avoid overwriting data
    left_format = wb.add_format({'right':True})  

    # iterating over all our rows in our table:
    for row_num in range(header_offset, total_rows):
        # apply the left format to the first column before our table
        sheet.write(row_num, column_offset - 1, "", left_format) 


######################## TITLE FORMATTING ##################################

###                      ANY SHAPE DATAFRAMES                        ###

def insert_title(df, wb, sheet, title, font_size=16, font_color='#000000', bg_color='#ffffff', align='left', row_num=0, col_num=0):

    # This function will insert a title for your table

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your data from your dataframe
    ### wb is your workbook
    ### sheet is your worksheet
    ### title is your title, entered as a string

    ## OPTIONAL:
    ### font_size is the font size for title. defaults to 16
    ### font_color is the font color for title. defaults to black
    ### bg_color is the background color the cell containing the title. defaults to white
    ### align is the horizontal text alignmnet. defaults to left
    ### cell is the cell where the title will be placed. default to A1    
    ### row_num is the row to place your title, defaults to excel row 1
    ### col_num is the column to place your title, defaults to excel column A

    # raising an error message to tell the user if they have entered an invalid alignmnet argument
    valid_alignments = ['left','center','right','fill','justify','center_across','distributed']

    if align not in valid_alignments:
        raise ValueError(f"{align} is not a valid alignment option. Valid options are: {valid_alignments}")

    # creating title format
    title_format = wb.add_format({'bold':True, 'font_color':font_color, 'bg_color':bg_color, 'font_size':font_size,'align':align})

    # applying title format and inserting title
    sheet.write(row_num, col_num, title, title_format)


######################## TWO TABLES PER SHEET FORMATTING ##################################

###                      ANY SHAPE DATAFRAMES                        ###

def create_skip_rows(df, header_offset=0, rows_between=2):

    # this function will generate the amount of rows to skip if you are putting two tables on the same sheet
    ## it is meant to be used when there are two tables, one under the top one

    # ARGUMENTS
    
    ## MANDATORY:
    ### df is your top dataframe

    ## OPTIONAL:
    ### header_offset is your blank rows for titles for your top table. it should match header_offset from functions you used to format it
    ### rows_between is the number of rows between tables. defaults to 

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
        num_col_indices = len(df.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        num_col_indices = 1  

    # funtion returns the number of rows in the dataframe + number of header rows + header_offset + rows _between for num rows to skip
    return len(df) + num_col_indices + header_offset + rows_between


def two_table_col_widths(df1, df2, wb, sheet,column_offset=0,method='headers',text_wrap=False,wrap_rows=2,\
    df1_col_multiindex=False,df2_col_multiindex=False):

    # this function will automatically set column widths for sheets with two tables of data on them
    ## the two data tables must be stacked on top of each other
    ## it will apply the width of whichever dataframe has the widest column to each column

    # ARGUMENTS
    
    ## MANDATORY:
    ### df1 is your data from your first dataframe
    ### df2 is your data from your second dataframe
    ### wb is your workbook
    ### sheet is your worksheet

    ## OPTIONAL:
    ### column_offset is the number of columns to shift to the right if you do not want your table to start on column A. defaults to 0
    ####    column_offset for both dataframes MUST be the same in all your other functions, if not 0
    ### method is how the width is set:
    #       'header' sets width based on the length of column names. is default
    #       'data' sets width based on the length of the longest data point in the column
    #       'all' sets width based off the column name or longest data point, whichever is larger
    ### text_wrap will wrap text in headers when True. default is False
    ### wrap_rows is how many rows wide the wrapped text should be. default is 2
    ### df1_col_multiindex should be changed to True if you have multiple header rows for df1. defaults to False
    ### df2_col_multiindex should be changed to True if you have multiple header rows for df2. defaults to False

    import numpy as np
    from math import ceil

    # check to make sure column_offset input is valid
    ## if column_offset is not an integer, raise an error
    if isinstance(column_offset, int) == False:
        raise TypeError(f"{column_offset} is not a valid argument for column_offset. column_offset must be an integer.")
    else:
        pass

    # check to make sure wrap_rows input is valid
    ## if wrap_rows is not an integer, raise an error
    if isinstance(wrap_rows, int) == False:
        raise TypeError(f"{wrap_rows} is not a valid argument for wrap_rows. wrap_rows must be an integer.")
    else:
        pass

    # list of valid method args
    valid_methods = ['headers', 'data', 'all']

    # error if valid method arg not used
    if method not in valid_methods:
        raise ValueError(f"{method} is not a valid method option. Valid methods are: {valid_methods}")
    else:
        pass

    # get data col counts
    df1_col_count = len(df1.columns)
    df2_col_count = len(df2.columns)

    # get the count of how many row indices
    # if there is no index set to 0 (pandas has a default index with no name)
    if None in df1.index.names:
        df1_num_row_indices = 0
    else:
        # else number of row indices is how many row index names there are    
        df1_num_row_indices = len(df1.index.names)

    if None in df2.index.names:
        df2_num_row_indices = 0
    else:
        df2_num_row_indices = len(df2.index.names)

    # get total columns per df
    df1_total_cols = df1_col_count + df1_num_row_indices
    df2_total_cols = df2_col_count + df2_num_row_indices

    max_cols = max(df1_total_cols, df2_total_cols)

    # this will try to get the count of column levels you have if it's a multiindex but if it fails since it's only one level
    try:
       df1_num_col_indices = len(df1.columns.levshape)
    # then it will assign a value of 1 for column_indices
    except:
        df1_num_col_indices = 1  

    try:
       df2_num_col_indices = len(df2.columns.levshape)
    except:
        df2_num_col_indices = 2  

    # if df1 does not have multiple header rows:
    if df1_col_multiindex == False:
        if df1_num_col_indices > 1:
            # raise error if there ARE multiple header rows
            raise Exception(f"{df1} has a column multiindex. {df1} has {df1_num_col_indices} header rows.")
        else:
            # create a list holding the length of the name of each column
            ## + 1 for 'wiggle room'
            df1_col_names = [name for name in df1.columns]

            # adjust col widths for text wrapping of headers
            if text_wrap == True:
                # if there is text wrapping, the width of columns is their previous col_name length divided by the number of rows for wrapping 
                ## rounded up so that it is an integer value to prevent errors
                df1_col_name_lengths = []
                for name in df1_col_names:
                    if " " in name:
                        df1_length = ceil(len(name)/wrap_rows) + 1
                    else:
                        df1_length = len(name) + 1
                    df1_col_name_lengths.append(df1_length)
            elif text_wrap==False:
                df1_col_name_lengths = [len(name) + 1 for name in df1_col_names]
            else:
                raise ValueError(f"{text_wrap} is not not a valid text_wrap argument. text_wrap must be True or False.")

            # create a list holding the max length of the data in each column
            df1_max_data_lengths = []
                
            # iterating over the data columns:    
            for col in list(df1):
                # store their values in a list
                df1_values = df1[col].tolist()
                # create an empty list to store the lengths
                df1_value_lengths = []
                # iterating over the values list:
                for row_num, value in enumerate(df1_values):
                    # get the length in characters of each value
                    df1_length = len(str(value))
                    # add it to the value_lengths list
                    df1_value_lengths.append(df1_length)
                    # if it is the final iteration over the values with the completed value_lengths list for the column:
                    ## + 1 since python numbering starts at 0
                    if row_num + 1 == len(df1_values):
                        # get the max width value
                        ## + 1 for 'wiggle room'
                        df1_max_data_width = max(df1_value_lengths) + 1
                        # append it to our data lengths list
                        df1_max_data_lengths.append(df1_max_data_width)

            # create a list for the max of data and column width, whichever is greater
            df1_max_all_lengths = np.maximum(df1_col_name_lengths, df1_max_data_lengths)

            # choosing list to use based on method:
            if method == 'headers':
                df1_col_widths = df1_col_name_lengths
            elif method == 'data':
                df1_col_widths = df1_max_data_lengths
            elif method ==  'all':
                df1_col_widths = df1_max_all_lengths

            # get widths of index columns
            # if there is no index, set width list to the column width list
            if df1_num_row_indices == 0:
                df1_width_list = df1_col_widths 
            # if there is a single row index:
            elif df1_num_row_indices == 1:
                # gets the length of all the values in the index
                df1_index_values = [len(value) for row_num, value in enumerate(df1.index.values)]

                # gets the max of the index values or the name of the index, whichever is greater
                ## + 1 for 'wiggle room'
                ## in list form so .extend() will work below
                df1_max_index_lengths = [max(max(df1_index_values), len(df1.index.name)) + 1]

                # creates a copy of the index length list so we do not modify the original
                df1_width_list = df1_max_index_lengths.copy()

                # adds the col widths to the index list
                df1_width_list.extend(df1_col_widths)
            # if these is a row multiindex:
            else:
                # create empty list to hold max_index_lengths
                df1_max_index_lengths = []

                # iterating over row indices:
                for col_num in range(df1_num_row_indices):
                    # store the length of all index values in a list
                    df1_index_values = [len(value) for i, value in enumerate(df1.index.get_level_values(col_num))]
                    # get the max width of the longest value or title, whichever is longer
                    ## + 1 for 'wiggle room'
                    df1_max_index_length = max(max(df1_index_values), len(df1.index.names[col_num])) + 1
                    # add that to the max_index_lengths list
                    df1_max_index_lengths.append(df1_max_index_length)

                # creates a copy of the index length list so we do not modify the original
                df1_width_list = df1_max_index_lengths.copy()

                # adds the col widths to the index list
                df1_width_list.extend(df1_col_widths)          
            
            # if the total number of columns for the dataframe is less than the total columns to be formatted:
            if df1_total_cols < max_cols:
                # iterating through our number of columns to be formatted:
                for col_num in range(max_cols):
                    try:
                        # keep list values that exist
                        df1_width_list[col_num] == df1_width_list[col_num]
                    except IndexError:
                        # add 0 to end of list for number of extra columns the other dataframe has
                        df1_width_list.append(0)
    
    # if the dataframe has two header rows:
    elif df1_col_multiindex == True:
        if df1_num_col_indices != 2:
            # raise error if there are NOT exactly two header rows
            raise Exception(f"{df1} has {df1_num_col_indices} header rows. df1_col_multiindex only works for 2 header rows.")
        else:
            # create a list holding the length of the name of each column
            ## + 1 for 'wiggle room'
            df1_col_names = [name for name in df1.columns.get_level_values(1)]

            # adjust col widths for text wrapping of headers
            if text_wrap == True:
                df1_col_name_lengths = []
                # if there is text wrapping, the width of columns is their previous col_name length 
                # divided by the number of rows for wrapping 
                ## rounded up so that it is an integer value to prevent errors
                for name in df1_col_names:
                    if " " in name:
                        df1_length = ceil(len(name)/wrap_rows) + 1
                    else:
                        df1_length = len(name) + 1
                    df1_col_name_lengths.append(df1_length)
            elif text_wrap==False:
                df1_col_name_lengths = [len(name) + 1 for name in df1_col_names]
            else:
                raise ValueError(f"{text_wrap} is not not a valid text_wrap argument. text_wrap must be True or False.")

            # create a list holding the max length of the data in each column
            df1_max_data_lengths = []
                
            # iterating over the data columns:    
            for col in list(df1):
                # store their values in a list
                df1_values = df1[col].tolist()
                # create an empty list to store the lengths
                df1_value_lengths = []
                # iterating over the values list:
                for row_num, value in enumerate(df1_values):
                    # get the length in characters of each value
                    df1_length = len(str(value))
                    # add it to the value_lengths list
                    df1_value_lengths.append(df1_length)
                    # if it is the final iteration over the values with the completed value_lengths list for the column:
                    ## + 1 since python numbering starts at 0
                    if row_num + 1 == len(df1_values):
                        # get the max width value
                        ## + 1 for 'wiggle room'
                        df1_max_data_width = max(df1_value_lengths) + 1
                        # append it to our data lengths list
                        df1_max_data_lengths.append(df1_max_data_width) 
            
            # create a list for the max of data and column width, whichever is greater
            df1_max_all_lengths = np.maximum(df1_col_name_lengths, df1_max_data_lengths)

            # choosing list to use based on method:
            if method == 'headers':
                df1_col_widths = df1_col_name_lengths
            elif method == 'data':
                df1_col_widths = df1_max_data_lengths
            elif method ==  'all':
                df1_col_widths = df1_max_all_lengths

            # get row index col widths
            
            # if there is no row index:  
            if df1_num_row_indices == 0:
                # set the width list to the column widths
                df1_width_list = df1_col_widths 
            # if these is a single row index:
            elif df1_num_row_indices == 1:
                # gets the length of all the values in the index
                df1_index_values = [len(value) for row_num, value in enumerate(df1.index.values)]

                # gets the max of the index values or the name of the index, whichever is greater
                ## + 1 for 'wiggle room'
                df1_max_index_lengths = [max(max(df1_index_values), len(df1.index.name)) + 1]

                # create a copy so we don't modify the original
                df1_width_list = df1_max_index_lengths.copy()

                # add the data column widths to the list
                df1_width_list.extend(df1_col_widths)
            # else for row multiindex:
            else:
                # create empty list to hold max_index_lengths
                df1_max_index_lengths = []

                # iterating over row indices:
                for col_num in range(df1_num_row_indices):
                    # store the length of all index values in a list
                    df1_index_values = [len(value) for i, value in enumerate(df1.index.get_level_values(col_num))]
                    # get the max width of the longest value or title, whichever is longer
                    ## + 1 for 'wiggle room'
                    df1_max_index_length = max(max(df1_index_values), len(df1.index.names[col_num])) + 1
                    # add that to the max_index_lengths list
                    df1_max_index_lengths.append(df1_max_index_length)

                # create a copy of index lengths so we don't modify original
                df1_width_list = df1_max_index_lengths.copy()

                # add col width lengths to list
                df1_width_list.extend(df1_col_widths)          
            
            # if the total number of columns for the dataframe is less than the total columns to be formatted:
            if df1_total_cols < max_cols:
                # iterating through our columns to be formatted:
                for col_num in range(max_cols):
                    try:
                        # keep list values
                        df1_width_list[col_num] == df1_width_list[col_num]
                    except IndexError:
                        # add 0 to end of list for number of extra columns the other dataframe has
                        df1_width_list.append(0) 
        

    # if df2 does not have multiple header rows:
    if df2_col_multiindex == False:
        if df2_num_col_indices > 1:
            # raise error if there ARE multiple header rows
            raise Exception(f"{df2} has a column multiindex. {df2} has {df2_num_col_indices} header rows.")
        else:
            # create a list holding the length of the name of each column
            ## + 1 for 'wiggle room'
            df2_col_names = [name for name in df2.columns]

            # adjust col widths for text wrapping of headers
            if text_wrap == True:
                df2_col_name_lengths = []
                for name in df2_col_names:
                    if " " in name:
                        df2_length = ceil(len(name)/wrap_rows) + 1
                    else:
                        df2_length = len(name) + 1
                    df2_col_name_lengths.append(df2_length)
            elif text_wrap==False:
                df2_col_name_lengths = [len(name) + 1 for name in df2_col_names]
            else:
                raise ValueError(f"{text_wrap} is not not a valid text_wrap argument. text_wrap must be True or False.")

            # create a list holding the max length of the data in each column
            df2_max_data_lengths = []
                
            # iterating over the data columns:    
            for col in list(df2):
                # store their values in a list
                df2_values = df2[col].tolist()
                # create an empty list to store the lengths
                df2_value_lengths = []
                # iterating over the values list:
                for row_num, value in enumerate(df2_values):
                    # get the length in characters of each value
                    df2_length = len(str(value))
                    # add it to the value_lengths list
                    df2_value_lengths.append(df2_length)
                    # if it is the final iteration over the values with the completed value_lengths list for the column:
                    ## + 1 since python numbering starts at 0
                    if row_num + 1 == len(df2_values):
                        # get the max width value
                        ## + 1 for 'wiggle room'
                        df2_max_data_width = max(df2_value_lengths) + 1
                        # append it to our data lengths list
                        df2_max_data_lengths.append(df2_max_data_width)

            # create a list for the max of data and column width, whichever is greater
            df2_max_all_lengths = np.maximum(df2_col_name_lengths, df2_max_data_lengths)

            # choosing list to use based on method:
            if method == 'headers':
                df2_col_widths = df2_col_name_lengths
            elif method == 'data':
                df2_col_widths = df2_max_data_lengths
            elif method ==  'all':
                df2_col_widths = df2_max_all_lengths

            # get widths of index columns
            # if there is no index, set width list to the column width list
            if df2_num_row_indices == 0:
                df2_width_list = df2_col_widths 
            # if there is a single row index:
            elif df2_num_row_indices == 1:
                # gets the length of all the values in the index
                df2_index_values = [len(value) for row_num, value in enumerate(df2.index.values)]

                # gets the max of the index values or the name of the index, whichever is greater
                ## + 1 for 'wiggle room'
                ## in list form so .extend() will work below
                df2_max_index_lengths = [max(max(df2_index_values), len(df2.index.name)) + 1]

                # creates a copy of the index length list so we do not modify the original
                df2_width_list = df2_max_index_lengths.copy()

                # adds the col widths to the index list
                df2_width_list.extend(df2_col_widths)
            # if these is a row multiindex:
            else:
                # create empty list to hold max_index_lengths
                df2_max_index_lengths = []

                # iterating over row indices:
                for col_num in range(df2_num_row_indices):
                    # store the length of all index values in a list
                    df2_index_values = [len(value) for i, value in enumerate(df2.index.get_level_values(col_num))]
                    # get the max width of the longest value or title, whichever is longer
                    ## + 1 for 'wiggle room'
                    df2_max_index_length = max(max(df2_index_values), len(df2.index.names[col_num])) + 1
                    # add that to the max_index_lengths list
                    df2_max_index_lengths.append(df2_max_index_length)

                # creates a copy of the index length list so we do not modify the original
                df2_width_list = df2_max_index_lengths.copy()

                # adds the col widths to the index list
                df2_width_list.extend(df2_col_widths)          
            
            # if the total number of columns for the dataframe is less than the total columns to be formatted:
            if df2_total_cols < max_cols:
                # iterating through our number of columns to be formatted:
                for col_num in range(max_cols):
                    try:
                        # keep list values that exist
                        df2_width_list[col_num] == df2_width_list[col_num]
                    except IndexError:
                        # add 0 to end of list for number of extra columns the other dataframe has
                        df2_width_list.append(0)
    
    # if the dataframe has two header rows:
    elif df2_col_multiindex == True:
        if df2_num_col_indices != 2:
            # raise error if there are NOT exactly two header rows
            raise Exception(f"{df2} has {df2_num_col_indices} header rows. df2_col_multiindex only works for 2 header rows.")
        else:
            # create a list holding the length of the name of each column
            ## + 1 for 'wiggle room'
            df2_col_names = [name for name in df2.columns.get_level_values(1)]

            # adjust col widths for text wrapping of headers
            if text_wrap == True:
                df2_col_name_lengths = []
                # if there is text wrapping, the width of columns is their previous col_name length 
                # divided by the number of rows for wrapping 
                ## rounded up so that it is an integer value to prevent errors
                for name in df2_col_names:
                    if " " in name:
                        df2_length = ceil(len(name)/wrap_rows) + 1
                    else:
                        df2_length = len(name) + 1
                    df2_col_name_lengths.append(df2_length)
            elif text_wrap==False:
                df2_col_name_lengths = [len(name) + 1 for name in df2_col_names]
            else:
                raise ValueError(f"{text_wrap} is not not a valid text_wrap argument. text_wrap must be True or False.")

            # create a list holding the max length of the data in each column
            df2_max_data_lengths = []
                
            # iterating over the data columns:    
            for col in list(df2):
                # store their values in a list
                df2_values = df2[col].tolist()
                # create an empty list to store the lengths
                df2_value_lengths = []
                # iterating over the values list:
                for row_num, value in enumerate(df2_values):
                    # get the length in characters of each value
                    df2_length = len(str(value))
                    # add it to the value_lengths list
                    df2_value_lengths.append(df2_length)
                    # if it is the final iteration over the values with the completed value_lengths list for the column:
                    ## + 1 since python numbering starts at 0
                    if row_num + 1 == len(df2_values):
                        # get the max width value
                        ## + 1 for 'wiggle room'
                        df2_max_data_width = max(df2_value_lengths) + 1
                        # append it to our data lengths list
                        df2_max_data_lengths.append(df2_max_data_width) 
            
            # create a list for the max of data and column width, whichever is greater
            df2_max_all_lengths = np.maximum(df2_col_name_lengths, df2_max_data_lengths)

            # choosing list to use based on method:
            if method == 'headers':
                df2_col_widths = df2_col_name_lengths
            elif method == 'data':
                df2_col_widths = df2_max_data_lengths
            elif method ==  'all':
                df2_col_widths = df2_max_all_lengths

            # get row index col widths
            
            # if there is no row index:  
            if df2_num_row_indices == 0:
                # set the width list to the column widths
                df2_width_list = df2_col_widths 
            # if these is a single row index:
            elif df2_num_row_indices == 1:
                # gets the length of all the values in the index
                df2_index_values = [len(value) for row_num, value in enumerate(df2.index.values)]

                # gets the max of the index values or the name of the index, whichever is greater
                ## + 1 for 'wiggle room'
                df2_max_index_lengths = [max(max(df2_index_values), len(df2.index.name)) + 1]

                # create a copy so we don't modify the original
                df2_width_list = df2_max_index_lengths.copy()

                # add the data column widths to the list
                df2_width_list.extend(df2_col_widths)
            # else for row multiindex:
            else:
                # create empty list to hold max_index_lengths
                df2_max_index_lengths = []

                # iterating over row indices:
                for col_num in range(df2_num_row_indices):
                    # store the length of all index values in a list
                    df2_index_values = [len(value) for i, value in enumerate(df2.index.get_level_values(col_num))]
                    # get the max width of the longest value or title, whichever is longer
                    ## + 1 for 'wiggle room'
                    df2_max_index_length = max(max(df2_index_values), len(df2.index.names[col_num])) + 1
                    # add that to the max_index_lengths list
                    df2_max_index_lengths.append(df2_max_index_length)

                # create a copy of index lengths so we don't modify original
                df2_width_list = df2_max_index_lengths.copy()

                # add col width lengths to list
                df2_width_list.extend(df2_col_widths)          
            
            # if the total number of columns for the dataframe is less than the total columns to be formatted:
            if df2_total_cols < max_cols:
                # iterating through our columns to be formatted:
                for col_num in range(max_cols):
                    try:
                        # keep list values
                        df2_width_list[col_num] == df2_width_list[col_num]
                    except IndexError:
                        # add 0 to end of list for number of extra columns the other dataframe has
                        df2_width_list.append(0)

    # the final width list takes the maximum width value for each column from the two dataframes
    width_list = np.maximum(df1_width_list, df2_width_list)

    # iterating over the total columns:
    for col_num, width in enumerate(width_list):
        # apply the matching width to the column
        sheet.set_column(col_num + column_offset, col_num + column_offset, width)

    