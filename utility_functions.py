def return_divisible_ints(start_num, end_num, denominator):
    
    # this function will return the numbers divisible by the denominator between start_num and end_num, inclusive

    # ARGUMENTS

    # MANDATORY:
    ## start_num is the starting number
    ## end_num is the ending number
    ## denominator is the number to divide by
    
    # create empty list to hold results
    divisible_ints = []
    
    # for each number in the range of start_num to end_num + 1 (to keep end_num inclusive in results)
    for i in range(start_num, end_num+1):
        # if the remainder of i divided by the denominator is 0
        if i%denominator==0:
            # add to list of divisible ints
            divisible_ints.append(i)
    
    return divisible_ints


def clean_header_string(string):

    # this function will convert an underscore separted or camelcase string to title format
    ## ex: "birth_date" or "BirthDate" to "Birth Date"

    # MANDATORY:
    ## string is the string to convert

    import re

    # if string contains underscores
    #if re.fullmatch(r'^([a-z][A-Z][0-9]+)(_[a-z][A-Z][0-9]+)+$', string) != None:
    if '_' in string:
        # replace underscore with space and capitalize all first letters
        return string.title().replace('_',' ')
    # # if string is camelcase
    elif re.fullmatch(r'(?:[A-Z][a-z]+)+', string) != None:
        # insert space before the next captalized letter
        return re.sub(r'((?<=[a-z])[A-Z]|(?<!\A)[A-Z](?=[a-z]))', r' \1', string)
    else: 
        # else return string with first letter capitalized
        return string.title()