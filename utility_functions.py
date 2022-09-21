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