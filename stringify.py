def stringify(seq):
    '''
    sequence --> list
    
    Input an iterable squence, 
    return a list of strings of the objects in the source sequence
    '''
    strings = []
    for i in seq:
        strings.append(str(i))
    return strings