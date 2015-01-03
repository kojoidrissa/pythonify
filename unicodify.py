def unicodify(seq):
    '''
    sequence --> list

    Input an iterable squence, 
    return a list of unicode strings of the objects in the source sequence
    '''
    strings = []
    for i in seq:
        strings.append(unicode(i))
    return strings