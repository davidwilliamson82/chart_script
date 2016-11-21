def flatten(list, function = lambda x: x):
    # This just flattens a list, with an optional function applied to the elements.
    return [function(item) for sublist in list for item in sublist]

def to_alpha(number):
    # This takes a base-10 number and returns a base-26 alphabetical index.
    from string import ascii_uppercase as alphabet
    number += 1
    exp, output = 0, ''
    while divmod(number - 1, 26 ** (exp + 1))[0] > 0: exp += 1
    for ex in range(exp, -1, -1):
        quotient, number = divmod(number, 26 ** ex)
        output += alphabet[quotient - 1]
    return output

def from_alpha(string):
    # This takes a base-26 alphabetical index and returns a base-10 number.
    from string import ascii_uppercase as alphabet
    return sum([(alphabet.index(c) + 1) * 26 ** (len(string) - (i + 1)) for i, c in enumerate(string)]) - 1

def chart_strings(number_of_tables, interval, ranges, sheet = 'Sheet1'):
    """This function takes some information about multiple excel tables 
    and generates references for multiple columns in those tables, 
    for the purpose of generating line graphs from the data."""
    from string import ascii_uppercase
    from re import match, split as re_split
    def offset(index, table_index):
        # This adjusts the index of a column, based on the table_index.
        # It depends on re.match.
        if index[0] == '$':
            return index.replace('$', '')
        if match(r'[A-Z]+', index):
            return to_alpha(from_alpha(index) + table_index)
        else:
            return index 
    
    output = []
    prefix = "='{}'!".format(sheet)

    # To generate the list of range references, each name is followed by the prefix and the offset range indices.
    for table_index in range(0, number_of_tables * interval, interval):
        table_references = []
        for name, cells in ranges:
            reference = [offset(index, table_index) 
                         for index in flatten([re_split(r'(\d+)', index)[:-1] 
                                               for index in cells.split(':')])]
            if len(reference) > 2:
                table_references.append('{}{}${}${}:${}${}'.format(name, prefix, *reference))
            else:
                table_references.append('{}{}${}${}'.format(name, prefix, *reference))
        output.append(tuple(table_references)) 
    return tuple(output)

def make_range(top_row, bottom_row):
    """Given top and bottom row values, this returns a function which accepts the column as an argument. 
    The resulting function will return the reference for the range specified by the column index."""
    return lambda column: '{0}{1}:{0}{2}'.format(column, top_row, bottom_row)

# test it out
from pprint import pprint
column = make_range(4, 369)
pprint(chart_strings(8, 4, (('header', 'B1:D1'), 
                            ('horizontal', column('$A')), 
                            ('benefits', column('B')), 
                            ('costs', column('C')), 
                            ('ratios', column('D'))), 
                     sheet = 'Benefit-Cost Ratios'))
