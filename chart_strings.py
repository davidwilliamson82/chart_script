def flatten(list, function = lambda x: x):
    # This just flattens a list, with an optional function applied to the elements.
    return [function(item) for sublist in list for item in sublist]

def chart_strings(number_of_tables, interval, ranges, sheet = 'Sheet1'):
    """This function takes some information about multiple excel tables 
    and generates references for multiple columns in those tables, 
    for the purpose of generating line graphs from the data."""
    from string import ascii_uppercase
    from re import match, split as re_split
    def offset(index, table_index, indices_list):
        # This adjusts the index of a column, based on the table_index.
        # It depends on re.match.
        if index[0] == '$':
            return index.replace('$', '')
        if match(r'[A-Z]+', index):
            return indices_list[indices_list.index(index) + table_index]
        else:
            return index 
    
    output = []
    prefix = "='{}'!".format(sheet)

    #  These lines generate a list of the first 702 column index strings. 
    alpha = ascii_uppercase
    beta = flatten([[first + second for second in alpha] for first in alpha])
    alpha = list(alpha) + beta
    
    # To generate the list of range references, each name is followed by the prefix and the offset range indices.
    for table_index in range(0, number_of_tables * interval, interval):
        table_references = []
        for name, cells in ranges:
            reference = [offset(index, table_index, alpha) 
                         for index in flatten([re_split(r'(\d+)', index)[:-1] 
                                               for index in cells.split(':')])]
            table_references.append('{} {}${}${}:${}${}'.format(name, prefix, *reference))
        output.append(table_references) 
    return output

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
