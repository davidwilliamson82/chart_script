def chart_strings(number_of_tables, interval, ranges, sheet = 'Sheet1'):
    """This function takes some information about multiple excel tables 
    and generates references for multiple columns in those tables, 
    for the purpose of generating line graphs from the data."""
    from string import ascii_uppercase
    from re import match, split as re_split
    def offset(index, table_index):
        if index[0] == '$':
            return index.replace('$', '')
        if match(r'[A-Z]+', index):
            return alpha[alpha.index(index) + table_index]
        else:
            return index 
    output = []
    prefix = "='{}'!".format(sheet)
    alpha = ascii_uppercase
    beta = [item for sublist in 
            [[first + second for second in alpha] for first in alpha] 
            for item in sublist]
    alpha = list(alpha) + beta
    for table_index in range(0, number_of_tables * interval, interval):
        beta = alpha[table_index : table_index + interval]
        output.append(['{} {}${}${}:${}${}'.format(name, prefix, 
                                                   *[offset(index, table_index) 
                                                     for coordinates in [re_split(r'(\d+)', s)[:-1] 
                                                                         for s in cells.split(':')] 
                                                     for index in coordinates]) 
                       for name, cells in ranges])
    return output

def make_range(top_row, bottom_row):
    """A simple index factory. 
    Given top and bottom row values, it returns a function which accepts the column as an argument. 
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