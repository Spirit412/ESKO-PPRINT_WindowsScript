from itertools import groupby
elements = ['punkt', 'punkt', 'punkt', 'punkt', 'punkt', 'punkt' ]
for key, group in groupby(elements):
    print(key)