import time
import pandas as pd
from helpers import convert_to_xlsx, combine_workbooks, copy_data, filter_data_to_new_ws, add_column

src_files = [
    # src file paths
]
# destination file path'
dest_file = ''


start =  time.time()

# combine all workbooks into a single worksheet
combine_workbooks(convert_to_xlsx(src_files), dest_file)

# filter the combined worksheet and copy the filtered data into a new worksheet
def cond1(x):
    return x == 'Link clicked'
def cond2(x):
    return x == 'Actual user activities'
filter_condition1 = [
    ('Response', cond1),
    ('Sandbox Detected Status', cond2)
]
filter_data_to_new_ws(dest_file, 'Combined Data', 'Filtered data',filter_condition1)

# add Duplicates column and count the number of appearances of each email
add_column(dest_file, 'Filtered data', 'Duplicates')

# filter Duplicates column and add to new worksheet Phished
def cond3(x):
    return x == 1
filter_condition2 = [
    ('Duplicates', cond3)
]
filter_data_to_new_ws(dest_file, 'Filtered data', 'Phished', filter_condition2)

# filter Duplicates column and add to new worksheet Duplicates
def cond4(x):
    return x != 1
filter_condition3 = [
    ('Duplicates', cond4)
]
filter_data_to_new_ws(dest_file, 'Filtered data', 'Duplicates', filter_condition3)
end = time.time()
print(end - start)
