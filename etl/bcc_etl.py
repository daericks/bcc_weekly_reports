import xlrd
import pandas as pd
from os import listdir

# Functions
# Find date from filename
def dateFromFilename(file='20200120_report.xlsx'):
    # The filename is always a date followed either by '_report.xlsx' or '_update.xlsx'
    date = file.replace('_report.xlsx', '').replace('_update.xlsx', '')
    return date

# Return workbook and desired sheet objects from filename
def openBCCfile(file = '20200120_report.xlsx'):
    date = dateFromFilename(file)
    data_dir = '~/projects/bcc_weekly_reports/data/'
    loc = (data_dir + file)

    print('Opening {}'.format(loc))
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_name('COA Stats')
    print('Rows: {}, Cols: {}'.format(sheet.ncols, sheet.nrows))
    
    return wb, sheet

# Function to read COA data from files
def importCOAdata(table='totals', file = '20200120_report.xlsx'):
    '''
    Return a df of one of the 3 tables found in the COA weekly report
    '''
    # Which table to import
    if table == 'totals':
        rows = range(1,3)
    elif table == 'product_category':
        rows = range(4,9)
    elif table == 'fail_category':
        rows = range(9,sheet.nrows-1)

    # Create dataframe from table
    data = []
    for row in rows:
        data.append(sheet.row_values(row))
    data_df = pd.DataFrame(data=data[1:], columns=data[0], dtype='int')
    
    # Create date column and make dtype=datetime
    date = dateFromFilename(file)
    data_df['Date'] = pd.to_datetime(date)

    # Add 'Tested Batches' to fail_category'
    if table == 'fail_category':
        tested_batches = sheet.cell_value(2, 1)
        data_df['Tested Batches'] = int(tested_batches)
    
    return data_df

# Function to rename Columns. (Make new column from old, delete old column)
def removeCol(df, column, newname):
    try: 
        df[column]
    except: 
        return df
    else:
        df[newname] = df[column]
        del df[column]
        return df

# Some numeric values have unwanted characters (',') in them and are thus 'objects'
# Convert these to ints
def makeFloat(data):
    valid = '1234567890.' #valid characters for a float
    try: 
        ''.join(filter(lambda char: char in valid, data))
    except: 
        return data
    else: 
        return int(''.join(filter(lambda char: char in valid, data)))

# Sometimes unwanted characters appear in our strings (',', ':', etc)
# Remove specific string from 'Category' strings
def stringRemove(data, string):
    try: 
        data.replace(string,'')
    except: 
        return data
    else: 
        return data.replace(string,'')


# ----------- End of Functions -----------
# ----------- Begin script -----------


# Generate list of files
data_dir2 = '../../data/'
files = [f for f in listdir(data_dir2) if '.xlsx' in f]

# Use file list to generate list of dates
dates=[]
for file in files: 
    dates.append(file.replace('_report.xlsx', '').replace('_update.xlsx',''))
dates = sorted(dates)

# Read all data into dataframes
totals = pd.DataFrame()
product_categories = pd.DataFrame()
fail_categories = pd.DataFrame()
for file in files: 
    # Read file
    wb, sheet = openBCCfile(file=file)
    # Read tables and save as dataframes
    totals = totals.append(importCOAdata(table='totals', file=file))
    product_categories = product_categories.append(importCOAdata(table='product_category', file=file))
    fail_categories = fail_categories.append(importCOAdata(table='fail_category', file=file))

# Remove empty columns
del totals['']
del product_categories['']
del fail_categories['']

# ----------- Totals df -----------
# ----------- Totals df -----------
# Clean up totals df
del totals['Certificates of Analysis Received'] # This column is identical to 'Tested Batches'. Don't need both
totals['Percent Failed'] = 100* totals['Failed Batches'] / totals['Tested Batches']

# Save totals df
data_save_path = '../../etl_data/'
print('Saving totals.csv to ', data_save_path)
totals.to_csv(path_or_buf='../../etl_data/totals.csv', index=False)
print('Done.')


# ----------- Product_Categories df -----------
# ----------- Product_Categories df -----------
# Rename Columns
product_categories = removeCol(product_categories, 'Tested Batches By Category', newname='Category')
product_categories = removeCol(product_categories, 'Failed Batches By Category', newname='Failed Batches')

# Some values has ',' in them and are thus 'objects'
# Convert these to ints
product_categories['Failed Batches'] = product_categories['Failed Batches'].apply(makeFloat)
product_categories['Failed Batches'] = product_categories['Failed Batches'].astype('int')

# Remove specific string from 'Category' strings
product_categories['Category'] = product_categories['Category'].apply(stringRemove, args=':')
product_categories['Category'] = product_categories['Category'].apply(stringRemove, args=',')

# Calculated columns
product_categories['Percent Failed'] = 100* product_categories['Failed Batches'] / product_categories['Tested Batches']

# Aggregate for more normalizations
pc_totals_cols = ['Date', 'Tested Batches', 'Failed Batches']
pc_totals = product_categories[pc_totals_cols].groupby(by='Date').sum()/2
# Rename Aggregated Columns
pc_totals = removeCol(pc_totals, 'Failed Batches', newname='Total Failed')
pc_totals['Total Failed'] = pc_totals['Total Failed'].astype('int')
pc_totals = removeCol(pc_totals, 'Tested Batches', newname='Total Tested')
pc_totals['Total Tested'] = pc_totals['Total Tested'].astype('int')
# Merge with Category Data
product_categories = pc_totals.merge(product_categories, left_index=True, right_on='Date', how='right')
product_categories['Percent of Failures'] = 100 * product_categories['Failed Batches'] / product_categories['Total Failed']
product_categories['Percent Tested'] = 100 * product_categories['Tested Batches'] / product_categories['Total Tested']

# Save to file
print('Saving product_categories.csv to ', data_save_path)
product_categories.to_csv(path_or_buf='../../etl_data/product_categories.csv', index=False)
print('Done.')


# ----------- Fail_Categories df -----------
# ----------- Fail_Categories df -----------
#Rename columns
fail_categories = removeCol(fail_categories, 'Failed Batches By Category', newname='Failed Batches')
fail_categories = removeCol(fail_categories, '*Reasons For Failure', newname='Failure Reason')

# Remove ':'
fail_categories['Failure Reason'] = fail_categories['Failure Reason'].apply(stringRemove, args=':')
# Remove '*'
# Revisit later
fail_categories['Failure Reason'] = fail_categories['Failure Reason'].apply(stringRemove, args='*')

# Convert 'Failed Batches' to int
fail_categories['Failed Batches'] = fail_categories['Failed Batches'].apply(makeFloat)
fail_categories['Failed Batches'] = fail_categories['Failed Batches'].astype('int')

# Calculated columns
# 'Total Failed'
fc_totals = fail_categories.groupby(by='Date').sum()/2
fc_totals = removeCol(fc_totals, 'Failed Batches', newname='Total Failed')
fc_totals['Total Failed'] = fc_totals['Total Failed'].astype('int')
del fc_totals['Tested Batches']
fail_categories = fc_totals.merge(fail_categories, left_index=True, right_on='Date', how='right')
fail_categories['Percent of Failures'] = 100 * fail_categories['Failed Batches'] / fail_categories['Total Failed']
# Failure Rate in Percent
fail_categories['Failure Rate in Percent'] = 100 * fail_categories['Failed Batches'] / fail_categories['Tested Batches']

print('Saving fail_categories.csv to ', data_save_path)
fail_categories.to_csv(path_or_buf='../../etl_data/fail_categories.csv', index=False)
print('Done.')

print('NOTE:')
print('These value tables are cumulative b/c that is what BCC provides. Next steps are to break down into weekly rates.')