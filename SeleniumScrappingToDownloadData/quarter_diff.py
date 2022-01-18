import os
import pandas as pd

municipalities = os.listdir('data')
data = {}
for municipality in municipalities:
    data[municipality] = os.listdir(fr'data\{municipality}')

for municipality in municipalities:
    quarters = []
    for raw_quarter in data[municipality]:
        quarter_split = raw_quarter.split('_')
        last_month, _ = quarter_split[4].split('.')
        quarter = f'{quarter_split[2]} {quarter_split[3]} {last_month}'
        quarters.append(quarter)
    data[municipality] = quarters

all_quarters = ['2008 ENE MAR', '2008 ABR JUN', '2008 JUL SEP', '2008 OCT DIC', '2009 ABR JUN', '2009 ENE MAR',
                '2009 JUL SEP', '2009 OCT DIC', '2010 ABR JUN', '2010 ENE MAR', '2010 JUL SEP', '2010 OCT DIC',
                '2011 ABR JUN', '2011 ENE MAR', '2011 JUL SEP', '2011 OCT DIC', '2012 ABR JUN', '2012 ENE MAR',
                '2012 JUL SEP', '2012 OCT DIC', '2013 ABR JUN', '2013 ENE MAR', '2013 JUL SEP', '2013 OCT DIC',
                '2014 ABR JUN', '2014 ENE MAR', '2014 JUL SEP', '2014 OCT DIC', '2015 ABR JUN', '2015 ENE MAR',
                '2015 JUL SEP', '2015 OCT DIC', '2016 ABR JUN', '2016 ENE MAR', '2016 JUL SEP', '2016 OCT DIC',
                '2017 ABR JUN', '2017 ENE MAR', '2017 JUL SEP', '2017 OCT DIC', '2018 ABR JUN', '2018 ENE MAR',
                '2018 JUL SEP', '2018 OCT DIC']
presence_of_data = pd.DataFrame(columns=all_quarters, index=municipalities)
for municipality in municipalities:
    res = pd.Series()
    for quarter in all_quarters:
        if data[municipality].__contains__(quarter):
            res[quarter] = 'Exists'
        else:
            res[quarter] = 'Undefined'
    presence_of_data.loc[municipality] = res

writer = pd.ExcelWriter('available_years_quarters.xlsx', engine='xlsxwriter')
presence_of_data.to_excel(writer, startrow=1, sheet_name='Sheet1')
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Iterate through each column and set the width == the max length in that column. A padding length of 2 is also added.
for i, col in enumerate(presence_of_data.columns):
    # find length of column i
    column_len = presence_of_data[col].astype(str).str.len().max()
    # Setting the length if the column header is larger
    # than the max column value length
    column_len = max(column_len, len(col)) + 2
    # set the column length
    worksheet.set_column(i, i, column_len)
writer.save()
