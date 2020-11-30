
#import numpy as np
import pandas as pd
# import glob
import os

#### Combine, concatenate, join multiple excel files in a given folder into one dataframe, Each excel files having multiple sheets
#### All sheets in a single Excel file are first combined into a dataframe, then all the Excel Books in the folder
#### Are combined to make a single data frame. The combined data frame is the exported into a single Excel sheet.



### Dataframe Initialization
combiner_df = pd.DataFrame()

# filenames = glob.glob(path + "/*.xlsx")

for file in os.listdir(os.getcwd()):
    if file.endswith(".xlsx"):


        ### Get all the sheets in a single Excel File using  pd.read_excel command, with sheet_name=None
        ### Note that the result is given as an Ordered Dictionary File
        ### Hell can be found here: https://pandas.pydata.org/pandas-docs...

        # df = pd.read_excel(file, sheet_name=None,nrows=None,usecols=None,header = 0,index_col=None)
        df = pd.read_excel(file)
        # df = pd.read_excel(file, sheet_name=None, skiprows=None,nrows=None,usecols=None,header = 0,index_col=None)
        #df = pd.read_excel(file, sheet_name=None, skiprows=0,nrows=34,usecols=105,header = 9,index_col=None)

        ### Use pd.concat command to Concatenate pandas objects as a Single Table.
        combiner_df =  combiner_df.append(df)



         ### Use append command to append/stack the previous concatenated data on top of each other
        ### as the iteration goes on for every files in the folder

        # concat_all_sheets_all_files=concat_all_sheets_all_files.append(concat_all_sheets_single_file)
        # print(concat_all_sheets_all_files)

combiner_df.to_excel("master.xlsx",index=False,engine="xlsxwriter")
