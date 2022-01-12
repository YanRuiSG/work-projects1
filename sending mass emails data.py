import os
import pandas as pd
import numpy as np


wd = os.getcwd()

# create a list of available excel files in the current working drive
excel_files = [file for file in os.listdir(wd) if file.endswith('.xlsx')]

print(excel_files)


# import the raw dataset of the list of all workers to be discharged as a dataframe
main_df = pd.read_excel(excel_files[0])
# explore the dataframe and inspect the data values
print(main_df.head())
print(main_df.shape)


# import the data for the email addresses as a dataframe
email_df =  pd.read_excel(excel_files[-1])
print(email_df.info())

# Renaming column names/perform any other necessary data cleaning
emails_col = list(email_df.columns)
#display(emails_col)
email_df = email_df.rename(columns={emails_col[0]:'Company', emails_col[1]:'Name'})
print(email_df.head(5))



# create the list of unique companies to loop through
companies = sorted(list(main_df['Company'].unique()))
display(companies)



# create a list containing lists of companies in batches of 20
companies_sublist = [companies[i:i+20] for i in range(0, len(companies), 20)]


# create the excel file data for the mass email sending automation
import xlsxwriter
from datetime import date

for batch in companies_sublist:
    # create a unique list of company names for sheet names
    # only keep the first 10 elements of each company name in the list
    companies_sheetname = [name[:14] for name in batch]


    #Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('Email batch file ' + str(companies_sublist.index(batch) + 1) + '.xlsx', engine='xlsxwriter')


    # for each company, export the list of workers to be discharged to an individual excel worksheet

    for count in range(len(batch)):
        # display(companies[count])


        # subset the company from the wp_expired dataframe
        df1 = main_df[main_df['Company'] == batch[count]]

        # suset the company from the email_df dataframe
        df2 = email_df[email_df['Company'] == batch[count]]

        # join the 2 dataframe horizontally
        df3 = df1.merge(df2, how='outer', on='Company')


        # add the email subject information as a new column
        df3.to_excel(writer, sheet_name=companies_sheetname[count])


    writer.save()