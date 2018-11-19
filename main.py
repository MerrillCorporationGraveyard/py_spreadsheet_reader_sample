## import modules
import os
import numpy as np
import pandas as pd

## the main function
def main():
    sheet = input("Enter the name of the spreadsheet in the Spreadsheets folder: ")
    xd = load_file_contents(sheet) ## loads the excel data
    values = find_high_low(xd) ## finds the high and low values
    display_values(values) ## prints the values to the console

## load our excel file
def load_file_contents(sheet):
    ## if the user forgot to input the file extension, we'll add it
    if '.xlsx' not in sheet:
        sheet = sheet + '.xlsx'

    os.chdir('{}/Spreadsheet'.format(os.getcwd()))
    return pd.read_excel(sheet)

## find the high and the low value
## return: a dictionary of { high: "high value", low: "low value" }
def find_high_low(xd):
    xd['Change'] = xd['Close'] - xd['Open']
    hilo = {
        "high": find_top(xd, False),
        "low": find_top(xd, True)
    }

    return hilo

## sort the values by change and pick the top one
## return: a string value of the date column of the top sort value
def find_top(xd, asc):
    sort = xd.sort_values(['Change'], ascending=asc)
    return sort['Date'].iloc[0]

## display the data to the console
def display_values(val):
    print('\nThe largest positive change value is: {}'.format(val['high']))
    print('\nThe largest negative change value is: {}'.format(val['low']))

## alright, lets run our program...
main()
