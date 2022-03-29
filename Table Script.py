# Name: Thanh T. Tran
# Date: 2/17/2022
# Task: Generate a table with the following format:
#   DATE | FILENAME | COLUMN | 20 MINS | 40 MINS | 60 MINS
#                               EXP VS ACT

import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

OUTPUT_FILENAME = "Table Output.xlsx"

def create_table(date, file_name, col_b, col_c, col_d, ratio_bd, ratio_cd):
    """ Create a table with the desired formatting. """
    data = { 'Date Created': [date, np.nan, np.nan, np.nan],
                'File Name': [file_name, np.nan, np.nan, np.nan],
                'Minutes': [0, 20, 40, 60],
                'Ch 1': col_b,
                'Ch 2': col_c,
                'Ch 3' : col_d,
                'Ch1 / Ch3' : ratio_bd,
                'Ch2 / Ch3' : ratio_cd
    }
    table = pd.DataFrame(data)
    table = table.fillna('')
    return table


def calculate_intervals():
    """ Calculate the intervals between each sample. Change SECONDS_PER_SAMPLE based on sampling rate. """
    SECONDS_PER_SAMPLE = 10
    seconds_per_minute = 60
    minutes = [0, 20, 40, 60]
    intervals = [int(minutes[0] * seconds_per_minute / SECONDS_PER_SAMPLE + 1),         # Off by one
                    int(minutes[1] * seconds_per_minute / SECONDS_PER_SAMPLE),
                    int(minutes[2] * seconds_per_minute / SECONDS_PER_SAMPLE),
                    int(minutes[3] * seconds_per_minute / SECONDS_PER_SAMPLE - 1)]      # [, ) 
    return intervals

def plot_table(table):
    """ Plot the current table to inspect formatting. """
    plt.rcParams["figure.figsize"] = [10, 5]
    plt.rcParams["figure.autolayout"] = True
    fig, ax = plt.subplots()
    ax.axis("off")
    ax.set_title("Date: File Name")
    the_table = ax.table(cellText = table.values,
            colLabels = table.columns,
            loc = "center")
    the_table.auto_set_column_width(col=list(range(len(table.columns))))
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(7)
    plt.show()

def format_table(date, file_name):
    """ Format the table with the date of creation and file name. """
    if not ".xlsx" in file_name:
        return

    df = pd.read_excel(file_name)
    intervals = calculate_intervals()

    if intervals[3] > len(df.iloc[:, 1]) or intervals[3] > len(df.iloc[:, 1]) or intervals[3] > len(df.iloc[:, 1]):
        col_b = [df.iloc[intervals[0], 1], df.iloc[intervals[1], 1], df.iloc[intervals[2], 1], df.iloc[len(df.iloc[:, 1]) - 1, 1]]
        col_c = [df.iloc[intervals[0], 2], df.iloc[intervals[1], 2], df.iloc[intervals[2], 2], df.iloc[len(df.iloc[:, 2]) - 1, 2]]
        col_d = [df.iloc[intervals[0], 3], df.iloc[intervals[1], 3], df.iloc[intervals[2], 3], df.iloc[len(df.iloc[:, 3]) - 1, 3]]
    else:
        col_b = [df.iloc[intervals[0], 1], df.iloc[intervals[1], 1], df.iloc[intervals[2], 1], df.iloc[intervals[3] - 1, 1]]
        col_c = [df.iloc[intervals[0], 2], df.iloc[intervals[1], 2], df.iloc[intervals[2], 2], df.iloc[intervals[3] - 1, 2]]
        col_d = [df.iloc[intervals[0], 3], df.iloc[intervals[1], 3], df.iloc[intervals[2], 3], df.iloc[intervals[3] - 1, 3]]

    ratio_bd = [float(format(i / j, '.4f')) for i, j in zip(col_b, col_d)]
    ratio_cd = [float(format(i / j, '.4f')) for i, j in zip(col_c, col_d)]
    
    table = create_table(date, file_name, col_b, col_c, col_d, 
                                ratio_bd, ratio_cd)
    # plot_table(table)
    return table

def write_to_same_sheet(tables):
    """ Write tables into the same sheet, skipping NEW_ROW amount of rows before writing the next table."""
    start_row = 1
    NEW_ROW = 6

    for df in tables:
        if df is None:
            continue
        if start_row == 1:
            df.to_excel(OUTPUT_FILENAME)
        else:
            with pd.ExcelWriter(OUTPUT_FILENAME, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                df.to_excel(writer, startrow=start_row)
        start_row += NEW_ROW

def write_to_new_sheet(tables):
    """ Write tables to a new sheet every time. """
    start = 1
    for df in tables:
        if df is None:
            continue
        if start == 1:
            start += 1
            df.to_excel(OUTPUT_FILENAME)
        else:
            with pd.ExcelWriter(OUTPUT_FILENAME, mode="a", engine="openpyxl", if_sheet_exists="new") as writer:
                df.to_excel(writer)

def main():
    """ Loops through every .xlsx file in current directory and create tables to write to OUTPUT_FILENAME. """
    current_dir = os.getcwd()
    if os.path.isfile(os.path.join(current_dir, OUTPUT_FILENAME)):
        os.remove(os.path.join(current_dir, OUTPUT_FILENAME))

    tables = []
    files = os.listdir(current_dir)
    for file_name in files:
        date = datetime.fromtimestamp(os.stat(file_name).st_ctime)
        tables.append(format_table(date, file_name))
    
    write_to_same_sheet(tables) # UNCOMMENT TO WRITE EVERY FILE INTO THE SAME SHEET

    # write_to_new_sheet(tables) # UNCOMMENT TO WRITE EVERY FILE INTO NEW SHEETS

if __name__ == '__main__':
    main()
    print("Done! Please check " + OUTPUT_FILENAME)
