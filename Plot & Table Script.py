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

def calculate_time(time_col):
    """ Calculate the intervals between each sample. Change SECONDS_PER_SAMPLE based on sampling rate. """
    SECONDS_PER_SAMPLE = 10
    time = []
    sum = 0
    for t in time_col:
        time.append(sum)
        sum += SECONDS_PER_SAMPLE
    return time

def parse_other(df_other):
    """ Returns a string of Additional Notes and everything after. """
    additional_notes = ""
    found = False
    for cell in df_other:   
        if "Additional Notes" in cell:
            found = True
        if found:
            additional_notes += cell + "\n"
    return additional_notes

def generate_plot(date, file_name, table):
    """ Generate the plot with the date of creation and file name. """
    if not ".xlsx" in file_name or "~$" in file_name:
        return

    df = pd.read_excel(file_name)
    
    time = calculate_time(df.iloc[:, 0])

    # Begin plotting
    fig = plt.figure(figsize=(15,20))
    gs = fig.add_gridspec(4,5)
    plt.suptitle(file_name + "\n" + str(date), fontsize = 14)

    # Plotting raw data
    ax0 = fig.add_subplot(gs[0,0:2])
    plt.title("Raw")
    plt.xlabel("Time (s)")
    plt.ylabel("RFU")
    plt.xticks([0, 1200, 2400, 3600])
    plt.grid(True)

    plt.plot(time, df['Channel 1'], label='Channel 1')
    plt.plot(time, df['Channel 2'], label='Channel 2')
    plt.plot(time, df['Channel 3'], label='Channel 3')

    plt.legend()

    # Plotting normalized data
    ax1 = fig.add_subplot(gs[0,2:4])
    plt.title("Normalized")
    plt.xlabel("Time (s)")
    plt.ylabel("RFU Gain")
    plt.xticks([0, 1200, 2400, 3600])
    ax1.set_ylim(bottom=1, top=6)
    plt.yticks([1, 2, 3, 4, 5, 6])
    plt.grid(True)

    norm_ch1 = df.iloc[:,1].astype(float) / df.iloc[0,1]
    norm_ch2 = df.iloc[:,2].astype(float) / df.iloc[0,2]
    norm_ch3 = df.iloc[:,3].astype(float) / df.iloc[0,3]

    plt.plot(time, norm_ch1, label='Channel 1')
    plt.plot(time, norm_ch2, label='Channel 2')
    plt.plot(time, norm_ch3, label='Channel 3')

    plt.legend()

    # Format plot with date and Additional Notes
    ax2 = fig.add_subplot(gs[0,4])
    ax2.set_axis_off()
    df_other = pd.read_excel(file_name, sheet_name='Others')
    additional_notes = parse_other(df_other.iloc[:,0])
    plt.text(0, 0.65, additional_notes)

    # Format plot with table
    ax3 = fig.add_subplot(gs[1,:])
    ax3.set_axis_off()
    plt.title("RFU Comparisons")
    cropped_table = table.iloc[:,2:]
    the_table = ax3.table(cellText=cropped_table.values, colLabels=cropped_table.columns, loc='center')
    the_table.scale(1, 2)

    # Save PNG
    png_name = file_name[:len(file_name)-5]
    plt.savefig(png_name + '.png', dpi = 300)
    # plt.show()

def main():
    """ Loops through every .xlsx file in current directory and create tables to write to OUTPUT_FILENAME. """
    current_dir = os.getcwd()
    if os.path.isfile(os.path.join(current_dir, OUTPUT_FILENAME)):
        os.remove(os.path.join(current_dir, OUTPUT_FILENAME))

    tables = []
    files = os.listdir(current_dir)
    for file_name in files:
        date = datetime.fromtimestamp(os.stat(file_name).st_ctime)
        table = format_table(date, file_name)
        tables.append(table)
        generate_plot(date, file_name, table)
    
    write_to_same_sheet(tables) # UNCOMMENT TO WRITE EVERY FILE INTO THE SAME SHEET

    # write_to_new_sheet(tables) # UNCOMMENT TO WRITE EVERY FILE INTO NEW SHEETS

if __name__ == '__main__':
    main()
    print("Done! Please check PNG files and " + OUTPUT_FILENAME)
