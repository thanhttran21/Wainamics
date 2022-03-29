# Name: Thanh T. Tran
# Date: 2/17/2022
# Task: Generate PNG of 3 channels vs. plots. 
#           Include a legend, date created, and "Additional Notes" text box

import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt

SECONDS_PER_SAMPLE = 10

def calculate_time(time_col):
    """ Calculate the intervals between each sample. Change SECONDS_PER_SAMPLE based on sampling rate. """
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

def generate_plot(date, file_name):
    """ Generate the plot with the date of creation and file name. """
    if not ".xlsx" in file_name or "~$" in file_name:
        return

    df = pd.read_excel(file_name)
    
    time = calculate_time(df.iloc[:, 0])

    # Begin plotting
    plt.figure(figsize=(15,10))

    # Plotting raw data
    plt.subplot(2, 2, 1)
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
    ax1 = plt.subplot(2, 2, 2)
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
    ax2 = plt.subplot(2, 2, 3)
    ax2.set_axis_off()
    df_other = pd.read_excel(file_name, sheet_name='Others')
    additional_notes = parse_other(df_other.iloc[:,0])
    plt.suptitle(file_name + "\n" + str(date), fontsize = 14)
    plt.text(0, 0.65, additional_notes)

    png_name = file_name[:len(file_name)-5]
    plt.savefig(png_name + '.png', dpi = 300)
    # plt.show()

def main():
    """ Loops through every .xlsx file in current directory and create tables to write to OUTPUT_FILENAME. """
    current_dir = os.getcwd()
    files = os.listdir(current_dir)

    for file_name in files:
        date = datetime.fromtimestamp(os.stat(file_name).st_ctime)
        generate_plot(date, file_name)

if __name__ == '__main__':
    main()
    print("Done! Please check PNG Files")
