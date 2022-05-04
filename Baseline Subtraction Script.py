# Name: Thanh T. Tran
# Date: 3/29/2022
# Task: Graph (Actual - Expected) vs. Time using a line of best fit between t_0 and t_f

import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Change STARTING_ROW for when to begin parsing data. This is to account for any initial data
STARTING_ROW = 5

# Change SECONDS_PER_SAMPLE depending on how often the machine samples data
SECONDS_PER_SAMPLE = 10

# Line of best fit times in seconds
T_0 = 60 * 5
T_F = 60 * 15

def calculate_time(time_col):
    """ Calculate the interval in seconds between each sample. Change SECONDS_PER_SAMPLE based on sampling rate. """
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

def generate_baseline_eq(ch_col):
    """ Generate a line of best fit equation of a given channel from [t_0, t_f). """
    sample_initial = int(T_0 / SECONDS_PER_SAMPLE)
    sample_final = int(T_F / SECONDS_PER_SAMPLE)
    desired_data = ch_col[sample_initial:sample_final]
    return np.polyfit(np.arange(T_0, T_F, SECONDS_PER_SAMPLE), desired_data, 1)

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

    plt.plot(time[STARTING_ROW:], df.iloc[STARTING_ROW:,1], label='Channel 1')
    plt.plot(time[STARTING_ROW:], df.iloc[STARTING_ROW:,2], label='Channel 2')
    plt.plot(time[STARTING_ROW:], df.iloc[STARTING_ROW:,3], label='Channel 3')

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

    norm_ch1 = df.iloc[STARTING_ROW:,1].astype(float) / df.iloc[STARTING_ROW,1]
    norm_ch2 = df.iloc[STARTING_ROW:,2].astype(float) / df.iloc[STARTING_ROW,2]
    norm_ch3 = df.iloc[STARTING_ROW:,3].astype(float) / df.iloc[STARTING_ROW,3]

    plt.plot(time[STARTING_ROW:], norm_ch1, label='Channel 1')
    plt.plot(time[STARTING_ROW:], norm_ch2, label='Channel 2')
    plt.plot(time[STARTING_ROW:], norm_ch3, label='Channel 3')

    plt.legend()

    # Format plot with Baseline Subtraction
    ax2 = plt.subplot(2, 2, 3)
    plt.title("Baseline Subtraction")
    plt.xlabel("Time (s)")
    plt.ylabel("Actual - Expected RFU")
    plt.xticks([0, 1200, 2400, 3600])
    plt.grid(True)

    ch1_m, ch1_b = generate_baseline_eq(df.iloc[STARTING_ROW:,1])
    ch2_m, ch2_b = generate_baseline_eq(df.iloc[STARTING_ROW:,2])
    ch3_m, ch3_b = generate_baseline_eq(df.iloc[STARTING_ROW:,3])

    ch1_eq = lambda x: ch1_m * x + ch1_b
    ch2_eq = lambda x: ch2_m * x + ch2_b
    ch3_eq = lambda x: ch3_m * x + ch3_b

    ch1_exp = [ch1_eq(x) for x in time[STARTING_ROW:]]
    ch2_exp = [ch2_eq(x) for x in time[STARTING_ROW:]]
    ch3_exp = [ch3_eq(x) for x in time[STARTING_ROW:]]

    ch1_baseline_sub = np.subtract(df.iloc[STARTING_ROW:,1], ch1_exp)
    ch2_baseline_sub = np.subtract(df.iloc[STARTING_ROW:,2], ch2_exp)
    ch3_baseline_sub = np.subtract(df.iloc[STARTING_ROW:,3], ch3_exp)

    ch1_exp_form = "Channel 1: y = " + str(round(ch1_m, 4)) + "x + " + str(round(ch1_b, 4))
    ch2_exp_form = "Channel 1: y = " + str(round(ch2_m, 4)) + "x + " + str(round(ch2_b, 4))
    ch3_exp_form = "Channel 1: y = " + str(round(ch3_m, 4)) + "x + " + str(round(ch3_b, 4))

    plt.plot(time[STARTING_ROW:], ch1_baseline_sub, label=ch1_exp_form)
    plt.plot(time[STARTING_ROW:], ch2_baseline_sub, label=ch2_exp_form)
    plt.plot(time[STARTING_ROW:], ch3_baseline_sub, label=ch3_exp_form)

    plt.legend()

    # Format plot with date and Additional Notes
    ax3 = plt.subplot(2, 2, 4)
    ax3.set_axis_off()
    df_other = pd.read_excel(file_name, sheet_name='Others')
    additional_notes = parse_other(df_other.iloc[:,0])
    plt.suptitle(file_name + "\n" + str(date), fontsize = 14)
    plt.text(0, 0, additional_notes)

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
