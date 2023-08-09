"""
Program to avergae the wind direction for each hour
using data taken from NOAA ASOS airport data

Note: Special formatting is used. See tet_winds.xlsx
for template sheet. Each column must match the template
for the program to work.

Written by Harrison Weiss and Chris Laganella

2018
"""

import xlrd
import xlwt
import sys
import json


class wind_average:
    m_day = 0
    m_month = 0
    m_hour = 0.0
    m_avg = 0


def write_default_config(outfile):
    config = dict()
    config["month_col"] = 2
    config["day_col"] = 3
    config["hour_col"] = 4
    config["dir_col"] = 8
    config["sheet_no"] = 0

    json.dump(config, outfile)
    return


def load_config():
    try:
        config_file = open("config.json", 'r')
    except FileNotFoundError:
        write_default_config(open("config.json", 'w+'))
        config_file = open("config.json", 'r')

    return json.load(config_file)


""" if __name__ == __main___: """

# Load config settings and set values
config = load_config()
# print(config)

month_col = int(config["month_col"])
day_col = int(config["day_col"])
hour_col = int(config["hour_col"])
dir_col = int(config["dir_col"])
sheet_no = int(config["sheet_no"])

out_book = xlwt.Workbook()


for file_name in sys.argv[1:]:
    # Attempt to load file from command line, skip file if invalid
    try:
        sheet = xlrd.open_workbook(file_name).sheet_by_index(sheet_no)
    except (xlrd.biffh.XLRDError, FileNotFoundError):
        print(
            f"There was an error loading file \"{file_name}\", skipping file...")
        continue

    print(f"WORKING ON SHEET {sheet.name}...")

    results = list()
    results_index = 0

    current_month = -1
    current_day = -1
    current_hour = -1
    current_sum = 0.0
    current_count = 0

    bool_first_time = True
    for row in range(0, sheet.nrows):
        # The first row is just the labels, but we can snag the first valid hour this time around
        if row == 0:
            continue

        # Dont need to check for times outside of 12 and 23
        if sheet.cell_value(row, hour_col) < 12 or sheet.cell_value(row, hour_col) > 23:
            continue

        # If this is the first time we've encountered a valid hour in the sheet (note we had to make it past above statement)
        if bool_first_time:
            bool_first_time = False
            current_month = sheet.cell_value(row, month_col)
            current_day = sheet.cell_value(row, day_col)
            current_hour = sheet.cell_value(row, hour_col)
            #print(f"JUST GOT FIRST VALID HOUR {current_hour} ON ROW {row}")

        # If this row has a different hour calculate average and add the data to our results
        if sheet.cell_value(row, hour_col) != current_hour or row == (sheet.nrows - 1):
            results.append(wind_average())
            results[results_index].m_month = current_month
            results[results_index].m_day = current_day
            results[results_index].m_hour = current_hour

            try:
                results[results_index].m_avg = round(
                    current_sum / current_count)
            except ZeroDivisionError:
                print(
                    f"Just attempted to divide by zero... Was this intentional or is there a formatting error?")
                print(
                    f"{sheet.name} - Row: {row}, reading direction from col: {dir_col}")
                print(
                    f"current_sum {current_sum}, current_count {current_count}, current_hour {current_hour}")

            # Increment results_index and reset "current" variables
            results_index += 1
            current_sum = 0
            current_count = 0
            current_month = sheet.cell_value(row, month_col)
            current_day = sheet.cell_value(row, day_col)
            current_hour = sheet.cell_value(row, hour_col)

        # Accounts for missing direction data
        if sheet.cell_value(row, dir_col) == "M":
            continue
        if sheet.cell_value(row, dir_col) == "]":
            continue
        if sheet.cell_value(row, dir_col) == "][M]":
            continue
        if sheet.cell_value(row, dir_col) == "D":
            continue
        if sheet.cell_value(row, dir_col) == "][D]":
            continue
        if sheet.cell_value(row, dir_col) == "":
            continue

        # If this is the same hour just add up the sum and increment the count
        # print(sheet.cell_value(row, dir_col))
        current_sum += sheet.cell_value(row, dir_col)
        current_count += 1

    # Note - Still in outter for loop (iterating through given files)
    #out_sheet = out_book.add_sheet(sheet.name)

    # Try to name the out_sheet after the orig sheet
    out_sheet_name = sheet.name
    while True:
        try:
            out_sheet = out_book.add_sheet(out_sheet_name)
        except:
            print(f"Error naming sheet {out_sheet_name}...")
            # Add an I kinda like roman numerals
            out_sheet_name += "I"
            print(f"Renaming to {out_sheet_name}")
            continue

        # If we finally got a valid sheet name then break
        break

    # First row is just the column names
    out_sheet.write(0, 0, "Month")
    out_sheet.write(0, 1, "Day")
    out_sheet.write(0, 2, "Hour")
    out_sheet.write(0, 3, "Average")

    # Fill in the results
    for i in range(0, len(results)):
        out_sheet.write(i + 1, 0, results[i].m_month)
        out_sheet.write(i + 1, 1, results[i].m_day)
        out_sheet.write(i + 1, 2, results[i].m_hour)
        out_sheet.write(i + 1, 3, results[i].m_avg)

print()
while True:
    print("Please enter a name for the output spreadsheet.")
    print("DO NOT include .xlsx, that will be inserted for you.")
    out_book_name = input("")
    out_book_name += ".xlsx"

    try:
        out_book.save(out_book_name)
    except:
        print(f"Unexpected error while saving, {sys.exc_info()[0]}")
        print()
        continue

    break
