"""
Program to average the wind direction per day for each
day in a given month taken from data created by:
    field_research_hweiss_v3.py

Written by Harrison Weiss

2018
"""

import xlrd
import xlwt
import sys
import json

class average:
    m_day = 0
    m_dir = 0.0
    m_conut = 0
    m_avg = 0

def write_default_config(outfile):
    config = dict()
    config["sheet_no"] = 0
    config["day_col"] = 1
    config["avg_col"] = 3
    
    json.dump(config, outfile)
    return

def load_config():
    try:
        config_file = open("config.json", 'r')
    except FileNotFoundError:
        write_default_config(open("config.json", 'w+'))
        config_file = open("config.json", 'r')

    return json.load(config_file)

config = load_config()

day_col = int(config["day_col"])
avg_col = int(config["avg_col"])
sheet_no = int(config["sheet_no"])

out_book = xlwt.Workbook()

for file_name in sys.argv[1:]:

    # TEST 
    while sheet_no < 5:
        try:
            sheet = xlrd.open_workbook(file_name).sheet_by_index(sheet_no)
        except (xlrd.biffh.XLRDError, FileNotFoundError):
            print(
                f"There was an error loading file \"{file_name}\", skipping file . . .")
            continue
        sheet_no += 1

        # Attempt to load file from command line, skip file if invalid
        # try:
        #     sheet = xlrd.open_workbook(file_name).sheet_by_index(sheet_no)
        # except (xlrd.biffh.XLRDError, FileNotFoundError):
        #     print(
        #         f"There was an error loading file \"{file_name}\", skipping file...")
        #     continue


        print(f"WORKING ON SHEET {sheet.name}...")

        results = list()
        results_index = 0

        current_day = -1
        current_sum = 0.0
        current_count = 0

        bool_first_time = True
        for row in range(0, sheet.nrows):
            # The first row is just the labels, but we can snag the first valid hour this time around
            if row == 0:
                continue

            # If this is the first time we've encountered a valid hour in the sheet (note we had to make it past above statement)
            if bool_first_time:
                bool_first_time = False
                current_day = sheet.cell_value(row, day_col)

            # If this row has a different day, calculate the average and add the data to our results
            if sheet.cell_value(row, day_col) != current_day or row == (sheet.nrows - 1):
                results.append(average())
                results[results_index].m_day = current_day

                try:
                    results[results_index].m_avg = round(
                        current_sum / current_count)
                except ZeroDivisionError:
                    print(
                        f"Just attempted to divide by zero... Was this intentional or is there a formatting error?")
                    print(
                        f"{sheet.name} - Row: {row}, reading direction from col: {avg_col}")
                    print(
                        f"current_sum {current_sum}, current_count {current_count}, current_day {current_day}")
                
                # Increment results_index and reset "current" variables
                results_index += 1
                current_sum = 0
                current_count = 0
                current_day = sheet.cell_value(row, day_col)

            # Accounts for missing direction data
            if sheet.cell_value(row, avg_col) == "0":
                continue

            # If this is the same hour, add up the sum and increment the count
            current_sum += sheet.cell_value(row, avg_col)
            current_count += 1
            
        # Try to name the out_sheet after the original sheet
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
            # Break when we get a valid sheet name
            break

        # First row is just the column names
        out_sheet.write(0, 0, "Day")
        out_sheet.write(0, 1, "Average")

        # Fill in the results
        for i in range(0, len(results)):
            out_sheet.write(i + 1, 0, results[i].m_day)
            out_sheet.write(i + 1, 1, results[i].m_avg)

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


        