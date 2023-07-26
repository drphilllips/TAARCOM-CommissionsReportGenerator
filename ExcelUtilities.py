import os

import numpy as np
import pandas as pd
from xlrd import XLRDError

import EnumTypes


default_sheet_name = "Data"
default_more_sheet_name = "Customers Ranked"


def saveError(*excel_files):
    """Checks for obstacles with saving the output file

    :param: excel_files: path to output directory
    :return: whether there is an error with saving
    """
    for file in excel_files:
        try:
            open(file, 'r+')
        except FileNotFoundError:
            pass
        except PermissionError:
            return True
    return False


def createExcelFile(output_path, sheet_data, more_sheet_data=pd.DataFrame()):
    """Creates an Excel file from a dataframe

    :param output_path: path for our created file
    :param sheet_data: dataframe which will be copied to this file
    :param more_sheet_data: dataframe which will be copied as a second sheet
    :return: xlsxwriter object for formatting this file
    """

    # Verify output path
    if saveError(output_path):
        print("..One or more files are currently open in Excel!\n"
              "..Please close the files and try again.\n"
              "*Program Terminated*")
        return

    # Write the output file
    writer = pd.ExcelWriter(output_path, engine="xlsxwriter", date_format="yyyy-mm-dd", datetime_format="yyyy-mm-dd")
    sheet_data.to_excel(writer, sheet_name=default_sheet_name, index=False)

    more_sheet_data.to_excel(writer, sheet_name=default_more_sheet_name, index=False)

    print("> New file saved at: " + output_path)

    return writer


def loadLookupFile(filename, sheet_name):
    """Loads the specified sheet from the lookup file to a dataframe

    :param: filename: name of the lookup file
    :param: sheet_name: name of the main sheet we pull data from
    :return: dataframe with sheet data
    """

    # Assume file is in the lookup directory
    look_dir = "I:/Lookup/"
    filepath = look_dir + filename

    if os.path.exists(filepath):
        try:
            sheet_data = pd.read_excel(filepath, sheet_name).fillna("")
        except XLRDError:
            print("..Error reading sheet name for " + filename + "!\n"
                  "..Please make sure the main tab is named \"" + sheet_name + "\".\n"
                  "*Program Terminated*")
            return
    else:
        print("..No " + filename + " file found!\n"
              "..Please make sure " + filename + " is in the directory.\n"
              "*Program Terminated*")
        return

    return sheet_data


def formatSheet(sheet_data, sheet_name, writer, col_widths):
    """Formats our output file to make it look nice :)

    :param: sheet_data: working data frame for output
    :param: sheet_name: name of the sheet we are working on
    :param: writer: working xlsxwriter for Excel tools
    :param: col_widths: pre-defined widths of columns
    :return: void; output formatted Excel file
    """

    print("..Formatting sheet (" + sheet_name + ")..")

    # Store the working sheet from the output Excel file
    sheet = writer.sheets[sheet_name]

    # --------------------
    #  Define all formats
    # --------------------

    fmt_default = writer.book.add_format({'font': 'Calibri Light',
                                          'font_size': 8})
    fmt_left_aligned = writer.book.add_format({'font': 'Calibri Light',
                                               'font_size': 8,
                                               'align': 'left'})
    fmt_center_aligned = writer.book.add_format({'font': 'Calibri Light',
                                                 'font_size': 8,
                                                 'align': 'center'})
    fmt_right_aligned = writer.book.add_format({'font': 'Calibri Light',
                                                'font_size': 8,
                                                'align': 'right'})
    fmt_accounting = writer.book.add_format({'font': 'Calibri Light',
                                             'font_size': 8,
                                             'num_format': '$#,##0'})  # num_format 44 + rounded to nearest $
    fmt_number_with_commas = writer.book.add_format({'font': 'Calibri Light',
                                                     'font_size': 8,
                                                     'num_format': 3})

    # -------------------
    #  Format the header
    # -------------------

    header_left_cols = ['T-End Cust', 'Reported Customer']

    # Write the DataFrame header to the worksheet with the defined format
    header_row = 0
    for col_num, value in enumerate(sheet_data.columns.values):
        if value in header_left_cols:
            sheet.write(header_row, col_num, value, fmt_left_aligned)
        else:
            sheet.write(header_row, col_num, value, fmt_center_aligned)

    # Freeze header, so it remains stationary when scrolling up or down
    sheet.freeze_panes(1, 0)

    # Set auto filter
    sheet.autofilter(0, 0, sheet_data.shape[0], sheet_data.shape[1] - 1)

    # -----------------
    #  Format the body
    # -----------------

    # Ignore number stored as text error
    sheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})

    # Determine which data columns need which format
    accounting_cols = ['Revenue']
    number_with_commas_cols = ['Qty']
    definitely_text_cols = ['P/N']  # Make sure it's not interpreted as a number
    center_aligned_cols = ['FSR', 'Principal', 'Comm Month', 'Invoice Date', 'Channel', 'EM/CM']
    right_aligned_cols = []

    # -------------------------
    #  Format and size columns
    # -------------------------

    for col in sheet_data.columns:
        # Setting each column's style
        fmt = fmt_default
        if col in accounting_cols:
            fmt = fmt_accounting
        elif col in number_with_commas_cols:
            fmt = fmt_number_with_commas
        elif col in definitely_text_cols:
            fmt = fmt_default
        elif col in center_aligned_cols:
            fmt = fmt_center_aligned
        elif col in right_aligned_cols:
            fmt = fmt_right_aligned

        # Set column width and formatting
        col_idx = sheet_data.columns.get_loc(col)
        col_width = col_widths[col_idx]
        sheet.set_column(col_idx, col_idx, col_width, fmt)

    # Set the row height for all rows
    row_height = 10.8
    for row_num in range(sheet_data.shape[0]):
        sheet.set_row(row_num, row_height)



