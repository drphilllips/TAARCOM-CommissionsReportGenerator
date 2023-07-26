
import os
import pandas as pd
import time
import subprocess

import EnumTypes
import ExcelUtilities


def main(cms_df, output_path, customer, principal, date_column, start_date, end_date):
    """
    Run.main executes "running a report" over TAARCOM's Commissions
    Master file based on several query options

    :param cms_df: loaded DataFrame of selected Commissions file
    :param output_path: filepath for the output report
    :param customer: drop-down selection for customer query
    :param principal: drop-down selection for principal query
    :param date_column: selected date column to use for time period query (invoice date, paid date, or n/a)
    :param start_date: first date of time interval for query
    :param end_date: last date of time interval for query
    :return: void; export, format, and open generated report
    """

    print("..Running report..")

    # -----------------------------
    #  Create Report (Output) File
    # -----------------------------

    # Load ReportColumns.xlsx
    rcl_df = ExcelUtilities.loadLookupFile("ReportColumns.xlsx", "Columns")

    # Pull desired commissions columns and their preferred names from Root Column Library
    actual_cols = list(rcl_df.columns)
    preferred_cols = list(rcl_df.iloc[0])

    # Pull desired column widths, including those for the "Customers Ranked" sheet
    col_widths = list(rcl_df.iloc[1])
    cust_rank_col_widths = list(rcl_df[['T-End Cust', 'Paid-On Revenue']].iloc[1])

    # Populate report dataframe with values from all desired columns from commissions dataframe
    rpt_df = cms_df[actual_cols]
    rpt_df.columns = preferred_cols

    # Create query text
    main_query = ''

    # --------------------
    #  Time Period Query       // Filter by time period first, as this should reduce the dataset most
    # --------------------

    if date_column == EnumTypes.DateColumn.PAID:
        main_query += '`Comm Month` >= "' + str(start_date) + '" and `Comm Month` <= "' + str(end_date) + '"'
    elif date_column == EnumTypes.DateColumn.INVOICE:
        main_query += '`Invoice Date` >= "' + str(start_date) + '" and `Invoice Date` <= "' + str(end_date) + '"'
    elif date_column == EnumTypes.DateColumn.NA:
        pass

    # -----------------
    #  Principal Query
    # -----------------

    # Convert full name to abbreviation to add to the query
    pcp_active = ExcelUtilities.loadLookupFile(filename="principalList.xlsx", sheet_name="Principals")
    pcp_inactive = ExcelUtilities.loadLookupFile(filename="principalList.xlsx", sheet_name="Inactive")
    pcp_lookup = pd.concat([pcp_active, pcp_inactive])
    dict_principal_to_abbrev = dict(zip(pcp_lookup['Principal'], pcp_lookup['Abbreviation']))
    abbreviation = dict_principal_to_abbrev.get(principal)

    if principal == EnumTypes.Principal.ALL:
        pass
    else:
        main_query += (' & ' if main_query else '') + 'Principal == "' + abbreviation + '"'

    # -----------------------
    #  Manual Customer Query
    # -----------------------

    if customer == EnumTypes.Customer.ALL:
        pass
    elif not isinstance(customer, EnumTypes.Customer):
        main_query += (' & ' if main_query else '') + '`T-End Cust` == "' + customer + '"'

    # --------------------
    #  Execute Main Query
    # --------------------

    # Use the df.query function
    if main_query:
        rpt_df = rpt_df.query(main_query, engine='python')

    # -----------------------
    #  Ranked Customer Query
    # -----------------------

    # Create sorted list of customers by their total paid-on revenue
    grouped_df = rpt_df[['T-End Cust', 'Revenue']]

    # Remove all strings from the Paid-On Revenue column
    grouped_df['Revenue'] = pd.to_numeric(grouped_df['Revenue'], errors='coerce')
    grouped_df = grouped_df.dropna(subset=['Revenue'])

    # Roll-up data based on paid-on revenue
    grouped_df = grouped_df.groupby('T-End Cust')['Revenue'].sum().reset_index()

    # Sort customers from most to least total paid-on revenue
    sorted_df = grouped_df.sort_values(by='Revenue', ascending=False)

    # Define which customers we keep (assume it is all customers, then work down from there)
    num_cust = sorted_df.shape[0]
    if num_cust > 10 and customer == EnumTypes.Customer.T10:
        num_cust = 10
    elif num_cust > 25 and customer == EnumTypes.Customer.T25:
        num_cust = 25
    elif num_cust > 50 and customer == EnumTypes.Customer.T50:
        num_cust = 50

    # Filter customers to top-n
    sorted_df = sorted_df.iloc[:num_cust]
    keep_cust = sorted_df['T-End Cust'].values.tolist()

    # Reduce file
    rpt_df = rpt_df[rpt_df['T-End Cust'].isin(keep_cust)]

    # ---------------------
    #  Export Final Report
    # ---------------------

    # Change date columns to string because xlsx writer can't format them... for some weird reason I can't find
    rpt_df['Invoice Date'] = rpt_df['Invoice Date'].astype(str).replace("NaT", "")
    rpt_df['Comm Month'] = rpt_df['Comm Month'].astype(str).replace("NaT", "")

    # Create file
    writer = ExcelUtilities.createExcelFile(output_path, rpt_df, sorted_df)

    if writer:
        # Format columns in Excel
        ExcelUtilities.formatSheet(rpt_df, "Data", writer, col_widths)
        ExcelUtilities.formatSheet(sorted_df, "Customers Ranked", writer, cust_rank_col_widths)

        # Save the file
        writer.save()

        # Success message
        print("> File successfully saved!")

        # Open the Excel file
        excel_app_path = 'C:/Program Files (x86)/Microsoft Office/Office14/EXCEL.EXE'
        subprocess.Popen([excel_app_path, output_path])
    else:
        print("> File NOT successfully saved.\n"
              "> Make sure to close all files with matching names in the Output directory.")
