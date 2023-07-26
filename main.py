import os
import sys

import numpy as np
import pandas as pd

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import pyqtSlot, QDate
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog
from PyQt5.uic import loadUi

import EnumTypes
import ExcelUtilities
import Run

VERSION = "Beta v1.0"


class Stream(QtCore.QObject):
    """Redirects console output to text widget"""
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))

    # Pass the flush so we don't get an attribute error.
    def flush(self):
        pass


class MainWindow(QDialog):
    """Generates the main window for our program"""

    def __init__(self):
        super(MainWindow, self).__init__()

        # External UI design w/ QTDesigner ;)
        loadUi("CRG.ui", self)

        # Initialize the threadpool for handling worker jobs
        self.threadpool = QtCore.QThreadPool()

        # State variables
        self.filepath = ""
        self.cms_df = pd.DataFrame()

        # Connect GUI buttons to methods
        self.btnSelectFile.clicked.connect(self.selectFile)
        self.btnDeselectFile.clicked.connect(self.deselectFile)
        self.btnClearConsole.clicked.connect(self.clearConsole)
        self.btnRun.clicked.connect(self.runClicked)

        # Initialize query option date edits and drop-downs
        self.initializeQueryOptions()

        # Disable Run button and drop-downs until file selected
        self.lockButtons()
        self.btnClearConsole.setEnabled(True)
        self.btnSelectFile.setEnabled(True)

        # Custom output stream
        sys.stdout = Stream(newText=self.writeToConsole)

        # Show welcome message
        self.clearConsole()

    # ------------------------
    #  One Excel Op at a Time
    # ------------------------

    def runClicked(self):
        """Send the run execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.run)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    # -------------------------
    #  Execute Operation Files
    # -------------------------

    def run(self):
        """Runs function for run (report)"""

        # Check if we have the necessary lookup files
        look_dir = "I:/Lookup/"
        rcl_exists = os.path.exists(look_dir + "ReportColumns.xlsx")  # Root Column Library for CRG

        # Only run if all lookup files can be found
        if self.filepath and rcl_exists:
            # Run the Run.py file.
            try:
                # Store values for drop-down options into variables
                customer = self.getEnumType(self.drpdwnCustomer)
                principal = self.getEnumType(self.drpdwnPrincipal)
                date_column = self.getEnumType(self.drpdwnDateColumn)

                # Store start and end dates as Python datetime objects
                start_date = self.dateStartDate.date().toPyDate()
                end_date = self.dateEndDate.date().toPyDate()

                # Isolate the input file's name
                filename = os.path.basename(self.filepath).split(".xls")[0]

                # Convert full name to abbreviation for the file's unique principal tag
                pcp_active = ExcelUtilities.loadLookupFile(filename="principalList.xlsx", sheet_name="Principals")
                pcp_inactive = ExcelUtilities.loadLookupFile(filename="principalList.xlsx", sheet_name="Inactive")
                pcp_lookup = pd.concat([pcp_active, pcp_inactive])
                dict_principal_to_abbrev = dict(zip(pcp_lookup['Principal'], pcp_lookup['Abbreviation']))
                abbreviation = dict_principal_to_abbrev.get(principal)

                # Create default unique name for file
                uq_tag = "{"
                uq_tag += (customer.name if isinstance(customer, EnumTypes.Customer) else customer[0:3]) + "-"
                uq_tag += (principal.name if isinstance(principal, EnumTypes.Principal) else abbreviation) + "-"
                uq_tag += date_column.name + ("-" if date_column != EnumTypes.DateColumn.NA else "")
                uq_tag += (start_date.strftime("%m.%d.%y") + "-") if date_column != EnumTypes.DateColumn.NA else ""
                uq_tag += end_date.strftime("%m.%d.%y") if date_column != EnumTypes.DateColumn.NA else ""
                uq_tag += "}"

                # Automatically output to Output directory
                output_path = "I:/Output/" + filename + "_" + uq_tag + ".xlsx"

                Run.main(self.cms_df, output_path, customer, principal, date_column, start_date, end_date)

            except Exception as error:
                print("..Unexpected Python error:\n" +
                      "?" + str(error) + "\n" +
                      "..Please contact your local coder.")
            # Clear file.
            self.unlockButtons()
        elif not self.filepath:
            print("..No Commissions file selected!\n"
                  "..Use the Select File button to select files.")
        elif not rcl_exists:
            print("..File ReportColumns.xlsx not found!\n"
                  "..Please check file location and try again.")

    # -----------------------
    #  GUI Utility Functions
    # -----------------------

    def getEnumType(self, drpdwn):
        """Converts drop-down option text to an enum type
           If not enum type, then returns the actual text"""

        drpdwn_txt = drpdwn.currentText()

        # Check values under each unique enum type SO you don't check same value across different enum types
        if drpdwn == self.drpdwnCustomer:
            for x in EnumTypes.Customer:
                if drpdwn_txt == x.value:
                    return x
        elif drpdwn == self.drpdwnPrincipal:
            for x in EnumTypes.Principal:
                if drpdwn_txt == x.value:
                    return x
        elif drpdwn == self.drpdwnDateColumn:
            for x in EnumTypes.DateColumn:
                if drpdwn_txt == x.value:
                    return x

        # If we can't find it, return the original text
        return drpdwn_txt

    def lockButtons(self):
        """Disable user interaction"""

        self.btnSelectFile.setEnabled(False)
        self.btnDeselectFile.setEnabled(False)
        self.btnClearConsole.setEnabled(False)
        self.btnRun.setEnabled(False)
        self.drpdwnCustomer.setEnabled(False)
        self.drpdwnPrincipal.setEnabled(False)
        self.drpdwnDateColumn.setEnabled(False)
        self.dateStartDate.setEnabled(False)
        self.dateEndDate.setEnabled(False)

    def unlockButtons(self):
        """Enable user interaction"""

        self.btnSelectFile.setEnabled(True)
        self.btnDeselectFile.setEnabled(True)
        self.btnClearConsole.setEnabled(True)
        self.btnRun.setEnabled(True)
        self.drpdwnCustomer.setEnabled(True)
        self.drpdwnPrincipal.setEnabled(True)
        self.drpdwnDateColumn.setEnabled(True)
        self.dateStartDate.setEnabled(True)
        self.dateEndDate.setEnabled(True)

    def writeToConsole(self, text):
        """Write console output to text widget."""

        cursor = self.txtConsole.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.txtConsole.setTextCursor(cursor)
        self.txtConsole.ensureCursorVisible()

    def clearConsole(self):
        """Clear console print statements"""

        self.txtConsole.clear()
        print("> Welcome to the TAARCOM, Inc. Commissions Report Generator Program!")
        print("> Make sure to pull the latest version from GitHub!")

    def selectFile(self):
        """Select file for Excel operations"""

        # Lock buttons until process concludes
        self.lockButtons()

        # Let user know the old selection is cleared
        if self.filepath:
            self.filepath = ""
            self.cms_df = pd.DataFrame()
            print("..Selecting new file, old selection cleared..")

        # Print before the open file dialog takes over runtime
        print("..Loading file..")

        # Grab Excel file for operations
        self.filepath, _ = QFileDialog.getOpenFileName(self, directory="I:/",
                                                       filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Make sure user doesn't cancel
        if self.filepath:
            # Store selected file into a dataframe
            self.cms_df = pd.read_excel(self.filepath, sheet_name=0).fillna("")

            # Populate drop-down options
            self.populateQueryOptions()

            # Print out the selected filename
            filename = os.path.basename(self.filepath)
            print("> File load complete: " + filename)

            # Shorten filename if too long for selected files label
            if len(filename) > 43:
                filename = filename[:43] + "..."

            # Update current file label
            self.lblSelectedFile.setText("> " + filename)

            # Enable buttons and drop-downs now that file is selected (or even if not selected)
            self.unlockButtons()
        elif not self.filepath:
            self.deselectFile()
            print("..Select file operation cancelled..")

    def deselectFile(self):
        """Deselect file and adjust GUI accordingly"""

        if self.filepath:
            self.filepath = ""
            self.cms_df = pd.DataFrame()
            self.lblSelectedFile.setText("<No File Selected>")
            print("> File selection cleared.")

        # Reset drop down and date options to their defaults
        self.initializeQueryOptions()

        # Disable buttons (except for clear console) and drop-downs now that file is deselected
        self.lockButtons()
        self.btnClearConsole.setEnabled(True)
        self.btnSelectFile.setEnabled(True)

    def initializeQueryOptions(self):
        """Initializes all drop-down options with their default values and enum types"""

        # Clear drop-down options
        self.drpdwnCustomer.clear()
        self.drpdwnPrincipal.clear()
        self.drpdwnDateColumn.clear()

        # Add enum type options... super keys (that's a cool name for 'em!)
        for x in EnumTypes.Customer: self.drpdwnCustomer.addItem(x.value)
        for x in EnumTypes.Principal: self.drpdwnPrincipal.addItem(x.value)
        for x in EnumTypes.DateColumn: self.drpdwnDateColumn.addItem(x.value)

        # Set start date as Jan 1st of current year, end date as current date
        current_date = QDate.currentDate()
        current_year = current_date.year()
        first_day_of_year = QDate(current_year, 1, 1)
        self.dateStartDate.setDate(first_day_of_year)
        self.dateEndDate.setDate(current_date)

    def populateQueryOptions(self):
        """Use selected commissions file program to populate drop-down options
           Occurs immediately after file selection
           We can assume that only one file has been selected"""

        # Check if we have the necessary lookup files
        look_dir = "I:/Lookup/"
        pcp_exists = os.path.exists(look_dir + "principalList.xlsx")  # Map principal abbrev to full name

        # Make sure we have a selected file
        if self.filepath and pcp_exists:
            # Initialize all inputs with default values and enum types
            self.initializeQueryOptions()

            # Read in Root Column Library (ReportColumns.xlsx)
            rcl_df = ExcelUtilities.loadLookupFile("ReportColumns.xlsx", "Columns")
            required_columns = rcl_df.columns

            # Make sure our commissions file contains all columns required for the report (i.e., the rcl_df cols)
            all_req_cols = True
            for col in required_columns:
                if col not in self.cms_df.columns:
                    all_req_cols = False

            if all_req_cols:
                # Convert all Q#YYYY date data to YYYY-mm
                def convert_quarter(date_str):
                    return date_str.str.replace(r'Q(\d)(\d{4})',
                                                lambda x: f'{int(x.group(2))}-{3 * int(x.group(1)) - 2:02d}-01')

                # Define date columns
                date_cols = ['Invoice Date', 'Comm Month']

                # Convert cols to string and apply convert quarter function
                self.cms_df[date_cols] = self.cms_df[date_cols].astype(str)
                self.cms_df[date_cols].apply(convert_quarter)

                # Convert date cols back to datetime
                self.cms_df[date_cols] = self.cms_df[date_cols].apply(pd.to_datetime, errors='coerce')

                # Reduce data to only required columns
                rpt_df = self.cms_df[required_columns]

                # Populate principal column
                # 1. get all unique 3-letter abbreviations from principal cols
                principal_options = rpt_df['Principal'].unique()
                # 2. convert abbreviations to full company names for drop-down options
                pcp_active = ExcelUtilities.loadLookupFile(filename="principalList.xlsx", sheet_name="Principals")
                pcp_inactive = ExcelUtilities.loadLookupFile(filename="principalList.xlsx", sheet_name="Inactive")
                pcp_lookup = pd.concat([pcp_active, pcp_inactive])
                dict_abbrev_to_pcp = dict(zip(pcp_lookup['Abbreviation'], pcp_lookup['Principal']))
                principal_options = [dict_abbrev_to_pcp.get(abbrev) for abbrev in principal_options]
                self.drpdwnPrincipal.addItems(sorted(principal_options))

                # Populate customer column
                customer_options = sorted(rpt_df['T-End Cust'].unique())
                self.drpdwnCustomer.addItems(customer_options)
            else:
                print("..Required columns not found.\n"
                      "..Make sure to select a commissions file with all the required columns for the report.")
                self.cms_df = pd.DataFrame()
                print("> File selection cleared.")
        elif not self.filepath:
            print("..Cannot populate drop-down options, no file is selected..")
        elif not pcp_exists:
            print("..File principalList.xlsx not found!\n"
                  "..Please check file location and try again.")


class Worker(QtCore.QRunnable):
    """Inherits from QRunnable to handle worker thread.

    param args -- Arguments to pass to the callback function.
    param kwargs -- Keywords to pass to the callback function.
    """

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    @pyqtSlot()
    def run(self):
        """Initialize the runner function with passed args, kwargs."""
        self.fn(*self.args, **self.kwargs)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # widget container for QT Designer UI
    widget = QtWidgets.QStackedWidget()
    widget.setWindowTitle("Commissions Report Generator (" + VERSION + ")")
    main_window = MainWindow()
    widget.addWidget(main_window)
    widget.setFixedWidth(900)
    widget.setFixedHeight(600)
    widget.show()

try:
    sys.exit(app.exec_())
except:
    print("..Exiting")
