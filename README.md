# CommissionsReportGenerator

Reads in the entire Commissions Master file and generates a report
based on several query parameters. Used for generating sales
reports/forecasts at TAARCOM, inc.

The Commissions Master file contains company commissions data, where
each line item contains information on a sale--i.e., supplier, part
number, distributor, quantity, unit price, invoiced dollars, etc.

The query parameters include customer, principal (supplier), 
date column (i.e., which column to use for selecting dates) and 
inputs for start and end dates.

**Customer** can be broken down into Top 10, Top 25, or Top 50 customers,
and the rest of the options are populated with all existing customers
within the commissions master file.

**Principal** is populated with all of our company's represented suppliers.

**Date Column** selection includes two options: invoiced date and paid date.

**Time Period** starts at the first of the current calendar year, 
and ends at the current date.
