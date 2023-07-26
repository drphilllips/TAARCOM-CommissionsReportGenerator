from enum import Enum


class Customer(Enum):
    ALL = "ALL"
    T10 = "Top 10"
    T25 = "Top 25"
    T50 = "Top 50"


class Principal(Enum):
    ALL = "ALL"


class DateColumn(Enum):
    PAID = "Paid Date"
    INVOICE = "Invoice Date"
    NA = "N/A"


COMBINED_LIST = []
COMBINED_LIST.extend(Customer)
COMBINED_LIST.extend(Principal)
COMBINED_LIST.extend(DateColumn)



