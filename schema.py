"""Column schema for Daily Sales Register workbooks."""

COLUMNS = [
    "Invoice Number",
    "Transaction number (Fusion)",
    "Advertiser Name",
    "AP Vendor Code",
    "AP Name",
    "Marketplace Seller ID",
    "Diamond Seller",
    "PAN Number",
    "International client",
    "Brand Name",
    "Team Responsible",
    "Accounts Manager",
    "Agency Eco-System",
    "RO No",
    "Billing Month",
    "Service Period",
    "Invoice Date",
    "FY",
    "Force Billings and Risky Debtors",
    "USD Booking $ Rate",
    "USD $ Billing value",
    "Invoice amount",
    "SGST",
    "CGST",
    "IGST",
    "Total Amount Due",
    "Currency",
    "Collection month",
    "Bank Realization date",
    "Cheque no",
    "Bank",
    "Cheque Amt",
    "TDS",
    "Bank Charges Month",
    "Bank Charges",
]

INR_FIELDS = {
    "Invoice amount",
    "SGST",
    "CGST",
    "IGST",
    "Total Amount Due",
    "Cheque Amt",
    "TDS",
    "Bank Charges",
}

USD_FIELDS = {
    "USD Booking $ Rate",
    "USD $ Billing value",
}

KPI_FIELDS = [
    "Invoice amount",
    "Total Amount Due",
    "TDS",
    "USD $ Billing value",
]

NUMERIC_FIELDS = INR_FIELDS | USD_FIELDS

SEARCH_PRIORITY = [
    "Invoice Number",
    "Advertiser Name",
    "Brand Name",
    "AP Name",
    "Accounts Manager",
    "Team Responsible",
    "PAN Number",
    "RO No",
    "Transaction number (Fusion)",
]


def sql_name(label: str) -> str:
    out = []
    prev_us = False
    for ch in label.lower():
        if ch.isalnum():
            out.append(ch)
            prev_us = False
        elif not prev_us:
            out.append("_")
            prev_us = True
    return "".join(out).strip("_") or "col"


SQL_COLUMNS = [sql_name(c) for c in COLUMNS]
LABEL_TO_SQL = dict(zip(COLUMNS, SQL_COLUMNS))
SQL_TO_LABEL = dict(zip(SQL_COLUMNS, COLUMNS))
NUMERIC_SQL = {LABEL_TO_SQL[c] for c in NUMERIC_FIELDS}
KPI_SQL = [LABEL_TO_SQL[c] for c in KPI_FIELDS]
