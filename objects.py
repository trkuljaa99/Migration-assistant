"""Migration objects of the SAP S/4HANA Migration Cockpit ("Migrate Your Data", file / staging-table approach),
grouped by area, for each target. The names follow SAP's migration object lists; the exact set depends on the
release and scope items, so check the "Migration Objects" list of your release before a real project."""

PUBLIC = {
    "Master data – business partners": [
        "Bank", "Customer", "Customer - extend existing record by new org levels", "Supplier",
        "Supplier - extend existing record by new org levels", "Contact person",
    ],
    "Master data – products & logistics": [
        "Product", "Product - extend existing record by new org levels", "Product classification",
        "Batch", "Batch unique at plant level", "Characteristic", "Class", "Customer material",
        "Purchasing info record", "Source list", "Material BOM", "Routing", "Work center", "Production version",
        "Inspection plan", "Equipment", "Functional location", "Maintenance plan", "Maintenance task list",
    ],
    "Finance – master data": [
        "G/L account", "Cost center", "Profit center", "Activity type", "Activity price", "Exchange rate",
        "Internal order", "Bank account balance",
    ],
    "Finance – balances & open items": [
        "FI - G/L account balance and open/line item", "FI - Accounts receivable open item",
        "FI - Accounts payable open item", "Fixed asset (incl. balances and transactions)",
        "FI - Historical balance",
    ],
    "Logistics – transactional data": [
        "Material inventory balance", "Purchase order (only open PO)", "Sales order (only open SO)",
        "Purchasing contract", "Scheduling agreement", "Pricing condition (purchasing and sales)",
        "Sales contract", "Service order (only open)",
    ],
    "Projects": ["Project", "Enterprise project"],
}

PRIVATE = {
    "Master data – business partners": [
        "Bank", "Customer", "Customer - extend existing record by new org levels", "Supplier",
        "Supplier - extend existing record by new org levels", "Contact person", "Business partner relationship",
        "Credit management role for customer",
    ],
    "Master data – products & logistics": [
        "Product", "Product - extend existing record by new org levels", "Product consumption", "Batch",
        "Characteristic", "Class", "Object classification (general template)", "Customer material",
        "Purchasing info record", "Source list", "Quota arrangement", "Material BOM", "Equipment BOM",
        "Functional location BOM", "Routing", "Master recipe", "Work center", "Resource", "Production version",
        "Inspection plan", "Inspection method", "Master inspection characteristic", "Equipment",
        "Functional location", "Maintenance plan", "Equipment task list", "General maintenance task list",
        "Measuring point", "Document info record",
    ],
    "Finance – master data": [
        "G/L account", "Cost center", "Profit center", "Cost element", "Activity type", "Activity price",
        "Statistical key figure", "Exchange rate", "Internal order", "House bank", "Bank account",
        "Fixed asset - master data", "Business area",
    ],
    "Finance – balances & open items": [
        "FI - G/L account balance", "FI - G/L account open item", "FI - Accounts receivable open item",
        "FI - Accounts payable open item", "Fixed asset (incl. balances)", "FI - Historical balance",
        "CO - Secondary cost", "Asset transaction",
    ],
    "Logistics – transactional data": [
        "Material inventory balance", "Purchase order (only open PO)", "Purchase requisition (only open)",
        "Sales order (only open SO)", "Purchasing contract", "Sales contract", "Scheduling agreement",
        "Pricing condition (general template)", "Warehouse fixed bin assignment", "Handling unit",
        "Inventory in EWM",
    ],
    "Projects & services": ["Project (project definition, WBS element)", "Network", "Service contract",
                            "Service order (only open)"],
    "Industry & other": ["Contract accounting - business partner", "Real estate contract",
                         "Commodity pricing condition", "Variant configuration - configuration profile",
                         "Joint venture accounting - venture"],
}

TARGETS = {"public": ("SAP S/4HANA Cloud Public Edition", PUBLIC),
           "private": ("SAP S/4HANA Cloud Private Edition", PRIVATE)}


def catalog(target):
    return TARGETS[target][1]


def label(target):
    return TARGETS[target][0]
