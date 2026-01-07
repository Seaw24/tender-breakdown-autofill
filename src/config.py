# config.py

# Maps keywords found in filenames to the official names in the Breakdown sheet
FILENAME_TO_MASTER_NAME = {
    "Beanyard": "Beanyard",
    "Crimson Corner": "Crimson Corner",
    "UnionShakeSmart": "Union Shake Smart",
    "gardner": "Gardner",
    "Honors": "HONORS",
    "KV": "Kahlert",
    "Shake Smart": "Student Life Center",
    "Epicenter": "Epicenter",
    "cv casheet": "Crimson View",
    "Satellite": "Einsteins",
    "HECS": "Hive"
}

# The starting column index for each location in Tender Breakdown.xlsx
BREAKDOWN_LOC_MAP = {
    "PHC": 2, "Crimson View": 14, "Einsteins": 26, "Hive": 38,
    "Union Food Court": 50, "HONORS": 62, "Gardner": 74,
    "Crimson Corner": 86, "Lassonde": 98, "Beanyard": 110,
    "Student Life Center": 122, "Kahlert": 134, "Union Shake Smart": 146,
    "Epicenter": 158, "Quartzdyne": 170, "Seagull Sunrise": 182, "City's Edge": 194
}
