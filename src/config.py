# config.py

# Maps keywords found in filenames to the official names in the Breakdown sheet
"""Settings: 
{
  "filepath" : [{"Location name 1 in Breakdown sheet" : "What to look for in casheet"}},
    ...         {"Location name 2 in Breakdown sheet" : "What to look for in casheet"}] 
  }
  """

FILENAME_TO_MASTER_NAME = {
    "Beanyard": [{"Beanyard": "TOTALS"}],
    "Crimson Corner": [{"Crimson Corner": "TOTALS"}],
    "UnionShakeSmart": [{"Union Shake Smart": "TOTALS"}],
    "gardner": [{"Gardner": "TOTALS"}],
    "Honors": [{"HONORS": "TOTALS"}],
    "KV": [{"Kahlert": "TOTALS"}],
    "Shake Smart": [{"Student Life Center": "TOTALS"}],
    "Epicenter": [{"Epicenter": "TOTALS"}],
    "cv casheet": [{"Crimson View": "TOTALS"}],
    "Lassonde": [{"Lassonde": "TOTALS"}],
    "Satellite": [{"Einsteins": "TOTALS"}],
    "HECS": [{"Hive": "Hive Tavlo"}, {"Quartz": "Quartz Café"}],
    "HUB": [{"City's Edge": "TOTALS"}],
    "CRSS": [{"Seagull Sunrise": "TOTALS"}],
    "csfs": [{"Union Food Court": "TOTALS"}],
}


# Directory paths for cash sheets and master file
DIRECTORY_PATHS = {
    "casheets_dir": "cash sheets",
    "master_file": "Tender Breakdown.xlsx"
}
