import re
 
def extract_rpt_id(filename):
    match = re.search(r"(RPT\d+[A-Z]{2})", filename.upper())
    return match.group(1) if match else None
 
def extract_tc(filename):

    match = re.search(r"TC(\d+)", filename.upper())

    return match.group(1) if match else None
 