from cwarning import CWarning

def analyze_pvs_report_line(line):
    columns = line.split(",")

    try:
        id = columns[3].replace('"', "").replace(")", "")
    except Exception:
        return None
    place = CWarning.get_filename(columns[8]) + ":" + columns[7].replace('"', "")
    if columns[4].find(":") > -1:
        desc = columns[4][1:columns[4].index(":")]
    else:
        desc = columns[4][1:].replace('"', "")
    try:
        source = columns[4][columns[4].index(":")+1:len(columns[4])-1]
    except Exception:
        source = ""
    warn = CWarning(id, desc, source, place)
    return warn