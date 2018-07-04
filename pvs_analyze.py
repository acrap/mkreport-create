from cwarning import CWarning
from worksheet_ext import WorksheetExt


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


def analyze_pvs_report(wbook, report_file, header_format, subheader_format, worksheet_dict, stats, only_stat):
    pvs_lines = report_file.readlines()
    for line in pvs_lines[2:]:
        warn = analyze_pvs_report_line(line)
        if warn is None:
            continue
        filename = CWarning.get_filename(warn.place)
        if not only_stat:
            if filename not in worksheet_dict:
                sheet = wbook.add_worksheet(filename)
                worksheet_dict[filename] = WorksheetExt(sheet)
                worksheet_dict[filename].worksheet.set_column('A:G', 50)
                worksheet_dict[filename].add_header_row(["Id", "Desc", "Place", "Source"], header_format)

            if not worksheet_dict[filename].is_pvs:
                worksheet_dict[filename].add_pvs(subheader_format)

            warn.write_to_book(worksheet_dict[filename].worksheet, worksheet_dict[filename].get_row())
        stats.add_warning(warn)
        if not only_stat:
            worksheet_dict[filename].row_inc()
