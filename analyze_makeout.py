import xlsxwriter
import sys
from cwarning import CWarning
from cwstatistics import CWStatistics
from worksheet_ext import WorksheetExt
from pvs_analyze import analyze_pvs_report_line

if __name__ == "__main__":
    pvs_report = None
    if len(sys.argv) < 2:
        print("pass path to make output file as argument")
        sys.exit(1)
    if len(sys.argv) > 2:
        pvs_report = open(sys.argv[2], "r")

    statistics = CWStatistics()

    workbook = xlsxwriter.Workbook('Report.xlsx')

    hcell_format = workbook.add_format()

    hcell_format.set_pattern(1)  # This is optional when using a solid fill.
    hcell_format.set_bg_color('#6E6E6E')
    hcell_format.set_bold()
    hcell_format.set_font_color("#FAFAFA")

    shcell_format = workbook.add_format()

    shcell_format.set_pattern(1)  # This is optional when using a solid fill.
    shcell_format.set_bg_color('#81F7D8')
    shcell_format.set_bold()

    # Set the columns widths.
    extwsheet_dict = dict()

    with open(sys.argv[1], "r") as makeout:
        row = 0
        lines = makeout.readlines()
        for i in range(0, len(lines)):
            if "warning: " in lines[i]:
                line = lines[i]
                place = CWarning.get_place(line)

                start_ind = 0
                if place.find("/") >= 0:
                    start_ind = place.rindex("/") + 1
                place_with_line_column = place[start_ind:]
                place = place[start_ind:place.index(":")]

                if place not in extwsheet_dict:
                    sheet = workbook.add_worksheet(place)
                    extwsheet_dict[place] = WorksheetExt(sheet)
                    extwsheet_dict[place].worksheet.set_column('A:G', 50)
                    extwsheet_dict[place].add_header_row(["Id", "Desc", "Place", "Source"], hcell_format)
                    extwsheet_dict[place].add_subheader("Compiler warnings:", shcell_format)

                id = CWarning.get_id(line)

                if CWarning.is_unique(place_with_line_column + lines[i+1]):
                    warning = CWarning(id, CWarning.get_desc(line), lines[i+1].replace("\n", ""), place_with_line_column)
                    warning.write_to_book(extwsheet_dict[place].worksheet, extwsheet_dict[place].get_row())
                    extwsheet_dict[place].row_inc()

                    statistics.add_warning(warning)

    if pvs_report is not None:
        pvs_lines = pvs_report.readlines()
        for line in pvs_lines[2:]:
            warn = analyze_pvs_report_line(line)
            if warn is None:
                continue
            filename = CWarning.get_filename(warn.place)
            if filename not in extwsheet_dict:
                sheet = workbook.add_worksheet(filename)
                extwsheet_dict[filename] = WorksheetExt(sheet)
                extwsheet_dict[filename].worksheet.set_column('A:G', 50)
                extwsheet_dict[filename].add_header_row(["Id", "Desc", "Place", "Source"], hcell_format)

            if not extwsheet_dict[filename].is_pvs:
                extwsheet_dict[filename].add_pvs(shcell_format)

            warn.write_to_book(extwsheet_dict[filename].worksheet, extwsheet_dict[filename].get_row())
            statistics.add_warning(warn)
            extwsheet_dict[filename].row_inc()
        pvs_report.close()

    # generate separate list for statistics
    statistics_wb = workbook.add_worksheet("Total")

    row = 1
    statistics_wb.set_column('A:A', 50)
    statistics_wb.set_column('B:B', 80)

    statistics_wb.write(0, 0, "ID", hcell_format)
    statistics_wb.write(0, 1, "Repeats", hcell_format)
    for key in statistics.by_id.keys():
        statistics_wb.write(row, 0, key)
        statistics_wb.write(row, 1, statistics.by_id[key])
        row += 1

    statistics_wb.write(row, 0, "FILE", hcell_format)
    statistics_wb.write(row, 1, "Warning", hcell_format)
    statistics_wb.write(row, 2, "Repeats", hcell_format)
    row += 1

    format1 = workbook.add_format()
    format2 = workbook.add_format()

    format1.set_bg_color('#F2F2F2')
    format1.set_bold()
    format1.set_border(1)
    format1.set_border_color("#000000")

    format2.set_bg_color('#D8D8D8')
    format2.set_bold()
    format2.set_border(1)
    format2.set_border_color("#000000")

    current_format = format1

    for key in statistics.by_file.keys():
        statistics_wb.write(row, 0, key, current_format)
        for id in statistics.by_file[key].keys():
            statistics_wb.write(row, 1, id, current_format)
            statistics_wb.write(row, 2, statistics.by_file[key][id], current_format)
            row += 1
            statistics_wb.write(row, 0, "", current_format)
        if current_format == format1:
            current_format = format2
        else:
            current_format = format1

    workbook.close()

