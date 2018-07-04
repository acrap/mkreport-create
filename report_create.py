import xlsxwriter
import sys
from cwstatistics import CWStatistics
from pvs_analyze import analyze_pvs_report
from make_analyze import makefile_analyze


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

    makefile_analyze(sys.argv[1], extwsheet_dict, workbook, statistics, hcell_format, shcell_format)

    if pvs_report is not None:
        analyze_pvs_report(workbook, pvs_report, hcell_format,
                           shcell_format, extwsheet_dict, statistics, only_stat=False)
        pvs_report.close()

    # generate separate list for statistics
    statistics_wb = workbook.add_worksheet("Summary")

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
    format1.set_border(1)
    format1.set_border_color("#000000")

    format2.set_bg_color('#D8D8D8')
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

