import xlsxwriter
import sys
from cwstatistics import CWStatistics
from pvs_analyze import analyze_pvs_report
from make_analyze import makefile_analyze
from summary_creator import create_summary

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

    create_summary(workbook, hcell_format, statistics)
    workbook.close()

