import xlsxwriter
import sys
from cwstatistics import CWStatistics
from pvs_analyze import analyze_pvs_report
from make_analyze import makefile_analyze
from summary_creator import create_summary, create_summary_diff

if __name__ == "__main__":
    pvs_report1 = open("pvs_report.csv")
    pvs_report2 = open("pvs_report_new.csv")

    statistics = CWStatistics()
    statistics2 = CWStatistics()

    workbook = xlsxwriter.Workbook('SummaryDiff.xlsx')

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

    makefile_analyze("warnings", extwsheet_dict, workbook, statistics, hcell_format, shcell_format, only_stat=True)
    makefile_analyze("warnings_new", extwsheet_dict, workbook, statistics2, hcell_format, shcell_format, only_stat=True)

    analyze_pvs_report(workbook, pvs_report1, hcell_format,
                       shcell_format, extwsheet_dict, statistics, only_stat=True)
    pvs_report1.close()

    analyze_pvs_report(workbook, pvs_report2, hcell_format,
                       shcell_format, extwsheet_dict, statistics2, only_stat=True)
    pvs_report2.close()

    create_summary_diff(workbook, hcell_format, statistics, statistics2)
    workbook.close()

