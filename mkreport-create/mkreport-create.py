#!/usr/bin/python3
import xlsxwriter
import argparse
from cwstatistics import CWStatistics
from pvs_analyze import analyze_pvs_report
from make_analyze import makefile_analyze
from summary_creator import create_summary

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    required = parser.add_argument_group('required arguments')
    required.add_argument("-makeout", help="GNU make output",
                        type=str, required=True)
    parser.add_argument("--pvs", help="PVS csv report",
                        type=str, default=None)
    parser.add_argument("--out", help="Output file",
                        type=str, default='report.xlsx', required=False)
    args = parser.parse_args()
    pvs_report = None

    if args.pvs is not None:
        pvs_report = open(args.pvs, "r")

    statistics = CWStatistics()

    workbook = xlsxwriter.Workbook(args.out)
    wsheet_summary_diff = workbook.add_worksheet("Summary")

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

    makefile_analyze(args.makeout, extwsheet_dict, workbook, statistics, hcell_format, shcell_format, only_stat=False)

    if pvs_report is not None:
        analyze_pvs_report(workbook, pvs_report, hcell_format,
                           shcell_format, extwsheet_dict, statistics, only_stat=False)
        pvs_report.close()

    create_summary(workbook, wsheet_summary_diff, hcell_format, statistics)
    workbook.close()
