#!/usr/bin/python3
import xlsxwriter
import argparse
from mkreport_lib.cwstatistics import CWStatistics
from mkreport_lib.pvs_analyze import analyze_pvs_report
from mkreport_lib.make_analyze import makefile_analyze
from mkreport_lib.summary_creator import create_summary_diff
from mkreport_lib.cwarning import cwarning_uniq

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    required = parser.add_argument_group('required arguments')
    required.add_argument("-makeout1", help="GNU make output for the previous state",
                        type=str, required=True)
    required.add_argument("-makeout2", help="GNU make output for the current state",
                        type=str, required=True)

    parser.add_argument("-pvs1", help="PVS csv report for the previous state",
                        type=str)

    parser.add_argument("-pvs2", help="PVS csv report for the current state",
                        type=str)

    parser.add_argument("--out", help="Output file",
                        type=str, default='SummaryDiff.xlsx', required=False)

    args = parser.parse_args()

    pvs_report1 = open(args.pvs1)
    pvs_report2 = open(args.pvs2)

    statistics = CWStatistics()
    statistics2 = CWStatistics()

    workbook = xlsxwriter.Workbook(args.out)

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

    makefile_analyze(args.makeout1, extwsheet_dict, workbook, statistics, hcell_format, shcell_format, only_stat=True)

    cwarning_uniq.clear()

    makefile_analyze(args.makeout2, extwsheet_dict, workbook, statistics2, hcell_format, shcell_format, only_stat=True)

    analyze_pvs_report(workbook, pvs_report1, hcell_format,
                       shcell_format, extwsheet_dict, statistics, only_stat=True)
    pvs_report1.close()

    analyze_pvs_report(workbook, pvs_report2, hcell_format,
                       shcell_format, extwsheet_dict, statistics2, only_stat=True)
    pvs_report2.close()

    wsheet_summary_diff = workbook.add_worksheet("Summary")
    create_summary_diff(workbook, wsheet_summary_diff, hcell_format, statistics, statistics2)
    workbook.close()

