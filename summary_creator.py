

def create_summary(workbook, hcell_format, statistics):
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
