

def create_summary(workbook, wsheet, hcell_format, statistics):
    # generate separate list for statistics
    statistics_wsheet = wsheet
    row = 1
    statistics_wsheet.set_column('A:A', 50)
    statistics_wsheet.set_column('B:B', 80)

    statistics_wsheet.write(0, 0, "ID", hcell_format)
    statistics_wsheet.write(0, 1, "Repeats", hcell_format)
    for key in statistics.by_id.keys():
        statistics_wsheet.write(row, 0, key)
        statistics_wsheet.write(row, 1, statistics.by_id[key])
        row += 1

    statistics_wsheet.write(row, 0, "FILE", hcell_format)
    statistics_wsheet.write(row, 1, "Warning", hcell_format)
    statistics_wsheet.write(row, 2, "Repeats", hcell_format)
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
        statistics_wsheet.write(row, 0, key, current_format)
        for id in statistics.by_file[key].keys():
            statistics_wsheet.write(row, 1, id, current_format)
            statistics_wsheet.write(row, 2, statistics.by_file[key][id], current_format)
            row += 1
            statistics_wsheet.write(row, 0, "", current_format)
        if current_format == format1:
            current_format = format2
        else:
            current_format = format1


def create_summary_diff(workbook, wsheet, hcell_format, statistics, statistics2):
    # generate separate list for statistics
    format1 = workbook.add_format()
    format2 = workbook.add_format()
    format_plus = workbook.add_format()
    format_minus = workbook.add_format()

    format1.set_bg_color('#F2F2F2')
    format1.set_border(1)
    format1.set_border_color("#000000")

    format2.set_bg_color('#D8D8D8')
    format2.set_border(1)
    format2.set_border_color("#000000")

    format_minus.set_bg_color('#2EFE9A')
    format_plus.set_bg_color('#DF3A01')

    statistics_wb = wsheet

    row = 1
    statistics_wb.set_column('A:A', 50)
    statistics_wb.set_column('B:B', 80)

    statistics_wb.write(0, 0, "ID", hcell_format)
    statistics_wb.write(0, 1, "Repeats", hcell_format)
    statistics_wb.write(0, 2, "Current", hcell_format)
    statistics_wb.write(0, 3, "Progress", hcell_format)
    all_keys = list(statistics.by_id.keys())
    for item in statistics2.by_id.keys():
        if item not in all_keys:
            all_keys.append(item)

    for key in all_keys:
        val1 = 0
        val2 = 0
        if key in statistics.by_id:
            val1 = statistics.by_id[key]
        statistics_wb.write(row, 0, key)
        statistics_wb.write(row, 1, val1)
        if key in statistics2.by_id:
            val2 = statistics2.by_id[key]
        statistics_wb.write(row, 2, val2)
        if val1 > val2:
            statistics_wb.write(row, 3, "-{diff}".format(diff=(val1-val2)), format_minus)
        if val1 < val2:
            statistics_wb.write(row, 3, "+{diff}".format(diff=(val2-val1)), format_plus)
        row += 1

    statistics_wb.write(row, 0, "FILE", hcell_format)
    statistics_wb.write(row, 1, "Warning", hcell_format)
    statistics_wb.write(row, 2, "Repeats", hcell_format)
    statistics_wb.write(row, 3, "Current", hcell_format)
    statistics_wb.write(row, 4, "Progress", hcell_format)
    row += 1

    current_format = format1

    all_keys = list(statistics.by_file.keys())
    for item in statistics2.by_file.keys():
        if item not in all_keys:
            all_keys.append(item)

    for key in all_keys:
        statistics_wb.write(row, 0, key, current_format)
        if key not in statistics.by_file:
            statistics.by_file[key] = dict()

        all_id = list(statistics.by_file[key])
        if key in statistics2.by_file:
            for item in statistics2.by_file[key]:
                if item not in all_id:
                    all_id.append(item)
        for id in all_id:
            val1 = 0
            val2 = 0
            statistics_wb.write(row, 1, id, current_format)
            if key in statistics.by_file:
                if id in statistics.by_file[key]:
                    val1 = statistics.by_file[key][id]
            statistics_wb.write(row, 2, val1, current_format)
            if key in statistics2.by_file:
                if id in statistics2.by_file[key]:
                    val2 = statistics2.by_file[key][id]
            statistics_wb.write(row, 3, val2, current_format)
            if val1 > val2:
                statistics_wb.write(row, 4, "-{diff}".format(diff=(val1-val2)), format_minus)
            if val1 < val2:
                statistics_wb.write(row, 4, "+{diff}".format(diff=(val2-val1)), format_plus)
            row += 1
            statistics_wb.write(row, 0, "", current_format)
        if current_format == format1:
            current_format = format2
        else:
            current_format = format1