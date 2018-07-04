import xlsxwriter
import sys

cwarning_uniq = dict()


class CWarning:
    def __init__(self, id, desc, source, place):
        self.id = id
        self.desc = desc
        self.source = source
        self.place = place

    def write_to_book(self, worksheet, row):
        worksheet.write(row, 0, self.id)
        worksheet.write(row, 1, self.desc)
        worksheet.write(row, 2, self.place)
        worksheet.write(row, 3, self.source)

    @staticmethod
    def is_unique(place):
        global cwarning_uniq

        if place not in cwarning_uniq:
            cwarning_uniq[place] = 1
            return True
        return False

    @staticmethod
    def get_id(line):
        start_ind = line.index("[")
        end_ind = line.index("]") + 1
        return line[start_ind:end_ind]

    @staticmethod
    def get_desc(line):
        start_ind = line.index("warning:")
        end_ind = line.index(" [") + 1
        return line[start_ind+8:end_ind]

    @staticmethod
    def get_place(line):
        end_ind = line.index(": ")
        return line[0:end_ind]

    @staticmethod
    def get_filename(place):
        start_ind = 0
        if place.find("/") >= 0:
            start_ind = place.rindex("/") + 1
        if place.find(":") == -1:
            res = place[start_ind:]
        else:
            res = place[start_ind:place.index(":")]
        res = res.replace("\n", "")
        return res


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


class CWStatistics:
    def __init__(self):
        self.by_id = dict()
        self.by_file = dict()

    def add_warning(self, cwarning):
        # add description to PVS warnings id's
        if cwarning.id.find("V") != -1:
            cwarning.id = cwarning.id + "(" + cwarning.desc + ")"

        if cwarning.id not in self.by_id:
            self.by_id[cwarning.id] = 1
        else:
            self.by_id[cwarning.id] += 1

        filename = CWarning.get_filename(cwarning.place)
        if filename not in self.by_file:
            self.by_file[filename] = dict()

        if cwarning.id not in self.by_file[filename]:
            self.by_file[filename][cwarning.id] = 1

        self.by_file[filename][cwarning.id] += 1


class WorksheetExt:
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.row = 0
        self.is_pvs = False

    def row_inc(self):
        self.row += 1

    def get_row(self):
        return self.row

    def is_pvs_added(self):
        return self.is_pvs

    def add_header_row(self, hlist, hformat):
        column = 0
        for item in hlist:
            self.worksheet.write(self.row, column, item, hformat)
            column += 1
        self.row_inc()

    def add_subheader(self, text, cformat):
        self.row_inc()
        self.worksheet.write(self.row, 0, text, cformat)
        for i in range(1, 4):
            self.worksheet.write(self.row, i, "", cformat)
        self.row_inc()

    def add_pvs(self, cformat):
        self.is_pvs = True
        self.add_subheader("PVS Report:", cformat)


if __name__ == "__main__":

    #test_str = ',error,"=HYPERLINK(""https://www.viva64.com/en/w/v114/"", ""V114"")","Dangerous explicit type pointer conversion: (int *) & val","=HYPERLINK(""file:///home/andrey/alexeyd_MW_2_internal_next/vobs/MW/DMAE/user_commands.c"", "" Open file"")","1829",/home/andrey/alexeyd_MW_2_internal_next/vobs/MW/DMAE/user_commands.c'
    #analyze_pvs_report_line(test_str)

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
    hcell_format.set_bg_color('#CEF6CE')
    hcell_format.set_bold()

    shcell_format = workbook.add_format()

    shcell_format.set_pattern(1)  # This is optional when using a solid fill.
    shcell_format.set_bg_color('#81F7D8')
    shcell_format.set_bold()

    # Set the columns widths.
    files_dict = dict()

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

                if place not in files_dict:
                    sheet = workbook.add_worksheet(place)
                    files_dict[place] = WorksheetExt(sheet)
                    files_dict[place].worksheet.set_column('A:G', 50)
                    files_dict[place].add_header_row(["Id", "Desc", "Place", "Source"], hcell_format)
                    files_dict[place].add_subheader("Compiler warnings:", shcell_format)

                id = CWarning.get_id(line)

                if CWarning.is_unique(place_with_line_column + lines[i+1]):
                    warning = CWarning(id, CWarning.get_desc(line), lines[i+1].replace("\n", ""), place_with_line_column)
                    warning.write_to_book(files_dict[place].worksheet, files_dict[place].get_row())
                    files_dict[place].row_inc()

                    statistics.add_warning(warning)

    if pvs_report is not None:
        pvs_lines = pvs_report.readlines()
        for line in pvs_lines[2:]:
            warn = analyze_pvs_report_line(line)
            if warn is None:
                continue
            filename = CWarning.get_filename(warn.place)
            if filename not in files_dict:
                sheet = workbook.add_worksheet(filename)
                files_dict[filename] = WorksheetExt(sheet)
                files_dict[filename].worksheet.set_column('A:G', 50)
                files_dict[filename].add_header_row(["Id", "Desc", "Place", "Source"], hcell_format)

            if not files_dict[filename].is_pvs:
                files_dict[filename].add_pvs(shcell_format)

            warn.write_to_book(files_dict[filename].worksheet, files_dict[filename].get_row())
            statistics.add_warning(warn)
            files_dict[filename].row_inc()
        pvs_report.close()

    # generate separate list for statistics
    statistics_wb = workbook.add_worksheet("Total")

    row = 1
    statistics_wb.set_column('A:A', 30)
    statistics_wb.set_column('B:C', 80)

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

    for key in statistics.by_file.keys():
        statistics_wb.write(row, 0, key)
        for id in statistics.by_file[key].keys():
            statistics_wb.write(row, 1, id)
            statistics_wb.write(row, 2, statistics.by_file[key][id])
            row += 1

    workbook.close()

