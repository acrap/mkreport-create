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
        return place[start_ind:place.index(":")]


class CWStatistics:
    def __init__(self):
        self.by_id = dict()
        self.by_file = dict()

    def add_warning(self, cwarning):
        if cwarning.id not in self.by_id:
            self.by_id[cwarning.id] = 1
        else:
            self.by_id[cwarning.id] += 1

        filename = CWarning.get_filename(cwarning.place)
        if filename not in self.by_file:
            self.by_file[filename] = 1
        else:
            self.by_file[filename] += 1


class WorksheetExt:
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.row = 0

    def row_inc(self):
        self.row += 1

    def get_row(self):
        return self.row


if __name__ == "__main__":

    if len(sys.argv)<2:
        print("pass path to make output file as argument")
        sys.exit(1)

    statistics = CWStatistics()
    workbook = xlsxwriter.Workbook('Report.xlsx')

    hcell_format = workbook.add_format()

    hcell_format.set_pattern(1)  # This is optional when using a solid fill.
    hcell_format.set_bg_color('#CEF6CE')

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
                    row = 1
                    sheet = workbook.add_worksheet(place)
                    files_dict[place] = WorksheetExt(sheet)
                    files_dict[place].worksheet.set_column('A:G', 50)
                    files_dict[place].worksheet.write(0, 0, "id", hcell_format)
                    files_dict[place].worksheet.write(0, 1, "desc", hcell_format)
                    files_dict[place].worksheet.write(0, 2, "place", hcell_format)
                    files_dict[place].worksheet.write(0, 3, "source", hcell_format)
                    files_dict[place].row_inc()

                id = CWarning.get_id(line)
                if len(id) == 0:
                    pass

                if CWarning.is_unique(place_with_line_column + lines[i+1]):
                    warning = CWarning(id, CWarning.get_desc(line), lines[i+1].replace("\n", ""), place_with_line_column)
                    warning.write_to_book(files_dict[place].worksheet, files_dict[place].get_row())
                    files_dict[place].row_inc()

                    statistics.add_warning(warning)

    # generate separate list for statistics
    statistics_wb = workbook.add_worksheet("Total")

    row = 1
    statistics_wb.set_column('A:C', 50)
    statistics_wb.write(0, 0, "ID", hcell_format)
    statistics_wb.write(0, 1, "Repeats", hcell_format)
    for key in statistics.by_id.keys():
        statistics_wb.write(row, 0, key)
        statistics_wb.write(row, 1, statistics.by_id[key])
        row += 1
    workbook.close()







