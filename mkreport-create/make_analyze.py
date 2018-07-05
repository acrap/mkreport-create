from cwarning import CWarning
from worksheet_ext import WorksheetExt


def makefile_analyze(filename, extwsheet_dict, workbook, statistics, hformat, shformat, only_stat):
    with open(filename, "r") as makeout:
        lines = makeout.readlines()
        for i in range(0, len(lines)):
            if "warning: " in lines[i]:
                line = lines[i]
                try:
                    place = CWarning.get_place(line)

                    start_ind = 0
                    if place.find("/") >= 0:
                        start_ind = place.rindex("/") + 1
                    place_with_line_column = place[start_ind:]
                    place = place[start_ind:place.index(":")]
                    if not only_stat:
                        if place not in extwsheet_dict:
                            sheet = workbook.add_worksheet(place)
                            extwsheet_dict[place] = WorksheetExt(sheet)
                            extwsheet_dict[place].worksheet.set_column('A:G', 50)
                            extwsheet_dict[place].worksheet.write(extwsheet_dict[place].get_row(), 0,
                                                                  'internal:Summary')
                            extwsheet_dict[place].row_inc()
                            extwsheet_dict[place].add_header_row(["Id", "Desc", "Place", "Source"], hformat)

                            extwsheet_dict[place].add_subheader("Compiler warnings:", shformat)

                    id = CWarning.get_id(line)
                except Exception:
                    continue

                if id is None:
                    continue

                if CWarning.is_unique(place_with_line_column + lines[i+1]):
                    warning = CWarning(id, CWarning.get_desc(line), lines[i+1].replace("\n", ""), place_with_line_column)
                    if not only_stat:
                        warning.write_to_book(extwsheet_dict[place].worksheet, extwsheet_dict[place].get_row())
                        extwsheet_dict[place].row_inc()

                    statistics.add_warning(warning)
