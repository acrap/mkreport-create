import xlsxwriter
from cwarning import CWarning


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
        else:
            self.by_file[filename][cwarning.id] += 1
