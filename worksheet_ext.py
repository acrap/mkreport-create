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
