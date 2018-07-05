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
        try:
            start_ind = line.index("[")
            end_ind = line.index("]") + 1
            return line[start_ind:end_ind]
        except Exception:
            return None

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

