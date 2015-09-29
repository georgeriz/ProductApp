from xlrd import open_workbook

class Product:
    def __init__(self, description_list):
        self.code = description_list[0]
        self.name = description_list[2]
        self.price_no_vat = description_list[21]
        self.price = description_list[22]
        self.cost = description_list[24]
        self.gain = description_list[25]
        self.gain_percentage = description_list[26]

        self._list = description_list

    def __repr__(self):
        return self.get_code() + ' ' + self.get_name()

    def get_code(self):
        return self.code
    def get_name(self):
        return self.name
    def get_price(self):
        if self.price == "":
            return "0.00"
        return '{0:,.2f}'.format(self.price)
    def get_price_no_vat(self):
        if self.price_no_vat == "":
            return "0.00"
        return '{0:,.2f}'.format(self.price_no_vat)
    def get_cost(self):
        if self.cost == "":
            return "0.00"
        return '{0:,.2f}'.format(self.cost)
    def get_gain(self):
        if self.gain == "":
            return "0.00"
        return '{0:,.2f}'.format(self.gain)
    def get_gain_percentage(self):
        if self.gain_percentage == "":
            return "-"
        return '{0:.0%}'.format(self.gain_percentage)


class ExcelReader:
    def __init__(self, filename, file_info):
        self.wb = open_workbook(filename)
        self.sheet = file_info[0]
        self.row = int(file_info[1])
        self.col = int(file_info[2])

    def run(self):
        my_list = []
        s = self.wb.sheet_by_name(self.sheet)
        for row in range(self.row, s.nrows):
            values = []
            for col in range(self.col - 1, s.ncols):
                values.append(s.cell(row,col).value)
            product = Product(values)
            my_list.append(product)
        return my_list