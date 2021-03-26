import random
import xlwt
import re
import os


class Report:

    def __init__(self, numofitems, numofdims):
        self.numofdims = numofdims
        self.numofitems = numofitems
        self.dimcounter = self.numofdims - (self.numofdims - 1)

    def get_req_dim(self):
        """
        Method to ask for the required dimension and the tolerances for it.
        "Input : Req dimension with tolerances"
        "Output : Req Dimension with tolerances in audit table style"
        """
        while True:
            try:
                allowed = "^[0-9]$|^--$|^[0-9]+\.?[0-9]+$"  # Pattern for tolerances (Integer,--,or float number)
                self.dimrequired = input("Dimension Number {} Required: ".format(self.dimcounter))
                self.plustolerance = input("Enter + Tolerance: ")
                self.minustolerance = input("Enter - Tolerance: ")
                if re.search(allowed, self.plustolerance) \
                        and re.search(allowed, self.minustolerance):
                    return self.get_dim_with_tols()
                else:
                    print("Only integers,floats or -- are only allowed in tolerances,try again.")
                    continue
            except ValueError:
                print("Exception point in get_req_dim method")
                print("Error: Tolerances must be only integers/floats or type "'"--"'" if None")
                continue

    def get_dim_with_tols(self):
        """
        Method to return the required dimension with the required tolerances
        in audit table style
        """

        self.dimrequired = self.dimrequired.replace('di', 'Ø').replace("DI", 'Ø').replace('de', '°') \
            .replace("DE", '°').replace("r", "R").replace("x", "X").replace("+-", "±").replace("-+", "±")

        if self.check_tols() is None:
            return self.dimrequired
        else:
            return self.dimrequired + self.check_tols()

    def check_tols(self):
        """
        Method to check and return the tolerances with the suitable symbols
        made this method so we don't have to write the same code more than once
        if we want to check the tolerances if they're the same or none or different.
        returns the required tolerances with the suitable symbols
        """
        try:
            if str(self.plustolerance) == "--" and str(self.minustolerance) == "--":
                return None

            elif self.dimrequired.find("°") != -1 and str(self.plustolerance) == "--" \
                    and str(self.minustolerance) != "--":
                return " -" + self.minustolerance + "°"

            elif self.dimrequired.find("°") != -1 and str(self.plustolerance) != "--" \
                    and str(self.minustolerance) == "--":
                return " +" + self.plustolerance + "°"

            elif str(self.plustolerance) == "--":
                return " -" + self.minustolerance

            elif str(self.minustolerance) == "--":
                return " +" + self.plustolerance

            elif self.dimrequired.find("°") != -1 and self.plustolerance == self.minustolerance:
                return " ±" + self.plustolerance + "°"

            elif self.dimrequired.find("°") != -1 and self.plustolerance != self.minustolerance \
                    and str(self.plustolerance) != "--" and str(self.minustolerance) != "--":
                return " +" + self.plustolerance + "°" + " -" + self.minustolerance + "°"

            elif float(self.plustolerance) != float(self.minustolerance):
                return " +" + str(self.plustolerance) + " -" + str(self.minustolerance)

            elif float(self.plustolerance) > 0.0 < float(self.minustolerance) \
                    and float(self.plustolerance) == float(self.minustolerance) \
                    or float(self.plustolerance) == 0 and float(self.minustolerance) == 0:
                return " ±" + str(self.plustolerance)
            else:
                return "This is in else in check_tols method/Check why"
        except ValueError:
            raise ValueError

    def min_and_max_dims(self, reqdim):
        """
        Input: Minimum dimension and maximum dimension for the required dimension
        Output: Random float number between the minimum and the maximum received ( 2 places after float point )
        """
        self.reqdim = reqdim
        self.mindiminput = input(
            "Enter the minimum dimension for Dim Number {} (".format(self.dimcounter) + self.reqdim + "): ")
        self.maxdiminput = input(
            "Enter the maximum dimension for DIm Number {} (".format(self.dimcounter) + self.reqdim + "): ")
        self.dimcounter += 1  # Increase the current dimension required number.

    def random_num_gen(self):
        """
        Method to generate Random number between the given minimum input and the minimum output
        if the min/max dimension is radius and not equal to it returns randomly one of them
        either the radius written in minimum dimension or the radius written in maximum dimension
        """
        for i in range(self.numofitems):
            try:
                # check if tolerances are radius to choose randomly between the 2.
                if self.mindiminput.find('R') != -1 or self.mindiminput.find("r") != -1 \
                        and self.maxdiminput.find('R') != -1 or self.maxdiminput.find("r") != -1:
                    yield random.choice((self.mindiminput, self.maxdiminput))

                elif self.mindiminput == 'OK' or self.mindiminput == 'ok' \
                        and self.maxdiminput == 'OK' or self.maxdiminput == 'ok':
                    yield 'OK'
                else:
                    yield round(random.uniform(float(self.mindiminput), float(self.maxdiminput)), 2)
            except ValueError:
                yield self.dimrequired


def sampling_qty(serving_qty):
    if 2 <= serving_qty <= 5:
        return serving_qty
    elif 6 <= serving_qty <= 8:
        return 5
    elif 9 <= serving_qty <= 15:
        return 5
    elif 16 <= serving_qty <= 25:
        return 5
    elif 26 <= serving_qty <= 50:
        return 5
    elif 51 <= serving_qty <= 90:
        return 7
    elif 91 <= serving_qty <= 150:
        return 11
    elif 151 <= serving_qty <= 280:
        return 13
    elif 281 <= serving_qty <= 500:
        return 16
    elif 501 <= serving_qty <= 1200:
        return 19
    elif 1201 <= serving_qty <= 3200:
        return 23
    elif 3201 <= serving_qty <= 10000:
        return 29
    elif 10001 <= serving_qty <= 35000:
        return 35
    elif 35001 <= serving_qty <= 150000:
        return 40
    elif 150001 <= serving_qty <= 500000:
        return 40
    elif serving_qty >= 500001:
        return 40
    else:
        return 1


def main():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("GeneratedDims")
    style = xlwt.easyxf('font: name Arial, bold off, height 240; align: horz center;borders: top medium')
    drawingnumber = input("Enter Drawing number: ")
    numofitems = sampling_qty(int(input("Enter serving quantity: ")))
    numofdimension = int(input("Enter number of dimensions (Balloons): "))
    report = Report(numofitems, numofdimension)
    rows = 1
    column = 0
    firstrowlist = ["Dim Required",
                    "Dim with Tols",
                    "Plus Tol",
                    "Minus Tol",
                    "Minimum",
                    "Maximum"]

    # for loop to write the first row (Titles) that is in firstrowlist
    for i in range(0, len(firstrowlist)):
        sheet.write(0, i, firstrowlist[i].capitalize(), style)

    for i in range(numofitems):
        sheet.write(0, len(firstrowlist) + i, "Item " + str(i + 1), style)

    for i in range(numofdimension):
        report.min_and_max_dims(report.get_req_dim())
        sheet.write(rows, column, report.dimrequired, style)
        column += 1
        sheet.write(rows, column, report.get_dim_with_tols(), style)
        column += 1
        sheet.write(rows, column, report.plustolerance, style)
        column += 1
        sheet.write(rows, column, report.minustolerance, style)
        column += 1
        sheet.write(rows, column, report.mindiminput.upper(), style)
        column += 1
        sheet.write(rows, column, report.maxdiminput.upper(), style)
        column += 1
        for item in report.random_num_gen():
            sheet.write(rows, column, item, style)
            column += 1
        column = 0
        rows += 1
    workbook.save(str(drawingnumber) + ".xls".capitalize())
    file = "C:\\Git\\QAReports\\" + str(drawingnumber) + ".xls"
    os.startfile(file)


main()
