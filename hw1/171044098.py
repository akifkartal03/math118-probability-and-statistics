import openpyxl
import statistics
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, NamedStyle


# I didn't want to bother with global variables therefore,
# I created a class to encapsulate whole homework.
class HW1:

    def __init__(self):
        # write class variables here.
        self.__data = None
        self.__countries = []
        self.__summary = []
        self.__result = [[]]
        self.__index = 1

    def get_results(self):
        # this is only public method to calculate all results
        sheet = openpyxl.load_workbook("owid-covid-data.xlsx", data_only=True)
        self.__data = sheet.active
        self.__countries = self.__get_list(self.__data["C"])
        self.__summary.append(["Total Country"])
        self.__summary.append([])
        self.__summary.append(["Earliest Date", "Country"])
        self.__q1()
        self.__q2()
        self.__q18()  # first create header for summary
        self.__q3()
        self.__q4()
        self.__q5()
        self.__q6()
        self.__q7()
        self.__q8()
        self.__q9()
        self.__q10()
        self.__q11()
        self.__q12()
        self.__q13()
        self.__q14()
        self.__q15()
        self.__q16()
        self.__q17()
        self.__write_summary()

    def __q1(self):
        first_column = self.__data["A"]
        unique_countries = self.__get_set(first_column)
        result = [len(unique_countries)]
        self.__summary[1] = result

    def __q2(self):
        dates = self.__get_list(self.__data["D"])
        # get rid of repeated elements
        # unique_dates = self.__get_set(dates)
        # sorted_dates = sorted(unique_dates, key=lambda x: tuple(map(int, x.split('-'))))
        min_date = min(dates)
        indexes = [i for i, date in enumerate(dates) if date == min_date]
        for index in indexes:
            res = [min_date, self.__countries[index]]
            self.__summary.append(res)
            self.__index = self.__index + 1

    def __q3(self):
        self.__common_3_4_11("E", 3)

    def __q4(self):
        self.__common_3_4_11("H", 4)

    def __q5(self):
        self.__common_5to13("Q")

    def __q6(self):
        self.__common_5to13("R")

    def __q7(self):
        self.__common_5to13("T")

    def __q8(self):
        self.__common_5to13("V")

    def __q9(self):
        self.__common_5to13("X")

    def __q10(self):
        self.__common_5to13("Z")

    def __q11(self):
        self.__common_3_4_11("AA", 11)

    def __q12(self):
        self.__common_5to13("AF")

    def __q13(self):
        self.__common_5to13("AG")

    def __q14(self):
        self.__common_3_4_11("AJ", 14)

    def __q15(self):
        self.__common_3_4_11("AK", 15)

    def __q16(self):
        self.__common_3_4_11("AI", 16)

    def __q17(self):
        self.__common_3_4_11("AS", 17)
        self.__common_3_4_11("AU", 17)
        self.__common_3_4_11("AV", 17)
        self.__common_3_4_11("AW", 17)
        self.__common_3_4_11("AX", 17)
        self.__common_3_4_11("AZ", 17)
        self.__common_3_4_11("BA", 17)
        self.__common_3_4_11("BB", 17)
        self.__common_3_4_11("BC", 17)
        self.__common_3_4_11("BD", 17)
        self.__common_3_4_11("BE", 17)
        self.__common_3_4_11("BF", 17)
        self.__common_3_4_11("BG", 17)

    def __q18(self):
        header = ["Country", "q#3", "q#4", "q#5_min", "q#5_max", "q#5_avg", "q#5_var",
                  "q#6_min", "q#6_max", "q#6_avg", "q#6_var", "q#7_min", "q#7_max", "q#7_avg", "q#7_var",
                  "q#8_min", "q#8_max", "q#8_avg", "q#8_var", "q#9_min", "q#9_max", "q#9_avg", "q#9_var",
                  "q#10_min", "q#10_max", "q#10_avg", "q#10_var", "q#11", "q#12_min", "q#12_max", "q#12_avg",
                  "q#12_var", "q#13_min", "q#13_max", "q#13_avg", "q#13_var", "q#14", "q#15", "q#16", "population",
                  "median age", "# of people aged 65 older", "# of people aged 70 older", "economic performance",
                  "death rates due to heart disease", "diabetes prevalence", "# of female smokers",
                  "# of male smokers", "handwashing facilities", "hospital beds per thousand people",
                  "life expectancy", "human development index"]
        self.__summary.append(header)

    def __get_set(self, data):
        my_set = set()
        for element in data[1:]:
            my_set.add(element.value)
        return my_set

    def __get_list(self, data):
        my_list = []
        for element in data[1:]:
            my_list.append(element.value)
        return my_list

    def __common_5to13(self, column):
        rate = self.__get_list(self.__data[column])
        my_set = set()
        rate_list = []
        i = 0
        j = 0
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country)
            if len(my_set) == size or i == 0:
                if rate[i] is not None:
                    rate_list.append(rate[i])
            else:
                # ct_name = self.__countries[i - 1]
                avg = None
                minimum = None
                maximum = None
                variation = None
                if len(rate_list) >= 1:
                    avg = round(sum(rate_list) / len(rate_list), 2)
                    minimum = min(rate_list)
                    maximum = max(rate_list)
                if len(rate_list) >= 2:
                    variation = round(statistics.variance(rate_list), 2)
                if rate[i] is not None:
                    rate_list = [rate[i]]
                else:
                    rate_list = []
                self.__result[j].append(minimum)
                self.__result[j].append(maximum)
                self.__result[j].append(avg)
                self.__result[j].append(variation)
                j = j + 1
            i = i + 1

    def __common_3_4_11(self, column, question):
        total = self.__get_list(self.__data[column])
        my_set = set()
        indexes = []
        temp = [None]
        i = 1
        k = 0
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country)
            if len(my_set) != size:
                indexes.append(i - 1)
                temp.append(None)
                k = k + 1
            if total[i - 1] is not None:
                temp[k] = total[i - 1]
            i = i + 1
        i = 0
        for index in indexes[1:]:
            if total[index - 1] is not None:
                res = total[index - 1]
            else:
                res = temp[i+1]
            if question == 3:
                if i == 0:
                    self.__result[i] = [self.__countries[index - 1], res]
                else:
                    self.__result.append([self.__countries[index - 1], res])
            else:
                self.__result[i].append(res)

            i = i + 1

    def __write_summary(self):
        o_wb = Workbook()
        out_filename = 'output.csv'
        o_sheet = o_wb.active
        for row in self.__summary:
            o_sheet.append(row)
        header_style = NamedStyle(name="header_style")
        header_style.font = Font(bold=True)
        header_row = o_sheet[1]
        for cell in header_row:
            cell.style = header_style
        header_row = o_sheet[3]
        for cell in header_row:
            cell.style = header_style
        header_row = o_sheet[3 + self.__index]
        for cell in header_row:
            cell.style = header_style
        # write results
        for row in self.__result:
            o_sheet.append(row)
        o_wb.save(filename=out_filename)
        print("Output file created!")


a = HW1()
a.get_results()
