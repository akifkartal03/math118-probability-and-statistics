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
        self.__output_sheet = None

    def get_results(self):
        # this is only public method to calculate all results
        sheet = openpyxl.load_workbook("owid-covid-data.xlsx", data_only=True)
        self.__data = sheet.active
        self.__countries = self.__get_list(self.__data["C"])
        self.__summary.append(["1) Total Country"])
        self.__summary.append([])
        self.__summary.append(["2) Earliest Date", "Country"])
        self.__q1()
        self.__q2()
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
                res = temp[i + 1]
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
        header_style = NamedStyle(name="header_style")
        header_style.font = Font(bold=True)
        self.__output_sheet = o_sheet

        self.__q1_write()
        self.__q2_write()
        header_row = o_sheet[1]
        for cell in header_row:
            cell.style = header_style
        header_row = o_sheet[3]
        for cell in header_row:
            cell.style = header_style

        self.__q3_write()
        header_row = o_sheet[3 + self.__index]
        for cell in header_row:
            cell.style = header_style

        self.__q4_write()
        header_row = o_sheet[4 + self.__index + len(self.__result)]
        for cell in header_row:
            cell.style = header_style

        self.__q5_write()
        header_row = o_sheet[5 + self.__index + (2 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q6_write()
        header_row = o_sheet[6 + self.__index + (3 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q7_write()
        header_row = o_sheet[7 + self.__index + (4 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q8_write()
        header_row = o_sheet[8 + self.__index + (5 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q9_write()
        header_row = o_sheet[9 + self.__index + (6 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q10_write()
        header_row = o_sheet[10 + self.__index + (7 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q11_write()
        header_row = o_sheet[11 + self.__index + (8 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q12_write()
        header_row = o_sheet[12 + self.__index + (9 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q13_write()
        header_row = o_sheet[13 + self.__index + (10 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q14_write()
        header_row = o_sheet[14 + self.__index + (11 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q15_write()
        header_row = o_sheet[15 + self.__index + (12 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q16_write()
        header_row = o_sheet[16 + self.__index + (13 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q17_write()
        header_row = o_sheet[17 + self.__index + (14 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        self.__q18_write()
        header_row = o_sheet[18 + self.__index + (15 * len(self.__result))]
        for cell in header_row:
            cell.style = header_style

        o_wb.save(filename=out_filename)
        print("Output file created!")

    def __q1_write(self):
        for row in self.__summary[0:2]:
            self.__output_sheet.append(row)

    def __q2_write(self):
        for row in self.__summary[2:]:
            self.__output_sheet.append(row)

    def __q3_write(self):
        self.__output_sheet.append(["3)Country", "q#3"])
        for row in self.__result:
            new_list = [row[0], row[1]]
            self.__output_sheet.append(new_list)

    def __q4_write(self):
        self.__output_sheet.append(["4)Country", "q#4"])
        for row in self.__result:
            new_list = [row[0], row[2]]
            self.__output_sheet.append(new_list)

    def __q5_write(self):
        self.__output_sheet.append(["5)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[3], row[4], row[5], row[6]]
            self.__output_sheet.append(new_list)

    def __q6_write(self):
        self.__output_sheet.append(["6)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[7], row[8], row[9], row[10]]
            self.__output_sheet.append(new_list)

    def __q7_write(self):
        self.__output_sheet.append(["7)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[11], row[12], row[13], row[14]]
            self.__output_sheet.append(new_list)

    def __q8_write(self):
        self.__output_sheet.append(["8)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[15], row[16], row[17], row[18]]
            self.__output_sheet.append(new_list)

    def __q9_write(self):
        self.__output_sheet.append(["9)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[19], row[20], row[21], row[22]]
            self.__output_sheet.append(new_list)

    def __q10_write(self):
        self.__output_sheet.append(["10)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[23], row[24], row[25], row[26]]
            self.__output_sheet.append(new_list)

    def __q11_write(self):
        self.__output_sheet.append(["11)Country", "q#11"])
        for row in self.__result:
            new_list = [row[0], row[27]]
            self.__output_sheet.append(new_list)

    def __q12_write(self):
        self.__output_sheet.append(["12)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[28], row[29], row[30], row[31]]
            self.__output_sheet.append(new_list)

    def __q13_write(self):
        self.__output_sheet.append(["13)Country", "minimum", "maximum", "average", "variation"])
        for row in self.__result:
            new_list = [row[0], row[32], row[33], row[34], row[35]]
            self.__output_sheet.append(new_list)

    def __q14_write(self):
        self.__output_sheet.append(["14)Country", "q#14"])
        for row in self.__result:
            new_list = [row[0], row[36]]
            self.__output_sheet.append(new_list)

    def __q15_write(self):
        self.__output_sheet.append(["15)Country", "q#15"])
        for row in self.__result:
            new_list = [row[0], row[37]]
            self.__output_sheet.append(new_list)

    def __q16_write(self):
        self.__output_sheet.append(["16)Country", "q#16"])
        for row in self.__result:
            new_list = [row[0], row[38]]
            self.__output_sheet.append(new_list)

    def __q17_write(self):
        self.__output_sheet.append(["17)Country", "population", "median age", "# of people aged 65 older",
                                    "# of people aged 70 older", "economic performance",
                                    "death rates due to heart disease",
                                    "diabetes prevalence", "# of female smokers", "# of male smokers",
                                    "handwashing facilities","hospital beds per thousand people",
                                    "life expectancy", "human development index"])
        for row in self.__result:
            new_list = [row[0], row[39], row[40], row[41], row[42],
                        row[43], row[44], row[45], row[46], row[47],
                        row[48], row[49], row[50], row[51]]
            self.__output_sheet.append(new_list)

    def __q18_write(self):
        header = ["18) Country", "q#3", "q#4", "q#5_min", "q#5_max", "q#5_avg", "q#5_var",
                  "q#6_min", "q#6_max", "q#6_avg", "q#6_var", "q#7_min", "q#7_max", "q#7_avg", "q#7_var",
                  "q#8_min", "q#8_max", "q#8_avg", "q#8_var", "q#9_min", "q#9_max", "q#9_avg", "q#9_var",
                  "q#10_min", "q#10_max", "q#10_avg", "q#10_var", "q#11", "q#12_min", "q#12_max", "q#12_avg",
                  "q#12_var", "q#13_min", "q#13_max", "q#13_avg", "q#13_var", "q#14", "q#15", "q#16", "population",
                  "median age", "# of people aged 65 older", "# of people aged 70 older", "economic performance",
                  "death rates due to heart disease", "diabetes prevalence", "# of female smokers",
                  "# of male smokers", "handwashing facilities", "hospital beds per thousand people",
                  "life expectancy", "human development index"]
        self.__output_sheet.append(header)
        for row in self.__result:
            self.__output_sheet.append(row)


hw1 = HW1()
hw1.get_results()
