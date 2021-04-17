import openpyxl
import statistics
from openpyxl.workbook import Workbook

# I didn't want to bother with global variables therefore,
# I created a class to encapsulate whole homework.
class HW1:

    def __init__(self):
        # write class variables here.
        self.__data = None
        self.__countries = []
        self.__summary = []

    def get_results(self):
        # this is only public method to calculate all results
        sheet = openpyxl.load_workbook("owid-covid-data.xlsx", data_only=True)
        self.__data = sheet.active
        self.__countries = self.__get_list(self.__data["C"])
        self.__summary.append(["Total Country"])
        self.__summary.append([])
        self.__summary.append(["Earliest Date","Country"])
        self.__q1()
        self.__q2()
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
            res = [min_date,self.__countries[index]]
            self.__summary.append(res)

    def __q3(self):
        self.__common_3_4_11("E")

    def __q4(self):
        self.__common_3_4_11("H")

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
        self.__common_3_4_11("AA")

    def __q12(self):
        self.__common_5to13("AF")

    def __q13(self):
        self.__common_5to13("AG")

    def __q14(self):
        self.__common_3_4_11("AJ")

    def __q15(self):
        self.__common_3_4_11("AK")

    def __q16(self):
        self.__common_3_4_11("AI")

    def __q17(self):
        self.__common_3_4_11("AS")
        self.__common_3_4_11("AU")
        self.__common_3_4_11("AV")
        self.__common_3_4_11("AW")
        self.__common_3_4_11("AX")
        self.__common_3_4_11("AZ")
        self.__common_3_4_11("BA")
        self.__common_3_4_11("BB")
        self.__common_3_4_11("BC")
        self.__common_3_4_11("BD")
        self.__common_3_4_11("BE")
        self.__common_3_4_11("BF")
        self.__common_3_4_11("BG")

    """
    
    
    """

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
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country)
            if len(my_set) == size or i == 0:
                if rate[i] is not None:
                    rate_list.append(rate[i])
            else:
                ct_name = self.__countries[i - 1]
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
            i = i + 1

    def __common_3_4_11(self, column):
        total = self.__get_list(self.__data[column])
        my_set = set()
        indexes = []
        i = 1
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country)
            if len(my_set) != size:
                indexes.append(i - 1)
            i = i + 1
        for index in indexes[1:]:
            print(self.__countries[index - 1], " ", total[index - 1])
    def __write_summary(self):
        wb = Workbook()
        dest_filename = 'output_book.xlsx'
        ws1 = wb.active
        for row in self.__summary:
            ws1.append(row)
        wb.save(filename=dest_filename)
a = HW1()
a.get_results()
