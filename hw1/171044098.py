import openpyxl
import statistics

class HW1:

    def __init__(self):
        # write class variables here.
        self.__data = None
        self.__countries = []

    def get_results(self):
        # this is only public method to calculate all results
        sheet = openpyxl.load_workbook("owid-covid-data.xlsx", data_only=True)
        self.__data = sheet.active
        self.__countries = self.__get_list(self.__data["C"])
        self.__q1()
        self.__q2()
        self.__q3()
        self.__q4()

    def __q1(self):
        first_column = self.__data["A"]
        unique_countries = self.__get_set(first_column)
        print("len: ", len(unique_countries))

    def __q2(self):
        dates = self.__get_list(self.__data["D"])
        # get rid of repeated elements
        # unique_dates = self.__get_set(dates)
        # sorted_dates = sorted(unique_dates, key=lambda x: tuple(map(int, x.split('-'))))
        min_date = min(dates)
        print(min_date)
        indexes = [i for i, date in enumerate(dates) if date == min_date]
        for index in indexes:
            print(self.__countries[index])

    def __q3(self):
        total_cases = self.__get_list(self.__data["E"])
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
            print(self.__countries[index - 1], " ", total_cases[index - 1])

    def __q4(self):
        total_deaths = self.__get_list(self.__data["H"])
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
            print(self.__countries[index - 1], " ", total_deaths[index - 1])

    def __q5(self):
        reproduction_rate = self.__get_list(self.__data["Q"])
        my_set = set()
        rate_list = []
        i = 0
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country)
            if len(my_set) == size or i == 0:
                rate_list.append(reproduction_rate[i])
            else:
                avg = round(sum(rate_list)/len(rate_list),2)
                minimum = min(rate_list)
                maximum = max(rate_list)
                variation = statistics.variance(rate_list)
                rate_list = [reproduction_rate[i]]
            i = i + 1
    """
    def __q6(self):
    def __q7(self):
    def __q8(self):
    def __q9(self):
    def __q10(self):
    def __q11(self):
    def __q12(self):
    def __q13(self):
    def __q14(self):
    def __q15(self):
    def __q16(self):
    def __q17(self):
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



a = HW1()
a.get_results()
