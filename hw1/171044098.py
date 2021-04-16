import openpyxl


class HW1:

    def __init__(self):
        # write class variables here.
        self.__data = None
        self.__countries = None

    def get_results(self):
        # this is only public method to calculate all results
        sheet = openpyxl.load_workbook("owid-covid-data.xlsx", data_only=True)
        self.__data = sheet.active
        self.__countries = self.__data["C"]
        self.__q1()

    def __q1(self):
        first_column = self.__data["A"]
        countries = self.__get_set(first_column)
        print("len: ", len(countries))

    def __q2(self):
        dates = self.__data["D"]
        my_list = self.__get_list(dates)
        # get rid of repeated elements
        # unique_dates = self.__get_set(dates)
        # sorted_dates = sorted(unique_dates, key=lambda x: tuple(map(int, x.split('-'))))
        min_date = min(my_list)
        print(min_date)
        indexes = [i for i, date in enumerate(my_list) if date == min_date]
        for index in indexes:
            print(self.__countries[index].value)

    def __q3(self):
        total_cases = self.__data["E"]
        my_set = set()
        indexes = []
        i = 1
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country.value)
            if len(my_set) != size:
                indexes.append(i - 1)
            i = i + 1
        for index in indexes[2:]:
            print(self.__countries[index - 1].value, " ", total_cases[index - 1].value)

    def __q4(self):
        total_deaths = self.__data["H"]
        my_set = set()
        indexes = []
        i = 1
        for country in self.__countries:
            size = len(my_set)
            my_set.add(country.value)
            if len(my_set) != size:
                indexes.append(i - 1)
            i = i + 1
        for index in indexes[2:]:
            print(self.__countries[index - 1].value, " ", total_deaths[index - 1].value)

    """

    def __q5(self):
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
        for element in data:
            my_set.add(element.value)
        return my_set

    def __get_list(self, data):
        my_list = []
        for element in data:
            my_list.append(element.value)
        return my_list


a = HW1()
a.get_results()
