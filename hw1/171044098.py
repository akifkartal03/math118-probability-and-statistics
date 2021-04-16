import openpyxl


class HW1:

    def __init__(self):
        # write class variables here.
        self.__data = None

    def get_results(self):
        # only public method
        sheet = openpyxl.load_workbook("owid-covid-data.xlsx", data_only=True)
        self.__data = sheet.active
        self.__q1()

    def __q1(self):
        first_column = self.__data["A"]
        countries = self.__get_set(first_column)
        print("len: ", len(countries))

    def __q2(self):
        dates = self.__data["D"]
        my_list = []
        for date in dates:
            my_list.append(date.value)
        # get rid of repeated elements
        # unique_dates = self.__get_set(dates)
        # sorted_dates = sorted(unique_dates, key=lambda x: tuple(map(int, x.split('-'))))
        min_date = min(my_list)
        indexes = [i for i, x in enumerate(my_list) if x == min_date]
        countries = self.__data["C"]
        for index in indexes:
            print(countries[index].value)

    """
    def __q3(self):
    def __q4(self):
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


a = HW1()
a.get_results()
