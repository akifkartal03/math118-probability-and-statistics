import openpyxl

wb_obj = openpyxl.load_workbook("Grades.xlsx", data_only=True)
sheet = wb_obj.active
values = sheet["I"]
myset = set()
mylist = []
for user in values:
    mylist.append(user.value)
# print(mylist)
min_date = min(mylist)
print(min_date)
indexes = [i for i, x in enumerate(mylist) if x == min_date]
countries = sheet["B"]
for index in indexes:
    print(countries[index].value)
print(indexes)
"""
for user in values:
    myset.add(user.value)
print("Normal Set:")
print(myset)
print("Sorted:")
last = sorted(myset, key=lambda d: tuple(map(int, d.split('-'))))
print(last)
"""
