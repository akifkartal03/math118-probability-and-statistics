import openpyxl

wb_obj = openpyxl.load_workbook("Grades.xlsx", data_only=True)
sheet = wb_obj.active
values = sheet["K"]
grades = sheet["L"]
myset = set()
indexes = []
i = 1
for element in values:
    size = len(myset)
    myset.add(element.value)
    if len(myset) != size:
        indexes.append(i - 1)
    i = i + 1

# print(indexes[2:])
for index in indexes[2:]:
    print(values[index - 1].value, " ", grades[index - 1].value)

"""
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

for user in values:
    myset.add(user.value)
print("Normal Set:")
print(myset)
print("Sorted:")
last = sorted(myset, key=lambda d: tuple(map(int, d.split('-'))))
print(last)
"""
