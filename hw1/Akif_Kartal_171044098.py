import scipy
from scipy import stats
import numpy as np
import matplotlib.pyplot as plt


# I didn't want to bother with global variables therefore,
# I created a class to encapsulate whole homework.
class HW2:

    def __init__(self):
        # write class variables here.
        self.__lines = [[]]
        self.__defects = []
        self.__lamda = 0.0
        self.__predictedCases = []

    def get_results(self):

        self.__q1()
        self.__q2()
        self.__q3()
        # self.__q4()

    def __q1(self):
        self.__readFile()
        for defect in range(5):
            total = 0
            for line in self.__lines:
                for item in line:
                    if int(item) == defect:
                        total += 1
            self.__defects.append(total)
        print("Number of defects:", self.__defects)

    def __q2(self):
        numerator = 0
        for i in range(5):
            numerator += i * self.__defects[i]
        denominator = sum(self.__defects)
        self.__lamda = numerator / denominator
        print("Lambda:", self.__lamda)

    def __q3(self):
        total = sum(self.__defects)
        for i in range(5):
            self.__predictedCases.append(round(total * scipy.stats.poisson.pmf(i, self.__lamda), 2))
        print("Predicted Cases:", self.__predictedCases)

    def __q4(self):
        barWidth = 0.25
        fig = plt.subplots(figsize=(12, 8))
        real = self.__defects
        predict = self.__predictedCases
        br1 = np.arange(len(real))
        br2 = [x + barWidth for x in br1]

        plt.bar(br1, real, color='g', width=barWidth,
                edgecolor='grey', label='Real Cases')
        plt.bar(br2, predict, color='b', width=barWidth,
                edgecolor='grey', label='Predicted Cases')

        plt.xlabel('Number of Defects', fontweight='bold', fontsize=15)
        plt.ylabel('Total Number Of Cases', fontweight='bold', fontsize=15)
        plt.xticks([r + barWidth for r in range(len(real))],
                   ['0', '1', '2', '3', '4'])

        plt.legend()
        plt.show()

    def __readFile(self):
        file = open("manufacturing_defects.txt", "r")
        for line in file:
            if len(line) > 0:
                self.__lines.append(line.split()[2:])
        self.__lines = self.__lines[1:len(self.__lines) - 1]


hw2 = HW2()
hw2.get_results()
