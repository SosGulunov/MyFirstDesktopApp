import random
import pandas as pd
import sys

sys.setrecursionlimit(4000)


class Chelik:
    def __init__(self, name, a=0, b=0):
        self.name = name
        self.a = a
        self.b = b
        self.count = 0
        self.canWork = []

    def app(self):
        self.count += 1

    def zeroCount(self):
        self.count = 0


# Workers.txt
taimingi = [[8, 10], [10, 12], [12, 15], [15, 17], [17, 18], [18, 20], [20, 24]]
my_file = open("Workers.txt", encoding="utf8")
n = my_file.read().split('\n')
# n = ["Вадим","Андрей","Эдик","Карина","София","Ксюша","Алеся","Алина","Виталий","Ирина Роман","Настя","Милана","Настя Ф","Оля","Мария"]
ChelikiStd = [Chelik(n[i]) for i in range((len(n)))]
Cheliki = ChelikiStd
RestCheliki = []


# Cheliki = [Chelik("Вадим",10,15),Chelik("Андрей",17,24),Chelik("Карина",18,24),Chelik("Эдик",20,24),Chelik("София",15,18),Chelik("Ксюша",12,17),Chelik("Алеся",12,20),Chelik("Алина",10,15),Chelik("Виталий",8,15),Chelik("Ирина Роман",8,18),Chelik("Настя",8,15),Chelik("Милана",12,18),Chelik("Настя Ф",15,20),Chelik("Оля",8,24),Chelik("Мария",8,24)]
def NewTable():
    for i in range(len(Cheliki)):
        Cheliki[i].zeroCount()
        Cheliki[i].canWork.clear()


def sort(Cheliki, taimingi):
    random.shuffle(Cheliki)
    for i in range(7):
        for j in range(len(Cheliki)):
            if ((Cheliki[j].a == taimingi[i][0] and Cheliki[j].b <= taimingi[i][1]) or (
                    Cheliki[j].a >= taimingi[i][0] and Cheliki[j].b == taimingi[i][1]) or (
                    Cheliki[j].a == taimingi[i][0] or Cheliki[j].b == taimingi[i][1]) or (
                    Cheliki[j].a < taimingi[i][0] and Cheliki[j].b > taimingi[i][1])):
                Cheliki[j].canWork.append(1)
            else:
                Cheliki[j].canWork.append(0)


def add2(Cheliki):
    mas = ["", "", "", "", "", ""]
    for j in range(len(Cheliki)):
        for i in range(6):
            if (Cheliki[j].canWork[i + 1] == 1) and (Cheliki[j].count < 3) and mas[i] == "":
                mas[i] = Cheliki[j].name
                Cheliki[j].canWork[i + 1] = 0
                Cheliki[j].app()
    return mas


def firstRow(Cheliki, taimingi):
    for j in range(len(Cheliki)):
        if Cheliki[j].a == taimingi[0][0]:
            Cheliki[j].app()
            Cheliki[j].app()
            return Cheliki[j].name


def main():
    # var1 = "ОБЩАЯ Биология Химия Информат История Общество БУ"
    # var2 = "ОБЩАЯ Математика Физика ОГЭ Хиструктор Мутаген Английский"
    # var3 = "ОБЩАЯ Родители Эконом Русский УУ матем УУ рус УУ общ"
    var1 = "Инфа"
    var2 = "Матем"
    var3 = "Общая"
    t = ["8-10", "10-12", "12-15", "15-17", "17-18", "18-20", "20-24"]

    MorningHuman = firstRow(Cheliki, taimingi)
    sort(Cheliki, taimingi)
    df = pd.DataFrame({'Time': t,
                       var1: [MorningHuman] + add2(Cheliki),
                       var2: [MorningHuman] + add2(Cheliki),
                       var3: [MorningHuman] + add2(Cheliki)})
    df.to_excel('./Result.xlsx', index=False)
    NewTable()
    sort(Cheliki, taimingi)


if __name__ == "__main__":
    main()