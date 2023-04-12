li = []
def FileOpener(file): #Открываем файл для создания списка значений
    f = open(f'{file}')
    for i in f:
        li.append(int(i))
    f.close()

def Fixer(): #Алгоритм исправления температуры
    f = open('fixed.txt', 'w')
    a = [int((li[x - 1] + li[x + 1])/2) if i == 300 else i for x, i in enumerate(li)]
    f.write("\n".join(map(str,a)))
    f.close()

FileOpener('file.txt')
Fixer()
