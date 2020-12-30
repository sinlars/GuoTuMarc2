import shutil


def copy2019():
    file = open('F:\PycharmProjects\GuoTuMarc\Marc\errors\文库2019年图书.txt', 'r')
    lines = file.readlines()
    file.close()
    print(lines)