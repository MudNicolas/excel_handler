from openpyxl import load_workbook
from os import listdir, remove
from os.path import exists


def choose_dir():
    exams = listdir('./试卷')
    for index, dir in enumerate(exams):
        print(f'{index}) {dir}')
    choice = input('enter the index of the exam: ')
    return exams[int(choice)]


def delete_file(path):
    if exists(path):
        remove(path)


if __name__ == '__main__':
    dir = f'./试卷/{choose_dir()}/成绩'
    classes = listdir(dir)  # 获取班级列表
    for c in classes:
        try:
            classdir = f'{dir}/{c}'
            files = listdir(classdir)  # 获取班级下的文件列表
            print(files)
        except Exception as e:
            pass
