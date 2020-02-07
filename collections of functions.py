###### Работа с os и список файлов в папке
def getListFileXlsxInPath(path):
    import os
    files = os.listdir(path)
    xlsfiles = filter(lambda x: x.endswith('.xlsx'), files)
    return xlsfiles


path = r'C:\Users\RostPK\Desktop\for py'
list = getListFileXlsxInPath(path)
for filename in list:
    print(filename.decode("cp1251"))
######
###### Работа с файлом
input_file = open(r'input.txt', 'r')  # w write and clear
a = input_file.read() # read all file in a
print (a)
# input_file.write('bla bla bla' + '\n') write in file

input_file.close()
######
###### Работа с кодировкой
a.decode('utf-8')
a.encode('utf-8')
######
###### Работа с исключениями
try:
    print (5/0)
except Exception:
    print('Ошибка')
######
###### Работа со времнем
print (time.strftime("%Y.%m.%d %H:%M:%S", time.localtime()))
