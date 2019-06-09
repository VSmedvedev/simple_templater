##### vladimir medvedev #####
##### stalker1984@mail.ru #####
import sys, csv, os
from docxtpl import DocxTemplate

#указываем имя файла в CSV формате из которого будут взяты данные
#для заполнения шаблона
fname = 'data.csv'
f = open(fname, 'r')
lis = [line.split(',') for line in f]
f.close()
context=dict()

# range(len(lis)) -> range(3) можно протестировать на первых трёх данных   
for i in range(len(lis)):
    #задайте столько фалов сколько у вас шаблонов
    doc = DocxTemplate("template.docx")
    docZF1 = DocxTemplate("templateZF1.docx")
    docZF2 = DocxTemplate("templateZF2.docx")
    for j in range(len(lis[i])):
        for k in range(len(lis[0][0].split(';'))):
            context[str(lis[0][0].split(';')[k])] = str(lis[i+1][0].split(';')[k])
    #задайте путь к папке где хранится скрипт
    #задайте имя папок в которых будут сохраняться сгенерированне документы, используйте значения из CSV файла для удобства
    # str(lis[i+1][0].split(';')[0] - первый столбик из файла, str(lis[i+1][0].split(';')[1] - второй столбик из файла, дальше по аналогии
    #меняете только последнюю цифру для нужного столбика
    #ВНИМАНИЕ: СЛЭШИ ДОЛЖНЫ БЫТЬ ОБРАТНЫЕ, А НЕ ПРЯМЫЕ!!!
    path = "C:/Users/526/Desktop/шаблоны МКК плазмид/"+ str(lis[i+1][0].split(';')[0]) + str(lis[i+1][0].split(';')[1]) + "/"
    os.mkdir(path)
    #рэндрим содержимое в файлы (ДОБАВТИ ЕЩЕ ЕСЛИ НУЖНО)
    doc.render(context)
    docZF1.render(context)
    docZF2.render(context)
    #заводим имена файлов для сохранения, используйте данные из CSV файла(ДОБАВТИ ЕЩЕ ЕСЛИ НУЖНО).
    fname =    path  + str(lis[i+1][0].split(';')[0] + ' ' + lis[i+1][0].split(';')[1]) #отформатировать имя файла
    fnameZF1 = path  + "ЗФ1-" + str(lis[i+1][0].split(';')[0] + " Внешний вид")
    fnameZF2 = path  + "ЗФ2-" + str(lis[i+1][0].split(';')[0] + " Специфическая активность")
    #сохраняем хайлы (ДОБАВТИ ЕЩЕ ЕСЛИ НУЖНО)
    doc.save(fname + '.docx')
    docZF1.save(fnameZF1 + '.docx')
    docZF2.save(fnameZF2 + '.docx')
    #делаем принт чтобы можно было в консоле наблюдать за ходом выполнения программы
    print ('файл ' + str(i+1))
    ### ОБЯЗАТЕЛЬНО УДАЛЯЙТЕ СГЕНЕРИРОВАННЫЕ ДОКУМЕНТЫ ПЕРЕД СЛЕДУЮЩЕЙ ГЕНЕРАЦИЕЙ ###
