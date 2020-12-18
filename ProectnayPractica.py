import fitz
from PIL import Image
import pytesseract
import cv2
import openpyxl
pytesseract.pytesseract.tesseract_cmd=r'/usr/local/Cellar/tesseract/4.1.1/bin/tesseract'
custom_config =r'--oem 3 --psm 6'

# Получаем адрес таблицы и пдф файла
print('Введите адрес пдф файла')
Pdf=str(input())
print('Введите адрес таблицы')
Xl=str(input())



# Достаём из пдф файла изображение
pdf_document = fitz.open(Pdf)
for current_page in range(len(pdf_document)):
   for image in pdf_document.getPageImageList(current_page):
       xref = image[0]
       pix = fitz.Pixmap(pdf_document, xref)
       if pix.n < 5:
           pix.writePNG("page%s-%s.png" % (current_page, xref))
       else:
           pix1 = fitz.Pixmap(fitz.csRGB, pix)
           pix1.writePNG("page%s-%s.png" % (current_page, xref))
           pix1 = None
       pix = None




# Переворачиваем изображение для дольнейшего анализа
im = Image.open('page1-15.png')
im.transpose(Image.ROTATE_270).save('page1-15.png')


# Получаем номер болезни пациента и записываем в массив
image=Image.open('page1-15.png')
cropped = image.crop((2100, 40, 2200, 70))
cropped.save('number.png')
img= Image.open('number.png')
num= pytesseract.image_to_string(img, config=custom_config)
num=num.split()
array=[0]*12
array[0]=num[0]

# Находим кординаты ключевых слов
img0 = cv2.imread('page1-15.png')
data= pytesseract.image_to_data(img0,config=custom_config)
for i,el in enumerate(data.splitlines()):
    if i==0:
        continue
    el=el.split()
    if 'Prescription' in el:
        x,y,z=int(el[6]),int(el[7]),int(el[9])
    if 'Patient' in el:
        x1,y1=int(el[6]),int(el[7])

# Считывем значения смещения и заносим в массив
cropped1 = image.crop((x1+138, y1+47,x1+310, y1+120))
cropped1.save('table2.png')
img2= Image.open('table2.png')
custom_config1 =r'--oem 3 --psm 7'
num2= pytesseract.image_to_string(img2)
num2=num2.split()
for i in range(6):
    array[i+1]=num2[i]

cropped1 = image.crop((x1+200, y1+145,x1+240, y1+170))
cropped1.save('table3.png')
img2= Image.open('table3.png')
num3= pytesseract.image_to_string(img2)
num3=num3.split()
array[7]=num3[0]

# Считываес данные второй таблицы и заносим в массив
cropped1 = image.crop((x, y+z,2225, y+z+28))
cropped1.save('table1.png')
img2= Image.open('table1.png')
num2= pytesseract.image_to_string(img2)

finish=num2.split()
array[9]=finish[8]
array[10]=finish[9]
array[11]=finish[10]


# Открываем таблицу, ищем свободную строчку и заноcим данные из массива
wb= openpyxl.reader.excel.load_workbook(filename=Xl)
wb.active=0
sheet=wb.active
l=True
n=0
i=1
while l:
    if (sheet['A'+str(i)].value)!=None:
        n+=1
        i+=1
    else: l=False
for i in range(len(array)):
    sheet.cell(n+1,1+i).value = array[i]
wb.save('/Users/arsenijsemenihin/Downloads/tablex.xlsx')











