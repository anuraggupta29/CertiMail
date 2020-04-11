from PIL import Image, ImageDraw, ImageFont
import xlrd

dataArr =[]
folder = 'certificates/'
# Give the location of the file
loc = ("excelfile.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# get no of rows
rows = sheet.nrows

#store name,event,email in array
for i in range(rows):
    dataArr.append((sheet.cell_value(i, 1), sheet.cell_value(i, 6), sheet.cell_value(i, 5)))


#GET IMG FILE
imgOriginal = Image.open('cop.jpg')

#INITIALIZE POSITION OF TEXT
pos1__, pos2__ = (1754,1360), (1754,1020)

#SELECT FONT
color = (26,26,26)
fnt = ImageFont.truetype('TCM.TTF', 96)

#LOOP----------------------------------------
for ind,i in enumerate(dataArr):
    img = imgOriginal.copy()
    name = i[0]
    event = i[1]
    txt1, txt2 = str(name).upper(), str(event).upper()

    (width1, height1) = fnt.font.getsize(txt1)[0]
    (width2, height2) = fnt.font.getsize(txt2)[0]

    pos1, pos2 = (1754 - int(width1/2), 2480 - pos1__[1] - 96), (1754 - int(width2/2), 2480 - pos2__[1] - 96)

    d = ImageDraw.Draw(img)
    d.text(pos1, txt1, font=fnt, fill=color)
    d.text(pos2, txt2, font=fnt, fill=color)

    #SAVE THE IMAGE
    img.save(mainEvent+'/'+'certificate'+str(ind+1)+'.jpg')
