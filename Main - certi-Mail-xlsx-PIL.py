from PIL import Image, ImageDraw, ImageFont
import xlrd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import io

dataArr =[]
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
    if i==0:
        continue
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


    #SEND MAIL-----------------------
    fromaddr = "agr@gmail.com" #your email address
    toaddr = i[2]
    print(i)

    msg = MIMEMultipart()
    msg['From'] = '"From Anurag Gupta"' #The from tag which the reciever will see
    msg['To'] = toaddr
    msg['Subject'] = "Certificate of Participation, "+i[1]+", "+" TCF'20"
    body = "Here is your certificate of participation for " +i[1]+ ", Corona & Melange, TCF20, NIT Patna. Please find the certificate attached. In case there's an error, please contact the event coordinator."
    msg.attach(MIMEText(body, 'plain'))
    filename = "Certificate.jpg"
    buf = io.BytesIO()
    img.save(buf, format='JPEG')
    byte_im = buf.getvalue()
    attachment = byte_im
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment))
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(p)
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(fromaddr, "your_password") #put your password in this
    text = msg.as_string()
    s.sendmail(fromaddr, toaddr, text)
    print("sent {}".format(ind+1))
