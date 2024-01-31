import openpyxl
from PIL import Image
from PIL import ImageDraw 
from PIL import ImageFont

woorkbook_students = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_students = woorkbook_students['Sheet1']
 

for indice, line in enumerate(sheet_students.iter_rows(min_row=2)):
    name_course = line[0].value
    name_participant = line[1].value
    type_participation = line[2].value
    data_home = line[3].value
    data_final = line[4].value
    qtd_hours = line[5].value
    data_emition = line[6].value

    font_name = ImageFont.truetype('./tahomabd.ttf')
    font_geral = ImageFont.truetype('./tahoma.ttf')
    font_data = ImageFont.truetype('./tahoma.ttf')

    image = Image.open('./certificado_padrao.jpg')
    draw = ImageDraw.Draw(image)

    draw.text((200,600), name_participant, fill='black', font=font_name)
    draw.text((1060,950), name_course, fill='black', font=font_geral)
    draw.text((1435,1065), type_participation, fill='black', font=font_geral)
    draw.text((14080, 1182),str(qtd_hours),fill='black',font=font_geral)


    draw.text((750, 1770), data_home,fill='black',font=font_data)
    draw.text((750,1930),data_final,fill='black',font=font_data)

    draw.text((2200,1930),data_emition, fill='blue',font=font_data)


    image.save(f'./{indice}{name_participant}certificado.png')

