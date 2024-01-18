import openpyxl
import PIL
from PIL import Image, ImageFont, ImageDraw


workbook_students = openpyxl.load_workbook('sheet.xlsx')
work_sheet = workbook_students['Planilha1']

for index, line in enumerate (work_sheet.iter_rows(min_row=2)):
    course_name = line[0].value
    participant_name = line[1].value
    finish_date = line[2].value
    finish_date_str = finish_date.strftime('%Y-%m-%d')
    certificate_code = line[3].value
    certificate_code_str = str(certificate_code)
    font_name = ImageFont.truetype('./fonts/Roboto-Bold.ttf', 90)
    general_font = ImageFont.truetype('./fonts/Roboto-Regular.ttf', 50)

    image = Image.open('./certificate_model.png')
    draw = ImageDraw.Draw(image)

    draw.text((760,600), participant_name, fill="black", font=font_name)
    draw.text((760,765), course_name, fill="black", font=general_font)
    draw.text((520,850), finish_date_str, fill="black", font=general_font)
    draw.text((910,1180), certificate_code_str, fill="black", font=general_font)
    image.save(f'./certifications/{participant_name}{index}.png')

    

