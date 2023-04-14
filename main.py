from PIL import Image, ImageDraw, ImageFont
import openpyxl
import os

os.chdir("./")

cert_image = Image.open("certificate_template_name.jpg")
print("Loaded Certificate")

wb = openpyxl.load_workbook("excel_sheet_name.xlsx")
print("Loaded Sheet")
sheet = wb.active

name_font = ImageFont.truetype("Montserrat SemiBold 600.ttf", 60)
inst_font = ImageFont.truetype("Montserrat SemiBold 600.ttf", 32)
prog_font = ImageFont.truetype("Montserrat SemiBold 600.ttf", 40)
print("Font and Sizes Added")

os.chdir("./certificates")
count = 1
for row in sheet.iter_rows(values_only=True):
    name = row[0]
    institution = row[1]
    program = row[2]

    img = Image.new('RGB', cert_image.size, color=(255, 255, 255))
    img.paste(cert_image, (0,0))

    draw = ImageDraw.Draw(img)
    name_width, name_height = draw.textsize(name, font=name_font)
    inst_width, inst_height = draw.textsize(institution, font=inst_font)
    prog_width, prog_height = draw.textsize(program, font=prog_font)
    draw.text(((img.width - name_width) / 2, 625), name, font=name_font, fill=(255,255,255))
    draw.text((((img.width - inst_width) / 2)-478, 797), institution, font=inst_font, fill=(255,255,255))
    draw.text((((img.width - prog_width) / 2)+550, 790), program, font=prog_font, fill=(255,255,255))

    img.save(f"{name}.jpg")
    print(f"Process {count} Completed")
    count += 1
