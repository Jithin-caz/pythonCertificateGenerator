from PIL import Image, ImageDraw, ImageFont
import openpyxl

# Load the poster template
poster_path = "2.png"
poster_template = Image.open(poster_path)

# Load the Excel spreadsheet
spreadsheet_path = "IEDC 2024 Passouts.xlsx"
wb = openpyxl.load_workbook(spreadsheet_path)
sheet = wb.active


text_color = (45, 45, 45)  
font_path = "American Captain.otf" 
font_size = 200
font = ImageFont.truetype(font_path, font_size)


name_position = (620,433)



for row in sheet.iter_rows(min_row=2, values_only=True):  
 
    poster = poster_template.copy()
    draw = ImageDraw.Draw(poster)

  
    name, phone = row[:2] 

   
    draw.text(name_position, f"{name}", fill=text_color, font=font)
   
    poster.save(f"certificates/{name}_poster.png") 
