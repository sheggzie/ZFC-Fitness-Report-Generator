from PIL import Image, ImageDraw, ImageFont
import pandas as pd

xls = pd.ExcelFile('specimen.xlsx')

df = xls.sheet_names

for i in range(len(df)):     
    sheet = pd.read_excel(xls, sheet_name=df[i])  
    
    row = sheet.iloc[31, 4:].to_string(index=False)
    col = sheet.columns
    
    # a = col[4:]
    a = []

    c = zip(a.append(col[4:]), row)   
    
    # Load the font you want to use
    font = ImageFont.truetype('arial.ttf', 25)

    # Create an empty image
    img = Image.new('RGB', (600, 500), color = (255, 255, 255))

    # Initialize ImageDraw object
    # draw = ImageDraw.Draw(img)
    
    # Set initial position for text drawing
    x = 10
    y = 10

    for s, t in c:
        text = (f"{s:<20}{t:>20}")
        print(len(s), len(t))

    
    # draw.text((x, y), text, fill=(0, 0, 0), font=font)

    # Save the image
    # img.save(df[i]+'.png')

    # print("Image created successfully!")  