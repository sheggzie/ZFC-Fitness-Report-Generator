from PIL import Image, ImageDraw, ImageFont
import pandas as pd

xls = pd.ExcelFile('specimen.xlsx')

df = xls.sheet_names

for i in range(len(df)):     
    sheet = pd.read_excel(xls, sheet_name=df[i])  
    head_rows = sheet.columns

    row = sheet.iloc[31, 4:].to_string(index=False)
    col = head_rows[4:]

    a = col
    b = row       
    c = zip(a, b)
    
    # Load the font you want to use
    font = ImageFont.truetype('arial.ttf', 25)

    # Create an empty image
    img = Image.new('RGB', (600, 500), color = (255, 255, 255))

    # Initialize ImageDraw object
    draw = ImageDraw.Draw(img)
    
    # Set initial position for text drawing
    x = 10
    y = 10

    for s, t in c:
        text = (f"{s:<25}{t:>25}")
        print(text)
    
    
    draw.text((x, y), text, fill=(0, 0, 0), font=font)

    # Save the image
    img.save(df[i]+'.png')

    print("Image created successfully!")
    print(" ")  