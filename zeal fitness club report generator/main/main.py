from PIL import Image, ImageDraw, ImageFont
import pandas as pd

sheet = pd.read_excel('april.xlsx', sheet_name='Actually')

text = sheet.iloc[31, [4, 5, 6]]

# Load the font you want to use
font = ImageFont.load_default()

# Create an empty image
img = Image.new('RGB', (600, 500), color = (255, 255, 255))

# Initialize ImageDraw object
draw = ImageDraw.Draw(img)

# Set initial position for text drawing
x = 10
y = 10

draw.text((x, y), text.to_string(), fill=(0, 0, 0), font=font)

# Save the image
img.save('output_image_new.png')

print("Image created successfully!")










# text = sheet.iloc[31, [4, 5, 6]]

# print(text.to_string())