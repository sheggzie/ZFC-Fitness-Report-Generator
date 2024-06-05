from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from datetime import datetime

# Get the current month name
current_month = datetime.now().strftime('%B')

# Load the Excel file
xls = pd.ExcelFile('may report.xlsx')

# Get sheet names
df = xls.sheet_names

# Define font sizes and padding
font_size_normal = 25
font_size_small = 15
padding = 30
row_spacing = 30
min_image_height = 200  # Minimum height to ensure space for very few rows

for i in range(len(df)):     
    # Read each sheet into a dataframe
    sheet = pd.read_excel(xls, sheet_name=df[i])  
    
    # Extract column headers starting from the fifth column
    col = sheet.columns[4:]
    
    # Extract the 32nd row data starting from the fifth column
    row = sheet.iloc[31, 4:].values

    # Remove decimal suffix by converting float to int where applicable
    row = [int(value) if isinstance(value, float) else value for value in row]

    # Zip the column headers and row data together
    c = list(zip(col, row))

    # Determine the font size based on the number of rows
    font_size = font_size_small if len(c) < 3 else font_size_normal

    # Load the appropriate font
    font = ImageFont.truetype('arial.ttf', font_size)

    # Initialize a dummy ImageDraw object to calculate text size
    dummy_img = Image.new('RGB', (1, 1), color=(255, 255, 255))
    dummy_draw = ImageDraw.Draw(dummy_img)
    
    # Calculate maximum width and total height needed for the text
    max_col_width = max(dummy_draw.textbbox((0, 0), s, font=font)[2] for s, _ in c)
    total_height = sum(dummy_draw.textbbox((0, 0), s, font=font)[3] - dummy_draw.textbbox((0, 0), s, font=font)[1] + row_spacing for s, _ in c)
    total_height += dummy_draw.textbbox((0, 0), f"Name: {df[i]}", font=font)[3] - dummy_draw.textbbox((0, 0), f"Sheet: {df[i]}", font=font)[1] + row_spacing
    total_height += dummy_draw.textbbox((0, 0), f"{current_month} workout stats.", font=font)[3] - dummy_draw.textbbox((0, 0), f"Workout stats for the month of {current_month}", font=font)[1] + row_spacing

    # Ensure minimum image height
    img_height = max(total_height + padding * 2, min_image_height)
    img_width = max_col_width * 2 + padding * 3

    # Create an empty image with calculated dimensions
    img = Image.new('RGB', (img_width, img_height), color=(255, 255, 255))

    # Initialize ImageDraw object
    draw = ImageDraw.Draw(img)
    
    # Set initial position for text drawing, centered vertically
    y = (img_height - total_height) // 2 if img_height > total_height else padding
    x = padding

    # Draw the sheet name at the top
    sheet_name_text = f"Name: {df[i]}"
    draw.text((x, y), sheet_name_text, fill=(0, 0, 0), font=font)
    
    # Move to the next line
    y += draw.textbbox((x, y), sheet_name_text, font=font)[3] - draw.textbbox((x, y), sheet_name_text, font=font)[1] + row_spacing

    # Draw the title text with the current month
    title_text = f"{current_month} workout stats."
    draw.text((x, y), title_text, fill=(0, 0, 0), font=font)
    
    # Move to the next line
    y += draw.textbbox((x, y), title_text, font=font)[3] - draw.textbbox((x, y), title_text, font=font)[1] + row_spacing

    # Draw each pair of column header and row value
    for s, t in c:
        # Calculate position for the row value to be right-aligned
        col_text = s
        row_text = str(t)
        row_x = x + max_col_width + padding
        
        draw.text((x, y), col_text, fill=(0, 0, 0), font=font)
        draw.text((row_x, y), row_text, fill=(0, 0, 0), font=font)
        
        # Move to the next line position using textbbox to get the height
        bbox = draw.textbbox((x, y), col_text, font=font)
        y += bbox[3] - bbox[1] + row_spacing

    # Save the image
    img.save(df[i] + '.png')

    print("Image created successfully for sheet:", df[i])
    print(" ")
