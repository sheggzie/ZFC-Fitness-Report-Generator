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
font_size_small = 20
top_padding = 30
side_padding = 30
additional_left_padding = 50
row_spacing = 30

# Define the base image paths
base_image_small_path = 'base image.png'
base_image_medium_path = 'base image medium.png'
base_image_large_path = 'base image large.png'

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

    # Determine the font size and the base image path based on the number of rows
    if len(c) == 3:
        font_size = font_size_normal
        base_image_path = base_image_medium_path
    elif len(c) > 4:
        font_size = font_size_normal
        base_image_path = base_image_large_path
    else:
        font_size = font_size_small
        base_image_path = base_image_small_path

    # Load the appropriate font for the header and title
    font_regular = ImageFont.truetype('Inter-ExtraBold.ttf', font_size)
    
    # Load the bold font for the content
    font_bold = ImageFont.truetype('Inter-SemiBold.ttf', font_size)

    # Initialize a dummy ImageDraw object to calculate text size
    dummy_img = Image.new('RGB', (1, 1), color=(255, 255, 255))
    dummy_draw = ImageDraw.Draw(dummy_img)
    
    # Calculate maximum width and total height needed for the text
    max_col_width = max(dummy_draw.textbbox((0, 0), s, font=font_bold)[2] for s, _ in c)
    total_height = sum(dummy_draw.textbbox((0, 0), s, font=font_bold)[3] - dummy_draw.textbbox((0, 0), s, font=font_bold)[1] + row_spacing for s, _ in c)

    # Load the selected base image
    base_img = Image.open(base_image_path)
    base_width, base_height = base_img.size

    # Create a copy of the base image to draw on
    img = base_img.copy()

    # Initialize ImageDraw object
    draw = ImageDraw.Draw(img)

    # Set initial position for text drawing
    x = side_padding
    y = top_padding

    # Draw the sheet name at the top left
    sheet_name_text = f"Hi {df[i]}"
    draw.text((x, y), sheet_name_text, fill=(255, 255, 255), font=font_regular)
    
    # Move to the next line
    y += draw.textbbox((x, y), sheet_name_text, font=font_regular)[3] - draw.textbbox((x, y), sheet_name_text, font=font_regular)[1] + row_spacing

    # Draw the title text with the current month below the sheet name
    title_text = f"{current_month} workout stats"
    draw.text((x, y), title_text, fill=(255, 255, 255), font=font_regular)
    
    # Move to the next line and add additional spacing to separate from the rest of the content
    y += draw.textbbox((x, y), title_text, font=font_regular)[3] - draw.textbbox((x, y), title_text, font=font_regular)[1] + row_spacing * 2

    # Adjust x for additional left padding for the rest of the content
    content_x = x + additional_left_padding

    # Draw each pair of column header and row value, starting below the title text
    for s, t in c:
        # Calculate position for the row value to be right-aligned
        col_text = s
        row_text = str(t)
        row_x = content_x + max_col_width + side_padding
        
        draw.text((content_x, y), col_text, fill=(255, 255, 255), font=font_bold)
        draw.text((row_x, y), row_text, fill=(255, 255, 255), font=font_bold)
        
        # Move to the next line position using textbbox to get the height
        bbox = draw.textbbox((content_x, y), col_text, font=font_bold)
        y += bbox[3] - bbox[1] + row_spacing

    # Save the image
    img.save(df[i] + '.png')

    print("Image created successfully for sheet:", df[i])
    print(" ")
