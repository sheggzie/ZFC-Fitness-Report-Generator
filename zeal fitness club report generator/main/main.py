from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

# Get the current month name
current_month = "August"
# current_month = datetime.now().strftime('%B')

# Load the Excel file
xls = pd.ExcelFile(os.path.join('resources', 'august report.xlsx'))

# Get sheet names
df = xls.sheet_names

# Define font sizes and padding
font_size_xl = 400
font_size_large = 90
font_size_normal = 80
font_size_small = 70
font_size_mini = 30
top_padding = 160
side_padding = 200
additional_left_padding = 120
row_spacing = 90

# Define the base image paths
base_image_small_path = os.path.join('resources', 'base image.png')
base_image_pro_path = os.path.join('resources', 'base image pro.png')
base_image_medium_path = os.path.join('resources', 'base image medium.png')
base_image_large_path = os.path.join('resources', 'base image large.png')
base_image_medium_pro_path = os.path.join('resources', 'base image medium pro.png')
base_image_extra_large_path = os.path.join('resources', 'base image extra large.png')


# Define the font paths
font_regular_path = os.path.join('resources', 'Inter-ExtraBold.ttf')
font_bold_path = os.path.join('resources', 'Inter-SemiBold.ttf')

# Define the width threshold for using the wide image
width_threshold = 1200

for i in range(len(df)):     
    # Read each sheet into a dataframe
    sheet = pd.read_excel(xls, sheet_name=df[i])  
    
    # Extract column headers starting from the fifth column
    col = sheet.columns[4:]
    
    # Extract the 32nd row data starting from the fifth column
    row = sheet.iloc[31, 4:].values

    # Remove decimal suffix by converting float to int where applicable
    row = [str(int(value)) if isinstance(value, float) else value for value in row]

    # Zip the column headers and row data together
    c = list(zip(col, row))

    # Load the appropriate font for the header and title
    font_regular = ImageFont.truetype(font_regular_path, font_size_normal)
    
    # Load the bold font for the content
    font_bold = ImageFont.truetype(font_bold_path, font_size_normal)

    # Initialize a dummy ImageDraw object to calculate text size
    dummy_img = Image.new('RGB', (1, 1), color=(255, 255, 255))
    dummy_draw = ImageDraw.Draw(dummy_img)
    
    # Calculate maximum width needed for the text
    max_col_width = max(dummy_draw.textbbox((0, 0), s, font=font_bold)[2] for s, _ in c)
    
    # Determine the font size and the base image path based on the number of rows and text width
    if max_col_width > width_threshold:
        base_image_path = base_image_extra_large_path
    elif len(c) == 3 or len(c) <= 2:
        font_size = font_size_large
        base_image_path = base_image_small_path
    elif len(c) == 4 or len(c) == 5:
        font_size = font_size_small
        base_image_path = base_image_small_path
    elif len(c) == 6:
        font_size = font_size_xl
        base_image_path = base_image_pro_path
    elif len(c) == 8:
        font_size = font_size_xl
        base_image_path = base_image_medium_path
    elif len(c) == 24:
        font_size = font_size_large
        base_image_path = base_image_medium_pro_path

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
    img.save(os.path.join('resources', df[i] + '.png'))

    print("Image created successfully for sheet:", df[i])
    print(" ")
