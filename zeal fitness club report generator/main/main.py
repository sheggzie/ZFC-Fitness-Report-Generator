from PIL import Image, ImageDraw, ImageFont
import pandas as pd

# Load the Excel file
xls = pd.ExcelFile('may report.xlsx')

# Get sheet names
df = xls.sheet_names

for i in range(len(df)):     
    # Read each sheet into a dataframe
    sheet = pd.read_excel(xls, sheet_name=df[i])  
    
    # Extract column headers starting from the fifth column
    col = sheet.columns[4:]
    
    # Extract the 32nd row data starting from the fifth column
    row = sheet.iloc[31, 4:].values

    # Zip the column headers and row data together
    c = list(zip(col, row))
    
    # Load the font you want to use
    font = ImageFont.truetype('arial.ttf', 25)

    # Create an empty image
    img = Image.new('RGB', (800, 600), color=(255, 255, 255))

    # Initialize ImageDraw object
    draw = ImageDraw.Draw(img)
    
    # Calculate the maximum width of the column headers
    max_col_width = max(draw.textbbox((0, 0), s, font=font)[2] for s, _ in c)

    # Set initial position for text drawing
    x = 10
    y = 10
    padding = 80  # Padding between column header and row value
    row_spacing = 20  # Spacing between rows

    # Draw each pair of column header and row value
    for s, t in c:
        # Calculate position for the row value to be right-aligned
        col_text = s
        row_text = str(t)
        col_width = draw.textbbox((0, 0), col_text, font=font)[2]
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