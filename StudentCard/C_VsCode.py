import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Load the Excel file
data = pd.read_excel("data.xlsx", sheet_name="SDATA")  # Adjust sheet name as needed

# Define dimensions for the card
width, height = 500, 250
border_size = 2
border_color = (0, 255, 0)  # Green color in RGB
font_size = 20
text_color = (0, 0, 0)  # Black color in RGB

# Iterate over each row in the DataFrame
for index, row in data.iterrows():
    # Create a new transparent image for each card
    card = Image.new("RGBA", (width, height), (0, 0, 0, 0))  # Transparent background
    draw = ImageDraw.Draw(card)

    # Draw a green border
    draw.rectangle(
        [(border_size // 2, border_size // 2), (width - border_size // 2, height - border_size // 2)],
        outline=border_color,
        width=border_size
    )

    # Load a bold font
    font = ImageFont.truetype("arialbd.ttf", 20)  # Use bold Arial or another bold font
    
    # Print the "ID CARD" text
    text = "ID CARD"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    text_x = (width - text_width) // 2
    text_y = 20  # Adjust as needed

    draw.text((text_x, text_y), text, fill=text_color, font=font)

    # Underline the text
    underline_y = text_y + text_height + 8
    draw.line((text_x, underline_y, text_x + text_width, underline_y), fill=text_color, width=2)

    # Define dimensions and position for the student's picture
    picture_width, picture_height = 100, 150
    picture_box_x = width - picture_width - 20
    picture_box_y = text_y + text_height + 30

    # Draw a box for the student's picture
    
    draw.rectangle(
        [picture_box_x, picture_box_y, picture_box_x + picture_width, picture_box_y + picture_height],
        outline=border_color,
        width=border_size
    )
    

    # Use the default picture for every student
    student_pic = Image.open("StudentPic.jpg").resize((picture_width, picture_height))
    card.paste(student_pic, (picture_box_x, picture_box_y))

    # Define starting positions for student data
    data_x = 25  # Horizontal position
    data_xx = 107  # Horizontal position
    data_y = 100  # Vertical position

    # Draw the student's name
    font = ImageFont.truetype("arialbd.ttf", 15)  # Use bold Arial or another bold font
    student_name = row['Name']  # Adjust column name for the student's name
    draw.text((data_x, data_y), f"Name", fill=text_color, font=font)
    draw.text((data_xx, data_y), f": {student_name}", fill=text_color, font=font)

    # Draw the student Roll#
    data_y += 20  # Vertical position
    student_id = row['Roll#']  # Adjust column name for the student ID
    draw.text((data_x, data_y), f"Roll No.", fill=text_color, font=font)
    draw.text((data_xx, data_y), f": {student_id}", fill=text_color, font=font)

    # Draw the student Campus
    data_y += 20  # Vertical position
    student_id = row['Campus']  # Adjust column name for the student ID
    draw.text((data_x, data_y), f"Campus", fill=text_color, font=font)
    draw.text((data_xx, data_y), f": {student_id}", fill=text_color, font=font)

     # Draw the student City
    data_y += 20  # Vertical position
    student_id = row['City']  # Adjust column name for the student ID
    draw.text((data_x, data_y), f"City", fill=text_color, font=font)
    draw.text((data_xx, data_y), f": {student_id}", fill=text_color, font=font)

     # Draw the student Day & Time
    data_y += 20  # Vertical position
    student_id = row['Day / Time']  # Adjust column name for the student ID
    draw.text((data_x, data_y), f"Day & Time", fill=text_color, font=font)
    draw.text((data_xx, data_y), f": {student_id}", fill=text_color, font=font)

    # Save the card using the roll number as the filename
    roll_number = row['Roll#']  # Adjust column name for the roll number
    card.save(f"{student_name} {roll_number}.png")

    # Optionally show the card
    card.show()
