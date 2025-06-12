from PIL import Image, ImageDraw, ImageFont

# Default parameters
DEFAULT_WATERMARK_COLOR = "#B3B3B3"  # Use a single color
DEFAULT_TRANSPARENCY = 128  # Set transparency
DEFAULT_Y_POSITION = 600  # Set an adjustable Y position for the watermark

def add_name_watermark(image, name, font, color=DEFAULT_WATERMARK_COLOR, transparency=DEFAULT_TRANSPARENCY, y_position=DEFAULT_Y_POSITION):
    # Create a new image for the text with transparency support
    text_image = Image.new('RGBA', image.size, (0, 0, 0, 0))
    text_draw = ImageDraw.Draw(text_image)

    # Calculate the width of the text to position it at the right edge
    text_width, text_height = text_draw.textbbox((0, 0), name, font=font)[2:]

    # Set the X position to align the text to the right edge without overflow
    x_position = max(0, image.width - text_width - 30)  # 20 pixels padding from the right edge

    # Apply color with transparency
    color_with_transparency = color + f"{transparency:02x}"

    # Draw the watermark text at the calculated X position and adjustable Y position
    text_draw.text((x_position, y_position), name, font=font, fill=color_with_transparency)

    # Rotate the watermark by 45 degrees
    rotated_text = text_image.rotate(0, expand=1)

    # Paste the watermark onto the original image
    image.paste(rotated_text, (0, 0), rotated_text)
