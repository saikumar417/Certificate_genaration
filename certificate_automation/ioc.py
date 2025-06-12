from PIL import Image

# Load the PNG image
img = Image.open("SAA_logo.png")

# Resize (optional): icons are usually 256x256, 128x128, etc.
img = img.resize((2304, 2304))

#img.save("your_icon.ico", format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])

# Save as .ico
img.save("SAA_logo.ico", format='ICO')
