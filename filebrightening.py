import io
import sys
import zipfile
from PIL import Image


#Checks is image bright or not
def color_check(img, pixels):
    bright_pixels = 0
    for i in range(img.size[0]):
        for j in range(img.size[1]):
            x, y, z = pixels[i, j][0], pixels[i, j][1], pixels[i, j][2]
            if x > 90 and y > 90 and z > 90:
                bright_pixels += 1

    if bright_pixels < (img.size[0] * img.size[1]) - bright_pixels:
        return True

    return False


def inverter(img, pixels):
    for i in range(img.size[0]):
        for j in range(img.size[1]):
            x, y, z = pixels[i, j][0], pixels[i, j][1], pixels[i, j][2]
            x, y, z = abs(x - 255), abs(y - 255), abs(z-255)
            pixels[i, j] = (x, y, z)

    return pixels


#Fetch images from docx file
def image_inverter(filename):
    inverted_file = filename.replace('.docx', '_inverted.docx')
    with zipfile.ZipFile(filename) as inzip:
        with zipfile.ZipFile(inverted_file, 'w') as outzip:
            for info in inzip.infolist():
                name = info.filename
                content = inzip.read(info)
                if info.filename.endswith(('.png', '.jpeg', '.gif')):
                    print(name)
                    file_extension = name.split('.')[-1]
                    img = Image.open(io.BytesIO(content))
                    pixels = img.load()

                    if color_check(img, pixels):
                        pixels = inverter(img, pixels)

                    outb = io.BytesIO()
                    img.save(outb, file_extension)
                    content = outb.getvalue()
                    info.file_size = len(content)
                    info.CRC = zipfile.crc32(content)
                outzip.writestr(info, content)


image_inverter(sys.argv[1])
