import os
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import sys

print("Python version:", sys.version)

folder_path = r"D:\JKD_Folder\JKD-PROJECT SITE\JKD_SITE_PHOTOS\WPR_09-2-26\28-2-26"
photos = [f for f in os.listdir(folder_path) if f.lower().endswith('.jpeg')]

if not photos:
    print("No photos found!")
    sys.exit()

photo_name = photos[0]
photo_path = os.path.join(folder_path, photo_name)
print(f"Testing with photo: {photo_path}")

try:
    img = OpenpyxlImage(photo_path)
    print("Successfully created OpenpyxlImage")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.add_image(img, "A1")
    wb.save("test_output.xlsx")
    print("Successfully saved test_output.xlsx")
except Exception as e:
    print(f"Error occurred: {type(e).__name__}: {str(e)}")
    import traceback
    traceback.print_exc()
