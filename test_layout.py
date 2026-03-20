import os
import xlsxwriter
from PIL import Image

excel_path = "test_layout.xlsx"
folder_path = r"D:\JKD_Folder\JKD-PROJECT SITE\JKD_SITE_PHOTOS\WPR_09-2-26\28-2-26"
photos = [f for f in os.listdir(folder_path) if f.lower().endswith('.jpeg')][:4]

workbook = xlsxwriter.Workbook(excel_path)
worksheet = workbook.add_worksheet('Photos')

worksheet.set_column('A:B', 48)
bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

TARGET_WIDTH = 321
TARGET_HEIGHT = 397

row_idx = 0
for i in range(0, len(photos), 2):
    photo_row_idx = row_idx + 2
    worksheet.set_row(row_idx, 20)      
    worksheet.set_row(row_idx + 1, 20)  
    worksheet.set_row(photo_row_idx, 305) 
    
    for j in range(2):
        if i + j < len(photos):
            photo_name = photos[i+j]
            photo_path = os.path.join(folder_path, photo_name)
            col_idx = j
            
            worksheet.write(row_idx, col_idx, f"Name: {photo_name}", bold_format)
            worksheet.write(row_idx + 1, col_idx, "Date: ", center_format)
            
            with Image.open(photo_path) as img:
                orig_width, orig_height = img.size
                x_scale = TARGET_WIDTH / orig_width
                y_scale = TARGET_HEIGHT / orig_height

            worksheet.insert_image(photo_row_idx, col_idx, photo_path, {
                'x_scale': x_scale,
                'y_scale': y_scale,
                'object_position': 1,
                'x_offset': 5,
                'y_offset': 5
            })
            
    row_idx += 3 # skip the photo row, go to next block

workbook.close()
print("Saved test_layout.xlsx")
