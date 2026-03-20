import os
import xlsxwriter

excel_path = "test_xlsxwriter.xlsx"
folder_path = r"D:\JKD_Folder\JKD-PROJECT SITE\JKD_SITE_PHOTOS\WPR_09-2-26\28-2-26"

photos = [f for f in os.listdir(folder_path) if f.lower().endswith('.jpeg')][:3]

workbook = xlsxwriter.Workbook(excel_path)
worksheet = workbook.add_worksheet('Photos')

worksheet.write('A1', 'Photo Name')
worksheet.write('B1', 'Image (12cm x 18cm)')

worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 66)

row_idx = 1 # 0-indexed in xlsxwriter

for photo_name in photos:
    photo_path = os.path.join(folder_path, photo_name)
    
    # 18cm ~ 510 points
    # Max row height in excel is 409 points. So we still use 2 rows.
    worksheet.set_row(row_idx, 255)
    worksheet.set_row(row_idx+1, 255)
    
    # Merge cells for text
    worksheet.merge_range(row_idx, 0, row_idx+1, 0, photo_name)
    
    # Insert image
    # We want 12cm x 18cm. Default DPI is typically 96, but xlsxwriter uses scale.
    # 12cm = 4.72 inches = 454 pixels
    # 18cm = 7.08 inches = 680 pixels
    
    # Provide the option with position
    worksheet.insert_image(row_idx, 1, photo_path, {'width': 454, 'height': 680})
    
    row_idx += 2

workbook.close()
print(f"Created {excel_path} successfully")
