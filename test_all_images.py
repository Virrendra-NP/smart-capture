import os
import traceback
import sys

def test():
    try:
        import openpyxl
        from openpyxl.drawing.image import Image as OpenpyxlImage
    except Exception as e:
        print("Import error:", str(e))
        return

    folder_path = r"D:\JKD_Folder\JKD-PROJECT SITE\JKD_SITE_PHOTOS\WPR_09-2-26\28-2-26"
    valid_extensions = ('.jpg', '.jpeg', '.png', '.bmp')
    try:
        photos = [f for f in os.listdir(folder_path) if f.lower().endswith(valid_extensions)]
    except Exception as e:
        print("OS listdir error:", str(e))
        return

    if not photos:
        print("No valid photos found")
        return

    print(f"Found {len(photos)} photos. Testing OpenpyxlImage...")
    errors = 0
    for photo_name in photos:
        photo_path = os.path.join(folder_path, photo_name)
        try:
            img = OpenpyxlImage(photo_path)
            # simulate adding parameters
            img.width = 454
            img.height = 680
        except Exception as e:
            print(f"Error on {photo_name}: {type(e).__name__} - {str(e)}")
            errors += 1
            
    print(f"Finished. {errors} errors out of {len(photos)} photos.")

if __name__ == '__main__':
    test()
