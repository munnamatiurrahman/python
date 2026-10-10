import os
import re
from PIL import Image

def get_sort_key(filename):
    """
    Extracts the first sequence of numbers from a filename to use for sorting.
    If '1-abc.jpg' -> returns 1.
    If 'image10.png' -> returns 10.
    """
    numbers = re.findall(r'\d+', filename)
    return int(numbers[0]) if numbers else 999999

def create_pdf_from_folder():
    # Current folder path where the script is run
    folder_path = os.getcwd()
    folder_name = os.path.basename(folder_path)
    output_pdf = os.path.join(folder_path, f"{folder_name}.pdf")

    # Common image extensions
    image_extensions = ('.jpg', '.jpeg', '.png', '.webp', '.bmp')
    
    # Get all image files and filter out the output PDF name if it already exists
    files = [f for f in os.listdir(folder_path) 
             if f.lower().endswith(image_extensions) and f != f"{folder_name}.pdf"]
    
    if not files:
        print(f"No images found in: {folder_path}")
        return

    print(f"Found {len(files)} images. Sorting...")
    
    # Sort files numerically
    files.sort(key=get_sort_key)
    
    images_list = []
    for file in files:
        img_path = os.path.join(folder_path, file)
        try:
            img = Image.open(img_path)
            # Convert to RGB (required for PDF saving)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            images_list.append(img)
            print(f"Added: {file}")
        except Exception as e:
            print(f"Could not process {file}: {e}")

    if images_list:
        print(f"Saving PDF to: {output_pdf}")
        images_list[0].save(output_pdf, save_all=True, append_images=images_list[1:])
        print("\nSuccessfully created PDF!")
    else:
        print("No valid images could be processed.")

if __name__ == "__main__":
    create_pdf_from_folder()
    # Keep window open if run by double-clicking
    input("\nPress Enter to close...")
