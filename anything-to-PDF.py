import os
from PIL import Image
from PyPDF2 import PdfMerger

def compress_image(image_path, max_width=1200):
    """Resize image to reduce size while maintaining aspect ratio."""
    img = Image.open(image_path)
    print(f"Original size of {image_path}: {img.size}")  # Log original size
    if img.width > max_width:
        width_percent = max_width / float(img.width)
        height_size = int((float(img.height) * width_percent))
        img = img.resize((max_width, height_size), Image.LANCZOS)
        print(f"Compressed size of {image_path}: {img.size}")  # Log new size
    return img.convert('RGB')  # Ensure compatibility for PDF

def combine_media_to_pdf(folder_path, output_pdf, sort_by='name', compress=False):
    print("Scanning folder for media files...")
    files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif', '.pdf'))
    ]

    if not files:
        print("No valid media files found in the folder!")
        return

    print(f"Found {len(files)} media files. Sorting by {sort_by}...")
    if sort_by == 'name':
        files.sort()
    elif sort_by == 'date':
        files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)))

    merger = PdfMerger()
    temp_pdf = "temp_images.pdf"

    try:
        image_list = []  # Store images for temporary PDF

        for idx, file in enumerate(files, 1):
            file_path = os.path.join(folder_path, file)
            print(f"Processing file {idx}/{len(files)}: {file}")

            if file.lower().endswith('.pdf'):
                try:
                    merger.append(file_path)
                    print(f"Added PDF: {file}")
                except Exception as e:
                    print(f"Error adding PDF {file}: {e}")
            else:
                try:
                    if compress:
                        img = compress_image(file_path)  # Compress image if enabled
                        temp_image_path = f"temp_{file}.jpg"
                        img.save(temp_image_path, format='JPEG', quality=85)  # Save as JPEG with compression
                        image_list.append(Image.open(temp_image_path))
                        print(f"Compressed and added image: {file}")
                    else:
                        img = Image.open(file_path).convert('RGB')  # Add original image
                        image_list.append(img)
                        print(f"Added original image: {file}")
                except Exception as e:
                    print(f"Error processing image {file}: {e}")

        if image_list:
            print("Saving images to temporary PDF...")
            image_list[0].save(temp_pdf, save_all=True, append_images=image_list[1:])  # Save all images to PDF
            print(f"Saved images to {temp_pdf}.")
            image_list.clear()  # Free memory
            merger.append(temp_pdf)
            print("Temporary PDF created and merged.")

        print(f"Writing final PDF to {output_pdf}...")
        merger.write(output_pdf)
        print(f"PDF successfully saved as {output_pdf}")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        merger.close()
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
                print("Temporary PDF removed.")
            except Exception as e:
                print(f"Error removing temporary PDF: {e}")

# Usage example
if __name__ == "__main__":
    folder = input("Enter the folder path containing images and PDFs: ").strip()
    output_file = input("Enter the full path and filename for the output PDF (e.g., 'C:\\Users\\YourName\\Desktop\\output.pdf'): ").strip()
    sort_option = input("Sort by 'name' or 'date'? ").strip().lower()
    compress_option = input("Do you want to compress images? (yes/no): ").strip().lower()

    if sort_option not in ['name', 'date']:
        print("Invalid sort option. Defaulting to 'name'.")
        sort_option = 'name'

    compress_images = compress_option == 'yes'
    combine_media_to_pdf(folder, output_file, sort_by=sort_option, compress=compress_images)
