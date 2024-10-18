import os
from PIL import Image
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

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
        return "No valid media files found in the folder!"

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
                    return f"Error adding PDF {file}: {e}"
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
                    return f"Error processing image {file}: {e}"

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
        return f"An error occurred: {e}"

    finally:
        merger.close()
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
                print("Temporary PDF removed.")
            except Exception as e:
                print(f"Error removing temporary PDF: {e}")

def run_combination():
    folder = folder_entry.get()
    output_file = output_entry.get()
    sort_option = sort_var.get()
    compress_images = compress_var.get()

    status_text.delete(1.0, tk.END)  # Clear previous status

    if not folder or not output_file:
        messagebox.showwarning("Input Error", "Please specify both folder and output file paths.")
        return

    # Perform the combination and compression
    status_text.insert(tk.END, "Combining media files...\n")
    result = combine_media_to_pdf(folder, output_file, sort_by=sort_option, compress=compress_images)

    if result:
        status_text.insert(tk.END, f"Status: {result}\n")
    else:
        status_text.insert(tk.END, "Completed successfully!\n")

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)

def browse_output():
    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_path)

# Create the GUI
root = tk.Tk()
root.title("Media to PDF Combiner")

# Folder Selection
folder_label = tk.Label(root, text="Folder Path:")
folder_label.grid(row=0, column=0, padx=5, pady=5)
folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=0, column=1, padx=5, pady=5)
folder_button = tk.Button(root, text="Browse", command=browse_folder)
folder_button.grid(row=0, column=2, padx=5, pady=5)

# Output File Selection
output_label = tk.Label(root, text="Output PDF Path:")
output_label.grid(row=1, column=0, padx=5, pady=5)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)
output_button = tk.Button(root, text="Browse", command=browse_output)
output_button.grid(row=1, column=2, padx=5, pady=5)

# Sorting Options
sort_label = tk.Label(root, text="Sort By:")
sort_label.grid(row=2, column=0, padx=5, pady=5)
sort_var = tk.StringVar(value="name")
sort_name_radio = tk.Radiobutton(root, text="Name", variable=sort_var, value="name")
sort_name_radio.grid(row=2, column=1, padx=5, pady=5, sticky='w')
sort_date_radio = tk.Radiobutton(root, text="Date", variable=sort_var, value="date")
sort_date_radio.grid(row=2, column=1, padx=5, pady=5)

# Compression Option
compress_var = tk.BooleanVar()
compress_check = tk.Checkbutton(root, text="Compress Images", variable=compress_var)
compress_check.grid(row=3, columnspan=3, padx=5, pady=5)

# Run Button
run_button = tk.Button(root, text="Combine to PDF", command=run_combination)
run_button.grid(row=4, columnspan=3, padx=5, pady=5)

# Status Text Area
status_text = scrolledtext.ScrolledText(root, width=60, height=15)
status_text.grid(row=5, columnspan=3, padx=5, pady=5)

# Start the GUI loop
root.mainloop()
