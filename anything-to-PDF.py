import os
from PIL import Image
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time

def compress_image(image_path, max_width=1200):
    """Resize image to reduce size while maintaining aspect ratio."""
    img = Image.open(image_path)
    if img.width > max_width:
        width_percent = max_width / float(img.width)
        height_size = int((float(img.height) * width_percent))
        img = img.resize((max_width, height_size), Image.LANCZOS)
    return img.convert('RGB')

def log_message(message):
    """Log messages to the text widget and auto-scroll."""
    status_text.configure(state=tk.NORMAL)
    status_text.insert(tk.END, message + "\n")
    status_text.see(tk.END)
    status_text.configure(state=tk.DISABLED)

def combine_media_to_pdf(folder_path, output_pdf, sort_by='name', order='ascending', compress=False):
    if not os.path.exists(folder_path):
        log_message("Error: Folder does not exist.")
        return

    files = []
    for root, _, filenames in os.walk(folder_path):
        for filename in filenames:
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif', '.pdf')):
                files.append(os.path.join(root, filename))

    if not files:
        log_message("Info: No valid media files found in the folder!")
        return

    # Sort files based on user's choice
    reverse_order = (order == 'descending')
    if sort_by == 'name':
        files.sort(reverse=reverse_order)
    elif sort_by == 'date':
        files.sort(key=lambda f: os.path.getmtime(f), reverse=reverse_order)

    merger = PdfMerger()
    temp_pdf = "temp_images.pdf"
    start_time = time.time()

    try:
        image_list = []

        for idx, file in enumerate(files, 1):
            log_message(f"Processing file {idx}/{len(files)}: {file}")

            if file.lower().endswith('.pdf'):
                try:
                    merger.append(file)
                    log_message(f"Added PDF: {file}")
                except Exception as e:
                    log_message(f"Error adding PDF {file}: {e}")
            else:
                try:
                    if compress:
                        img = compress_image(file)
                        log_message(f"Compressed and added image: {file}")
                    else:
                        img = Image.open(file).convert('RGB')
                        log_message(f"Added original image: {file}")
                    image_list.append(img)
                except Exception as e:
                    log_message(f"Error processing image {file}: {e}")

            elapsed_time = time.time() - start_time
            if idx > 0:
                estimated_time_remaining = (elapsed_time / idx) * (len(files) - idx)
                eta_label.config(text=f"ETA: {estimated_time_remaining:.2f} seconds")
                progress = (idx / len(files)) * 100
                progress_bar['value'] = progress

        if image_list:
            log_message("Saving images to temporary PDF...")
            image_list[0].save(temp_pdf, save_all=True, append_images=image_list[1:])
            merger.append(temp_pdf)

        merger.write(output_pdf)
        log_message(f"PDF successfully saved as {output_pdf}")

    except Exception as e:
        log_message(f"An error occurred: {e}")

    finally:
        merger.close()
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except Exception as e:
                log_message(f"Error removing temporary PDF: {e}")

        total_duration = time.time() - start_time
        log_message(f"Total time taken: {total_duration:.2f} seconds")
        progress_bar['value'] = 100

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)

def browse_output_file():
    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                               filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_path)

def create_pdf():
    folder_path = folder_entry.get()
    output_file = output_entry.get()
    sort_option = sort_var.get()
    order_option = order_var.get()
    compress_option = compress_var.get()

    compress_images = (compress_option == "yes")
    threading.Thread(target=combine_media_to_pdf, 
                     args=(folder_path, output_file, sort_option, order_option, compress_images), 
                     daemon=True).start()

# Create the main window
root = tk.Tk()
root.title("Media to PDF Combiner")

# Create and place labels and entries
tk.Label(root, text="Folder Path:").grid(row=0, column=0, padx=10, pady=10)
folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Button(root, text="Browse", command=browse_folder).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Output PDF Path:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="Browse", command=browse_output_file).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Sort By:").grid(row=2, column=0, padx=10, pady=10)
sort_var = tk.StringVar(value="name")
ttk.Combobox(root, textvariable=sort_var, values=["name", "date"]).grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Order:").grid(row=3, column=0, padx=10, pady=10)
order_var = tk.StringVar(value="ascending")
ttk.Combobox(root, textvariable=order_var, values=["ascending", "descending"]).grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Compress Images:").grid(row=4, column=0, padx=10, pady=10)
compress_var = tk.StringVar(value="yes")
ttk.Combobox(root, textvariable=compress_var, values=["yes", "no"]).grid(row=4, column=1, padx=10, pady=10)

tk.Button(root, text="Create PDF", command=create_pdf).grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# Status updates
status_text = tk.Text(root, height=10, width=70, state=tk.DISABLED)
status_text.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

# ETA label
eta_label = tk.Label(root, text="", width=70)
eta_label.grid(row=7, column=0, columnspan=3, padx=10, pady=5)

# Progress bar
progress_bar = ttk.Progressbar(root, length=600, mode='determinate')
progress_bar.grid(row=8, column=0, columnspan=3, padx=10, pady=10)

# Start the Tkinter event loop
root.mainloop()
