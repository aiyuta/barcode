#selete a zip file, or cancel the select window to display another window to select a folder which include pdf files.the program will run and generate a zip file included excel files 
import fitz
import io
from PIL import Image
from pyzbar.pyzbar import decode
import pandas as pd
from tkinter import filedialog
import tkinter as tk
import os
import zipfile
import time


total_time = 0

def calculate_time(func):
    def wrapper(*args, **kwargs):
        global total_time
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        elapsed_time = end_time - start_time
        total_time += elapsed_time
        print("Elapsed time for {}: {:.2f} seconds".format(func.__name__, elapsed_time))
        return result
    return wrapper


def unzip_file(file_path):
    
    # create the output folder
    folder_path = os.path.splitext(file_path)[0]
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # extract the contents of the zip file
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(folder_path)

    return folder_path

@calculate_time
# this function extract the images included barcod pics from the pdf, then decode adn saved to a excel file. 
def ex_de(folder_path):
    # loop through all the pdf files
    for root, dirs, files in os.walk(folder_path):
        
        for file in files:
            if file.endswith('.pdf'):
                file_path = os.path.join(root, file)
                file_save = file.split('.')[0]
                
                #open the file
                pdf_file = fitz.open(file_path)   
                results=[]
                #iterate over PDF pages
                for page_index in range(pdf_file.page_count):
                    #get the page itself
                    page = pdf_file[page_index]
                    image_li = page.get_images()
                    #printing number of images found in this page
                    #page index starts from 0 hence adding 1 to its content
                    if image_li:
                        print(f"[+] Found a total of {len(image_li)} images in page {page_index+1}")
                    else:
                        print(f"[!] No images found on page {page_index+1}")
                    for image_index, img in enumerate(page.get_images(), start=1):
                        #get the XREF of the image
                        xref = img[0]
                        #extract the image bytes
                        base_image = pdf_file.extract_image(xref)
                        image_bytes = base_image["image"]
                        #get the image extension
                        image_ext = base_image["ext"]
                        if image_index == 3:
                            #load it to PIL
                            image = Image.open(io.BytesIO(image_bytes))
                            # Recognize barcodes in the image using pyzbar
                            dc = decode(image)
                            s = dc[0].data.decode('utf-8')
                            result = "(" + "00" + ")" + s[1:4] + "0" + s[4:7] + s[7:10] + s[10:]
                            print(result)
                            results.append(result)
                            #save it to local disk        
                            #image.save(open(f"image{page_index+1}_{image_index}.{image_ext}", "wb"))

            df = pd.DataFrame({'Bar code': results})
            df_sorted = df.sort_values('Bar code')
            #df_sorted.to_excel(f'{file_save}.xlsx', index=False)
            df_sorted.to_excel(os.path.join(root, f'{file_save}.xlsx'), index=False)
            print(f'{file_save}.xlsx has saved')

       
def zip_files(folder_path):
    # zip the Excel files in the folder_path
    zip_path = os.path.join(folder_path, f'{folder_path}.zip')
    with zipfile.ZipFile(zip_path, mode='w') as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.xlsx'):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))
                    os.remove(file_path)
    print(f'Excel files zipped to {zip_path}.')            
    

def file_or_folder_path():
    # Create the tkinter window
    root = tk.Tk()
    root.withdraw()

    # Ask the user to select a file or folder
    file_path = filedialog.askopenfilename(filetypes=[("Zip Files", "*.zip")])
    if not file_path:
        folder_path = filedialog.askdirectory()
    else :
        folder_path= unzip_file(file_path)
    return folder_path

if __name__ == '__main__':
    
    folder_path = file_or_folder_path()

    # process the contents of the folder
    ex_de(folder_path)
    zip_files(folder_path)

    print('The processing is done') 
    