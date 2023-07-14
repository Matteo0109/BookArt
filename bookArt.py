#This program will help you to do book strip art
#To do that you need to have python3 and PIL installed
#import excel library
import openpyxl
from openpyxl.drawing.image import Image
import customtkinter
from tkinter import filedialog as fd
from tkinter.filedialog import asksaveasfile

#Let's create a function that will take as argument the number of pages of the book and the image that we want to use
def book_strip_art(pages, h_book, image_path):
    """
    This function will take as argument the number of pages, the height of the book and the image that we want to use
    and then will create an excel document with all the page on the top of the sheet et below the image resized to the
    height of the book

    :param pages: number of pages of the book
    :param h_book: height of the book
    :param image: image that we want to use

    :return: excel document with the image resized to the height of the book and the number of pages on the top of the
    sheet
    """
    #Create the excel document
    wb = openpyxl.Workbook()
    #Let's create the sheet
    sheet = wb.active
    #all column and row must be 0,5 inches
    #sheet.column_dimensions["A"].width = 5

    i=1
    l_number = []
    while i <= pages:
        l_number.append(i)
        #sheet.cell(row=1, column=i).border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                                                                     #right=openpyxl.styles.Side(style='thin'))
        i += 2

    sheet.append(l_number)


    #Set the column width to 0,5 inches
    for row in sheet.rows:
        #print(row)
        for cell in row:
            if cell.value:
                sheet.column_dimensions[cell.column_letter].width = 5

    #Print the resized image on the excel document below the page number, the image must fit the end of the column cell
    img = Image(image_path)

    #pixels = cms * dpi / 2.54
    print(h_book)
    img.height = h_book * 96
    print(img.height)
    img.width = (pages/4) * 96 
    sheet.add_image(img,'A2')

    #Let's save the excel document
    file_path=fd.asksaveasfilename( title="Select ",defaultextension=".*", filetypes=[('Excel File', ['.xlsx'])])
    print(file_path)
    if file_path:
        wb.save(file_path)

#Let's call the function
#book_strip_art(286, 9, '/Users/admin/Downloads/IMG_3658.jpeg')

app = customtkinter.CTk()

def open_file():
    file_path = fd.askopenfilename(filetypes=[('Image Files', ['.jpeg', '.jpg', '.png', '.gif','.tiff', '.tif', '.bmp'])])
    if file_path:
        print(file_path)
        book_strip_art(int(pages_nb.get()), int(h_book_nb.get()), file_path)
    
app.geometry("500x500")
app.title("Book Strip Art")
pages_l = customtkinter.CTkLabel(app,text="Nombre de pages du livre :")
pages_l.pack()
pages_nb = customtkinter.CTkEntry(app,placeholder_text="Insérer ici le nombre de pages du livre")
pages_nb.pack()
h_book_l = customtkinter.CTkLabel(app,text="Hauteur du livre :")
h_book_l.pack()
h_book_nb = customtkinter.CTkEntry(app,placeholder_text="Insérer ici la hauteur du livre")
h_book_nb.pack()
open_btn = customtkinter.CTkButton(app,text="Ouvrir une image",command=open_file)
print(open_btn)
open_btn.pack(pady=10)


app.mainloop()
