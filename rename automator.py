"""
TO-DO: 
    - add rename worksheet function to Settings 
    - test state of labels (colour & text)
        - check last few lines of readPdf and readExcel functions (search # here)

"""




import customtkinter
import os
import sys
from natsort import natsorted
import pandas as pd
from tkinter import filedialog
from PIL import Image

customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

# uh, just image importing thingy
def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Centers the window on the screen
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

app = customtkinter.CTk()  
app.title("Rename Automator")
center_window(app, 400, 200)

# global variable declaration
excelFilePath = ""
pdfFolderPath = ""
removeNumbers = "Remove Numbers"
excelWithHeader = "Without Header"

def defaultLabel0():
    label0.configure(bg_color="transparent")

def canDisplaySubmitButton():
    if excelFilePath != "" and pdfFolderPath != "":
        submitButton.place(relx=0.95, rely=0.95, anchor=customtkinter.SE)
    else:
        submitButton.place_forget()

def readExcelFile():
    global excelFilePath, pdfFolderPath

    # Open the file dialog to select a file
    excelFilePath = filedialog.askopenfilename()

    # input validation for excel file 
    if excelFilePath:
        if (excelFilePath[-5:] != ".xlsx"):
            label1.configure(text="Given file is not in .xlsx format", bg_color="red") 
        else:
            file_name = os.path.basename(excelFilePath)
            label1.configure(text=f"File selected: {file_name}", bg_color="transparent")
    else:
        label1.configure(text="No file selected", bg_color="transparent")

    # dont remove, lazy to explain, fuck around and find out
    defaultLabel0()

    # here 
    if pdfFolderPath:
        folder_name = os.path.basename(pdfFolderPath)
        label2.configure(text=f"Folder selected: {folder_name}")
    else:
        label2.configure(text="No folder selected", bg_color="transparent")
    
    canDisplaySubmitButton()

def readPdfFolder():
    global pdfFolderPath

    pdfFolderPath = filedialog.askdirectory()

    # input validation for pdf folder 
    if pdfFolderPath:
        folder_name = os.path.basename(pdfFolderPath)
        label2.configure(text=f"Folder selected: {folder_name}", bg_color="transparent")
    else:
        label2.configure(text="No folder selected", bg_color="transparent")

    # dont remove, lazy to explain, fuck around and find out
    defaultLabel0()

    # here
    if excelFilePath:
        file_name = os.path.basename(excelFilePath)
        label1.configure(text=f"File selected: {file_name}")
    else:
        label1.configure(text="No file selected", bg_color="transparent")

    canDisplaySubmitButton()

def submit():
    global excelFilePath, removeNumbers, excelWithHeader

    print(f"{removeNumbers}, {excelWithHeader}")

    # show progress
    label1.configure(text="Renaming the files...", bg_color="yellow")
    label2.configure(text="", bg_color="transparent")

    # read excel file
    try:
        if (excelWithHeader == "With Header"):
    if (excelWithHeader == "With Header"):
        ambassador = pd.read_excel(excelFilePath, "Ambassador", usecols="A")
    elif (excelWithHeader == "Without Header"):
        ambassador = pd.read_excel(excelFilePath, "Ambassador", usecols="A", header=None)
        elif (excelWithHeader == "Without Header"):
    except ValueError:
        label1.configure(text="Worksheet named 'Ambassador' not found", bg_color="red")
        label2.configure(text="", bg_color="transparent")

    # read and sort the pdf files within the folder
    files = os.listdir(pdfFolderPath)
    sorted_files = natsorted(files)

    # double check for number of files and number of records in excel
    if (len(ambassador) != len(files)):
        print("\n\n\n\033[31m" + "ERROR!! ooops number of records in excel not equal to the number of files in ilovepdf" + "\033[0m")
        print("\033[31m" + "Number of records in excel: \033[0;35m" + str(len(ambassador)) + "\033[0m")
        print("\033[31m" + "Number of files in PDF folder: \033[0;35m" + str(len(files)) + "\033[0m")

        label1.configure(text="Number of files in PDF folder and number", bg_color="red")
        label2.configure(text="of records in excel is not equal!", bg_color="red")

        raise Exception("Number of excel rows != number of files in folder")
    
    # renaming of pdf files
    for index, old_file_name in enumerate(sorted_files):
    # very inefficient but in all honesty
    # idc
        if (removeNumbers == "Remove Numbers"):
            new_file_name = ambassador.iat[index, 0] + '.pdf'
        elif (removeNumbers == "Don't Remove Numbers"):
            new_file_name = str(index+1) + ' - ' + ambassador.iat[index, 0] + '.pdf'

        # string formatting (etc. a/p, a/l, s/o, d/o)
        # when / is found, remove everything except for the first name 
        otherIndex = new_file_name.find('/')
        if otherIndex != -1:
            new_file_name = new_file_name[:otherIndex-2].strip() + '.pdf'

        old = os.path.join(pdfFolderPath + "/" + old_file_name)
        new = os.path.join(pdfFolderPath + "/" + new_file_name)

        if os.path.exists(old):
            print(new)

            if os.path.exists(new):
                new = new[0:-4] + " (" +  str(index+1) +  ")" + new[-4:]
                os.rename(old, new)
            else:
                os.rename(old, new)

    label0.configure(bg_color="green")
    label1.configure(text="Success! Check the PDF folder.", bg_color="green")
    label2.configure(text="", bg_color="green")

def settingsPage():
    global settingsPage, removeNumbers, excelWithHeader
    settingsPage = customtkinter.CTkToplevel()
    settingsPage.title("Settings")
    center_window(settingsPage, 410, 300)
    settingsPage.after(100, settingsPage.lift)

    def updateValues():
        global removeNumbers, excelWithHeader
        removeNumbers = combo1.get()
        excelWithHeader = combo2.get()
        print(removeNumbers)
        print(excelWithHeader)
        settingsPage.destroy()

    # remove numbers from renamed files?
    label1 = customtkinter.CTkLabel(settingsPage, text="Would you like to remove numbers from the renamed files?")
    label1.pack(pady=0)
    combo1 = customtkinter.CTkComboBox(settingsPage, values=["Remove Numbers", "Don't Remove Numbers"], state="readonly")
    combo1.pack(pady=1)
    combo1.set(removeNumbers)

    # excel with header?
    label2 = customtkinter.CTkLabel(settingsPage, text="Does your excel file has a header?")
    label2.pack(pady=0)
    combo2 = customtkinter.CTkComboBox(settingsPage, values=["With Header", "Without Header"], state="readonly")
    combo2.pack(pady=1)
    combo2.set(excelWithHeader)

    # Save Button
    saveButton = customtkinter.CTkButton(settingsPage, text="Save", command=updateValues)
    saveButton.pack(pady=10)

    # Back Button
    backButton = customtkinter.CTkButton(settingsPage, text="Back", command=settingsPage.destroy)
    backButton.pack(pady=0)

##########.
## here ##
##########
uploadButton1 = customtkinter.CTkButton(master=app, text="Click to upload excel file", command=readExcelFile)
uploadButton1.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)

uploadButton2 = customtkinter.CTkButton(master=app, text="Select folder with the pdf files.", command=readPdfFolder)
uploadButton2.place(relx=0.5, rely=0.7, anchor=customtkinter.CENTER)

submitButton = customtkinter.CTkButton(master=app, text="Submit", command=submit, width=80)
submitButton.place_forget()

settingImg = customtkinter.CTkImage(dark_image=Image.open(resource_path("resources\\settings-dark.png")), light_image=Image.open(resource_path("resources\\settings-light.png")), size=(30,30))
settingsButton = customtkinter.CTkButton(master=app, image=settingImg, text="", command=settingsPage, width=30, height=30, fg_color="transparent")
settingsButton.place(relx=0.99, rely=0.05, anchor=customtkinter.NE)

# grid configuration
app.grid_rowconfigure(0, weight=1)  # Stretch row 0
app.grid_rowconfigure(1, weight=1)  # Stretch row 1

# Add a label to display the selected file path
label0 = customtkinter.CTkLabel(app, text="", width=300, height=20)
label0.place(relx=0.5, rely=0.10, anchor=customtkinter.CENTER)
label1 = customtkinter.CTkLabel(app, text="No file selected", width=300, height=20)
label1.place(relx=0.5, rely=0.20, anchor=customtkinter.CENTER)
label2 = customtkinter.CTkLabel(app, text="No folder selected", width=300, height=20)
label2.place(relx=0.5, rely=0.30, anchor=customtkinter.CENTER)

app.mainloop()