import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from pdf2docx import parse
import time



root = Tk()

window_width = 400
window_hight = 400

monitor_width = root.winfo_screenwidth()
monitor_hight = root.winfo_screenheight()

x = (monitor_width / 2) - (window_width / 2)
y = (monitor_hight / 2) - (window_hight / 2)

root.geometry(f'{window_width}x{window_hight}+{int(x)}+{int(y)}')
root.iconbitmap("JK.ico")
root.resizable(False, False)
root.config(bg="#dbdbdb")
root.title("PDF to DOCX Converter")


def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("PDF File", "*.pdf")])
    file_path_label.config(width=35, text=file_path)

def select_save_path():
    global save_path
    save_path = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("DOCX Files", "*.docx")])
    save_path_label.config(width=35, text=save_path)

def reset():
    messagebox.showinfo("Info", "File conversion successful!")
    time.sleep(3)
    root.title("PDF to DOCX Converter")
    result_label.config(text="")
    file_path_label.config(text="")
    save_path_label.config(text="")

def convert():
    global duration
    if file_path.endswith(".pdf"):
        if save_path:
            root.title("Converting File...")
            save_path = os.path.splitext(save_path)[0] + ".docx "
            parse(file_path, save_path)
            result_label.config(text="Conversion successful!")
            reset()
        else:
            result_label.config(text="Please select a save path")
    else:
        result_label.config(text="Please select a PDF file")


file_path_label = Label(root, font=("arial", 12), bg="#dbdbdb", text="")
file_path_label.pack(pady=5)

file_select_button = Button(root, text="Select PDF File", font=("arial", 12), width=20, command=select_file)
file_select_button.pack(pady=0)

save_path_label = Label(root, font=("arial", 12), bg="#dbdbdb", text="")
save_path_label.pack(pady=0)

save_path_button = Button(root, text="Select Save Path", font=("arial", 12), width=20, command=select_save_path)
save_path_button.pack(pady=5)

result_label = Label(root, font=("arial", 12), bg="#dbdbdb", text="")
result_label.pack()

convert_button = Button(root, font=("arial", 12), width=20, text="Convert to DOCX", command=convert)
convert_button.pack(pady=30)

root.mainloop()
