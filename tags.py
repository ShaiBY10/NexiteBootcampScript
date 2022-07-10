import tkinter as tk
import tkinter.font as tkFont
from tkinter import *
from algo import *

class App:
    def __init__(self, root):
        #setting title
        self.input_file_path = None
        self.save_xlsx_path = None
        self.input_file = None
        root.title("Bootcamp Script")
        #setting window size
        width=500
        height=329
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        title=tk.Label(root)
        ft = tkFont.Font(family='Times',size=18)
        title["font"] = ft
        title["fg"] = "#333333"
        title["justify"] = "center"
        title["text"] = "Bootcamp Matrice Creator"
        title.place(x=40, y=10, width=267, height=52)

        input_browse_button=tk.Button(root)
        input_browse_button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        input_browse_button["font"] = ft
        input_browse_button["fg"] = "#000000"
        input_browse_button["justify"] = "center"
        input_browse_button["text"] = "Browse..."
        input_browse_button.place(x=90, y=100, width=130, height=30)
        input_browse_button["command"] = self.inputFileBrowseCommand

        self.input_file_message=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        self.input_file_message["font"] = ft
        self.input_file_message["fg"] = "#333333"
        self.input_file_message["justify"] = "center"
        self.input_file_message["text"] = "Waiting for action..."
        self.input_file_message["relief"] = "sunken"
        self.input_file_message.place(x=0, y=140, width=500, height=32)

        input_file_label=tk.Label(root)
        ft = tkFont.Font(family='Times',size=13)
        input_file_label["font"] = ft
        input_file_label["fg"] = "#333333"
        input_file_label["justify"] = "center"
        input_file_label["text"] = "Input File:"
        input_file_label.place(x=10, y=100, width=70, height=25)

        save_path_label=tk.Label(root)
        ft = tkFont.Font(family='Times',size=13)
        save_path_label["font"] = ft
        save_path_label["fg"] = "#333333"
        save_path_label["justify"] = "center"
        save_path_label["text"] = "Save Path:"
        save_path_label.place(x=10, y=190, width=76, height=30)

        browse_save_button=tk.Button(root)
        browse_save_button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        browse_save_button["font"] = ft
        browse_save_button["fg"] = "#000000"
        browse_save_button["justify"] = "center"
        browse_save_button["text"] = "Select Save Path"
        browse_save_button.place(x=90, y=190, width=130, height=30)
        browse_save_button["command"] = self.browseSavePath

        self.save_path_message=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        self.save_path_message["font"] = ft
        self.save_path_message["fg"] = "#333333"
        self.save_path_message["justify"] = "center"
        self.save_path_message["text"] = "Waiting for action..."
        self.save_path_message["relief"] = "sunken"
        self.save_path_message.place(x=0, y=240, width=500, height=30)

        create_file_button = tk.Button(root)
        create_file_button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times', size=10)
        create_file_button["font"] = ft
        create_file_button["fg"] = "#000000"
        create_file_button["justify"] = "center"
        create_file_button["text"] = "Create Excel File"
        create_file_button.place(x=180, y=280, width=130, height=30)
        create_file_button["command"] = self.create_file_command



    def inputFileBrowseCommand(self):
        input_file_path = filedialog.askopenfilename(initialdir=r"C:\Users\ShaiPC\Desktop\Python - SyncTrayzor\excel_script_nexite",
                                              title="Select a File",
                                              filetypes=[("Excel file","*.xlsx")])
        self.input_file_message['text'] = f"{input_file_path}"
        self.input_file_path = input_file_path
        return input_file_path

    def browseSavePath(self):
        save_file_path = filedialog.askdirectory(initialdir=r"C:\Users\ShaiPC\Desktop\New folder",
                                              title='Select a directory for output file')
        self.save_path_message['text'] = save_file_path + "/output.xlsx"
        self.save_xlsx_path = save_file_path
        print(self.input_file_path)

        return save_file_path

    def create_file_command(self):
        listToMatrix(self.input_file_path,self.save_xlsx_path)

    def status_popup(self):
        top = Toplevel(root)
        top.geometry('150x150')
        top.title('Status')
        Label(top,text=f'{self.input_file}')





if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
