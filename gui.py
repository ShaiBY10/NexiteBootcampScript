import tkinter.font as tkFont
import tkinter.messagebox
from tkinter import ttk, messagebox, END
from ttkthemes import ThemedTk
from algo import *


class App():
    def __init__(self, tk):
        self.tk = tk
        self.input_file_path = None
        self.save_xlsx_path = None
        self.input_file = None
        self.bootcamp_dir = r"G:\.shortcut-targets-by-id\1-9lckUfMsCWBy2igthBlY7BcDCwNSAii\Nexite Workspace\System\Trials and night tests\Bootcamps\2022"

        root.title("Bootcamp Script")
        root.tk.call('tk','scaling',2.0)
        # setting window size
        width = 500
        height = 250
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
        root.columnconfigure(0, weight=1, pad=10)
        root.columnconfigure(1, weight=1)
        root.columnconfigure(2, weight=1)

        # Bootcamp Matrix Title
        title = ttk.Label(root)
        ft = tkFont.Font(family='Times', size=18)
        title["font"] = ft
        title["text"] = "Bootcamp Matrix Creator"
        title.grid(column=0, row=0, columnspan=2)

        # Input Browse Button
        input_browse_button = ttk.Button(root)
        input_browse_button["text"] = "Select Input File..."
        input_browse_button["command"] = self.inputFileBrowseCommand
        input_browse_button.grid(column=0, row=1, pady=30)

        # Input Text Bar (where the selected path is shown)
        self.input_textbox = ttk.Entry(root)
        self.input_textbox.grid(column=1, row=1, sticky='ew')

        # Input File label
        input_file_label = ttk.Label(root)
        ft = tkFont.Font(family='Operator Mono', size=13)
        input_file_label["font"] = ft
        input_file_label["justify"] = "center"
        input_file_label["text"] = "Input File:"

        # Output File name label
        output_name_label = ttk.Label(root)
        ft = tkFont.Font(family='Operator Mono', size=12)
        output_name_label["font"] = ft
        output_name_label["justify"] = "center"
        output_name_label["text"] = "Output Name:"
        output_name_label.grid(column=0, row=2)

        # Create file button
        create_file_button = ttk.Button(root)
        create_file_button["text"] = "Create Excel File"
        create_file_button.grid(column=1, row=3, sticky="w", pady=5)
        create_file_button["command"] = self.create_file_command

        # Output file textbox
        self.output_file_textbox = ttk.Entry(root)
        self.output_file_textbox.grid(column=1, row=2, sticky='we')

    def inputFileBrowseCommand(self):
        input_file_path = filedialog.askopenfilename(
            initialdir=self.bootcamp_dir,
            title="Select a File",
            filetypes=[("Excel file", "*.xlsx")])
        self.input_textbox.insert(END, str(input_file_path))
        self.input_file_path = input_file_path
        print(input_file_path)
        return input_file_path

    def create_file_command(self):
        global raw_file_len
        output_textbox = self.output_file_textbox.get()
        print(output_textbox)
        sliced_path = self.input_file_path[:-5]
        raw_file_len = len(read_excel(self.input_file_path, header=None))

        try:

            if output_textbox == "":
                now = getTime()
                print(now)

                listToMatrix(self.input_file_path, f'{sliced_path}_{now}.xlsx')
            else:
                listToMatrix(self.input_file_path, f'{sliced_path}{output_textbox}.xlsx')

            df = read_excel(io=self.input_file_path, sheet_name=0, header=None)
            raw_list = list([value for cell, value in df[0].iteritems()])
            duplicates = duplicates_check(raw_list)
            if duplicates:
                tkinter.messagebox.showerror(
                    message='Found duplicates in the excel file! \n Check the file:' + self.input_file_path,
                    )
            else:
                tkinter.messagebox.showinfo(
                    message=f'No Duplicates Found! \n The length of the input file is {raw_file_len}\n Created {(raw_file_len // 104)} Boards\nPath: {sliced_path}{output_textbox}')
        except Exception as VE:
            messagebox.showerror(title="Raised ERROR MESSAGE", message=str(VE))


if __name__ == "__main__":
    root = ThemedTk(theme='breeze')
    app = App(root)
    root.mainloop()
