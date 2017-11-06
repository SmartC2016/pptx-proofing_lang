"""
This is a 'test' for the window for pptx-proofing language
"""

import tkinter as tk
from tkinter import filedialog
import os


class App(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.pack()
        self.master.title('pptx - proofing language')
        self.master.resizable(False, False)  # Now the window cannot be resized
        self.master.tk_setPalette(background='#ececec')  # Give it the same backgound color as regular Mac OS X

        # position the window in the middle of the screen and in the upper 3rd
        x = int((self.master.winfo_screenwidth() - self.master.winfo_reqwidth()) / 2)
        y = int((self.master.winfo_screenheight() - self.master.winfo_reqheight()) / 3)
        self.master.geometry(f'400x400+{x}+{y}')

        self.master.config(menu=tk.Menu(self.master))  # Disable the regular Python Menu ...

        # Bind keyboard keys to functions ...
        self.master.bind('<Escape>', self.click_cancel)

        self.myfilename = tk.StringVar(self.master)
        self.myfilename.set('')

        # Frame in the frame
        self.dialog_frame = tk.Frame(self)  # this frame is in the master frame
        self.dialog_frame.pack(padx=20, pady=15)

        # tk.Label(dialog_frame, text='This is your first GUI. (highfive)').pack()
        the_label_1 = tk.Label(self.dialog_frame, text='This little "helper" will change the proofing language\non all '
                                                'slides of your PowerPoint presentation.')
        the_label_1.grid(row=0, column=0)

        open_pptx_btn = tk.Button(self.dialog_frame, text='Open a Powerpoint Presentation',
                                  command=self.selectFile)
        open_pptx_btn.grid(row=1, column=0, padx=5, pady=5)

        the_label_2 = tk.Label(self.dialog_frame, text='Your selected file:')
        the_label_2.grid(row=3, column=0)
        the_label_3 = tk.Label(self.dialog_frame, textvariable=self.myfilename)
        the_label_3.grid(row=4, column=0)


        # The dropdown menu
        my_choices = {'Yellow', 'Blue', 'Green', 'Black', 'Red'}
        self.tkVar = tk.StringVar(self.master)
        self.tkVar.set('Green')

        dropdownMenu = tk.OptionMenu(self.dialog_frame, self.tkVar, *my_choices)
        dropdownMenu.grid(row=5, column=0)

        self.tkVar.trace('w', self.change_dropdown)




        # Seperate Frame for the buttons
        button_frame = tk.Frame(self)  # this frame also is in the master frame
        button_frame.pack(padx=15, pady=(0, 15), anchor='e')

        tk.Button(button_frame, text='OK', default='active', command=self.click_ok).pack(padx=5, pady=5, side='right')
        tk.Button(button_frame, text='Cancel', command=self.click_cancel).pack(padx=5, pady=5, side='right')

    def change_dropdown(self, *args):
        print(self.tkVar.get())
        return

    def selectFile(self):
        my_file_types = [('Powerpoint presentation', '.pptx'), ]
        answer = filedialog.askopenfilename(initialdir=os.getcwd(),
                                            title='Please select a Powerpoint presentation',
                                            filetypes=my_file_types)
        print(answer)
        self.myfilename.set(answer)
        return


    def click_ok(self, event=None):
        print('The user clicked "OK".')
        return

    def click_cancel(self, event=None):
        print('The user clicked "Cancel".')
        self.master.destroy()
        return


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    app.mainloop()
