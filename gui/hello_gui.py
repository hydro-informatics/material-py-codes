import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import askokcancel, askyesno, showinfo
from tkinter.filedialog import *


class MyApp(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.master.title("Graphical User Interface")
        self.master.iconbitmap("gui_icon.ico")
        self.pack()
        
        # Set geometry: upper-left corner of the window
        ww = 628  # width
        wh = 382  # height
        wx = (self.master.winfo_screenwidth() - ww) / 2
        wy = (self.master.winfo_screenheight() - wh) / 2
        # assign geometry
        self.master.geometry("%dx%d+%d+%d" % (ww, wh, wx, wy))

        # Menu Bar
        self.mbar = tk.Menu(self)  # create standard menu bar
        self.master.config(menu=self.mbar)  # make self.mbar standard menu bar
        # add menu entry
        self.ddmenu = tk.Menu(self, tearoff=0)
        self.mbar.add_cascade(label="Add A Menu", menu=self.ddmenu)  # attach entry it to standard menu bar
        self.ddmenu.add_command(label="Credits", command=lambda: showinfo("Hi!", "Info message"))

        # Image placement
        logo = tk.PhotoImage(file=os.path.dirname(os.path.abspath(__file__)) + "\\hello_gui.gif")
        logo = logo.subsample(1, 1)
        self.l_img = tk.Label(self, image=logo)
        self.l_img.image = logo
        self.l_img.grid(row=0, column=0)

        self.a_button = tk.Button(master, text="Click the Button")
        self.a_button.pack()


if __name__ == '__main__':
    MyApp().mainloop()
