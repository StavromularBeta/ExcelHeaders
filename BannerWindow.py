import tkinter as Tk
from tkinter import font as tkFont


class BannerWindow(Tk.Frame):
    def __init__(self, parent, **kwargs):
        Tk.Frame.__init__(self, parent, **kwargs)
        self.banner_font = tkFont.Font(size=48, weight='bold')

    def make_banner(self):
        main_banner = Tk.Label(self, text="ExcelHeader", font=self.banner_font)
        main_banner.grid(sticky=Tk.W, padx=10)
