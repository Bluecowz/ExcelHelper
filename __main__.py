from Tkinter import Tk
import TimeGUI
from Settings import Settings

"""
* This program handles spreadsheet manipulation for work. Because it is a pain in the ass.
* Author: Mike Gallant
* Version: 1.2
"""

"""
* Bugs/Improvements

* Improvements
    * Pass sending email/editing to new thread so UI does not lag.

"""


root = Tk()
settings = Settings()
gui = TimeGUI.TimeGUI(root, settings)
root.title("Time Manager")
root.resizable(height=False,width=False)
root.geometry('{}x{}'.format(400, 300)) #width x height
root.mainloop()
