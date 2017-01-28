from Tkinter import Tk
import TimeGUI
from Settings import Settings

"""

"""


def main():
    root = Tk()
    settings = Settings()
    TimeGUI.TimeGUI(root, settings)
    root.title("Time Manager")
    root.resizable(height=False, width=False)
    root.geometry('{}x{}'.format(400, 300)) #width x height
    root.mainloop()

main()

