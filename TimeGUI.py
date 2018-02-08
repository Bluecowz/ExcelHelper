from Tkinter import Frame, Button, Text,\
    TOP, RIGHT, END, LEFT, BOTTOM, Label
from ExcelHelper import ExcelHelper
from Email import Email
from PopupHelper import Popups
from threading import Thread
import tkMessageBox

"""
This class generates the GUI and events needed.
"""


class TimeGUI:
    _excel = ExcelHelper()

    def __init__(self, master, settings):
        self._settings = settings
        self._frame = Frame(master)
        self._frame.pack()
        self._check_in_b = Button(self._frame, text="  TIME  ", command=self._add_time)
        self._check_in_b.pack(side=TOP)

        self._desc = Text(self._frame, height=12)
        self._desc.bind('<Control-t>', self._add_time)
        self._desc.bind('<Control-d>', self._desc_and_distract)
        self._desc.pack(side=TOP)

        self._add_desc_b = Button(self._frame, text="  ADD  ", command=self._desc_and_distract)
        self._add_desc_b.pack(side=RIGHT)

        self._send_email_b = Button(self._frame, text="  SEND  ", command=self._send_email)
        self._send_email_b.pack(side=LEFT)

        self._frame.bind('<Control-t>', self._add_time)
        self._frame.bind('<Control-d>', self._add_desc)
        self._frame.bind('<Control-s>', self._send_email)

        self._instructions = Label(self._frame,
                                   text="Ctrl + t -> Add Time\nCtrl + d -> Add Desc\nCtrl + s -> Send Sheet")
        self._instructions.pack(side=BOTTOM)

        self._frame.focus_set()

    def _get_text(self):
        """
        Gets the text from the text box object
        :return: Current text in the text box
        """
        desc = self._desc.get("1.0", 'end-1c')
        return desc

    def _add_time(self, event=None):
        """
        Event used in hot keys and button presses. Adds time to spreadsheet.
        :param event: Used when event is called on hot key press.
        :return:
        """
        try:
            self._excel.add_time()
        except:
            Popups.failed_edit()

    def _add_desc(self, event=None):
        """
        Event used in hot keys and button presses. Adds description to spreadsheet.
        :param event: Used when event is called on hot key press.
        :return:
        """
        try:
            self._excel.add_desc(self._get_text())
            self._desc.delete("1.0", END)
        except:
            Popups.failed_edit()

    def _desc_and_distract(self, event=None):
        """
        This event is used to submit the description and then cear the focus.
        :param event: Used when event is called on hot key press.
        :return:
        """
        self._add_desc()
        self._lose_focus()

    def _lose_focus(self):
        """
        Sets the focus on the frame. So that hot keys can be called.
        :return:
        """
        self._frame.focus_set()

    def _send_email(self, event=None):
        """
        Event to send email on hot key call or button press
        :param event: Used when event is called on hot key press.
        :return:
        """
        if Popups.send_confirmation():
            thread = Thread(target=Email.send_email, args=(self._settings, self._excel,))
            thread.start()
            thread.join()




