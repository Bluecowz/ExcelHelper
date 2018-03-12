from tkinter import *

"""
Static class to handle feedback pop ups.
"""


class Popups:
    @staticmethod
    def email_sent():
        """
        Confirmation pop that email sent successfully.
        :return:
        """
        tkMessageBox.showinfo("Email Sent", "Email was sent successfully.",icon='info')

    @staticmethod
    def email_failed():
        """
        Displays an error message if email fails to send.
        :return:
        """
        tkMessageBox.showinfo("Email Failed", "Email failed to send.", icon='warning')

    @staticmethod
    def failed_edit():
        """
        If time manager failed to edit the spread sheet.
        :return:
        """
        tkMessageBox.showerror("Failure", "Failed to add to excel.", icon='warning')

    @staticmethod
    def full_warning():
        """
        Gives the user warning if the spread sheet is full.
        :return:
        """
        root = Tk()
        root.withdraw()
        tkMessageBox.showwarning("Time sheet full", "Time sheet is full. Recommend sending it in to create new sheet.", icon='warning')

    @staticmethod
    def send_confirmation():
        """
        Displays confirmation message to send spread sheet in email.
        :return: If sending is confirmed.
        """
        confirm = tkMessageBox.askquestion("Send", "Confirm sending time sheet?", icon='warning')
        if confirm == "yes":
            return True
        else:
            return False