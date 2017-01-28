from xml.dom.minidom import parse

"""
This class stores the variables used throughout the program.
"""


class Settings:

    # Paths
    my_path = ""
    archive_path = ""
    template_path = ""
    template_name = ""
    active_xl = ""

    # Important Excel Coordinates
    new_sheet_check_X = 5
    new_sheet_check_Y = 7
    data_colm = 1  # A column
    desc_colm = 6  # F column

    # Email Information
    dest_addr = ['', '']
    smtp_server = ''
    smtp_port = 587
    name = ''

    # Log in credentials
    user = ""
    password = ""

    def __init__(self):
        self._get_settings()
    
    @staticmethod
    def _get_xml():
        tree = parse('config.xml')
        settings = tree.getElementsByTagName('Settings')
        return settings

    def _get_settings(self):
        settings = Settings._get_xml()
        self.my_path = settings[0].getElementsByTagName('MyPath')[0].childNodes[0].nodeValue
        self.archive_path = settings[0].getElementsByTagName('ArchivePath')[0].childNodes[0].nodeValue
        self.template_path = settings[0].getElementsByTagName('TemplatePath')[0].childNodes[0].nodeValue
        self.template_name = settings[0].getElementsByTagName('TemplateName')[0].childNodes[0].nodeValue
        self.active_xl = settings[0].getElementsByTagName('ActiveXL')[0].childNodes[0].nodeValue
        self.user = settings[0].getElementsByTagName('YourEmail')[0].childNodes[0].nodeValue
        self.name = settings[0].getElementsByTagName('Name')[0].childNodes[0].nodeValue
        self.dest_addr = str(settings[0].getElementsByTagName('Recipient')[0].childNodes[0].nodeValue)
        self.smtp_server = settings[0].getElementsByTagName('STMPServer')[0].childNodes[0].nodeValue
        self.smtp_port = settings[0].getElementsByTagName('SMTPPort')[0].childNodes[0].nodeValue
        self.password = settings[0].getElementsByTagName('Password')[0].childNodes[0].nodeValue

