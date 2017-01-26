import shutil
from datetime import datetime

"""
This class handles file manipulation.
"""


class FileHelper():
    @staticmethod
    def archive(settings,excel):
        """
        Moves the active excel spreadsheet to the archive folder and renames it for that weeks range. Then copies the template to the active
        directory.
        :param settings: The currently used settings object.
        :param excel:
        :return:
        """
        end = datetime.strptime(excel.week_end(), "%m/%d/%Y")
        start = datetime.strptime(excel.week_start(), "%m/%d/%Y")
        _week_end = end.strftime("%m-%d-%Y")
        _week_start = start.strftime("%m-%d-%Y")
        new_name = _week_start + "_to_" + _week_end + ".xlsx"
        shutil.move(settings.my_path + "\\" + settings.active_xl, settings.archive_path +
                    "\\" + new_name)
        shutil.copy(settings.template_path + "\\" + settings.template_name, settings.my_path +
                    "\\" + settings.active_xl)
