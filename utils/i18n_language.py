import sys
import importlib
import gettext
import os
import tray
from utils.sqlite_tools import g_SqlLiteTool

importlib.reload(sys)


def i18n_main():
    try:
        sqlite_tool = g_SqlLiteTool
        result = sqlite_tool.start(table="client_settings", operation_type=2, settings=None, language=None)
        if result.get("result") != "Fail":
            data = result.get("data")
            tray.setLanguage(data['system_language'])
        else:
            tray.setLanguage('en_US')
    except Exception as e:
        tray.setLanguage('en_US')

    system_language = tray.getLanguage()
    work_path = os.getcwd() + '\locale'
    gettext.install('lang', work_path)
    gettext.translation('lang', work_path, languages=[system_language]).install()
