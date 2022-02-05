# python3
# coding: utf-8

import sys
import os
import logging
import logging.handlers

try:
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtGui import QIcon
    from PyQt5.QtCore import QSize
    used_Qt_Version = 5
    print("PyQt Version:", used_Qt_Version)
except Exception as e:
    print(e)
    exit()
    pass
import MainProg
from MainProg import __version__ as progversion

# example for documentation style:
# http://www.stack.nl/~dimitri/doxygen/manual/docblocks.html#pythonblocks


def compile_GUI():
    if used_Qt_Version == 4:
        print("Compile QUI for Qt Version: " + str(used_Qt_Version))
        os.system("pyuic4 -o Converter_ui.py Converter.ui")
    elif used_Qt_Version == 5:
        print("Compile QUI for Qt Version: " + str(used_Qt_Version))
        os.system("pyuic5 -o Converter_ui.py Converter.ui")


cwd = os.getcwd()
loggfilePath = os.path.sep.join([cwd,'log_file.log'])

# # \brief Short description.
# Longer description.
# \param self
# \param name
if __name__ == "__main__":
    # logger = init_log()
    app = QApplication(sys.argv)

    logger = logging.getLogger('SR_E2W2P')
    logger.setLevel(logging.INFO)

    if os.path.exists(loggfilePath):
        try:
            os.remove(loggfilePath)
        except:
            print("Error while deleting file ", loggfilePath)
    else:
        print("Can not delete the file as it doesn't exists")

    log_format = logging.Formatter(
        '%(asctime)s [%(name)s] %(levelname)s - %(message)s')
    log_handler = logging.handlers.RotatingFileHandler(
        loggfilePath, maxBytes=1048576, backupCount=10)
    log_handler.setFormatter(log_format)
    logger.addHandler(log_handler)

    logger.info('Start Logging')
    logger.info("cwd: " + str(loggfilePath))

    gui = MainProg.MainProg()
    gui.setWindowTitle("SchoolReport_Excel2Word2PDF (" + str(progversion) + ")")
    app_icon = QIcon()
    app_icon.addFile('GUI/icons/16x16.png', QSize(16, 16))
    app_icon.addFile('GUI/icons/24x24.png', QSize(24, 24))
    app_icon.addFile('GUI/icons/32x32.png', QSize(32, 32))
    app_icon.addFile('GUI/icons/48x48.png', QSize(48, 48))
    app_icon.addFile('GUI/icons/64x64.png', QSize(64, 64))
    app_icon.addFile('GUI/icons/128x128.png', QSize(128, 128))
    app_icon.addFile('GUI/icons/256x256.png', QSize(256, 256))
    app.setWindowIcon(app_icon)
    # gui.setWindowIcon(QIcon("img" + os.path.sep + "Logo.png"))

    gui.show()
    sys.exit(app.exec_())
