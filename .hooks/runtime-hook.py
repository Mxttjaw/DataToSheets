
import os
import sys

if sys.platform == 'linux':
    os.environ['TK_LIBRARY'] = os.path.join(sys._MEIPASS, 'tk')
    os.environ['TCL_LIBRARY'] = os.path.join(sys._MEIPASS, 'tcl')