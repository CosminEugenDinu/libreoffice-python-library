# testing
import sys
sys.path.append('.')
from libreoffice_py import document

calc = document.Calc()
sheet = calc.get_sheets()[0]
range = sheet.get_range(0, 0, 0, 0)
data = range.set_data([[1]])
calc.save('./test/testfile.ods')

