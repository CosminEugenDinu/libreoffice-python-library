# testing
import sys
sys.path.append('.')
from libreoffice_py import document

calc = document.Calc()
sheet = calc.get_sheets()[0]
range = sheet.get_range(0, 0, 1, 1)
data = [
    ['Hello', 'World'],
    [0, 1]
]
range.set_data(data)
calc.save('./test/testfile.ods')

