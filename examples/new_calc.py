import sys
sys.path.append('.')

from src import document

calc = document.Calc()

sheets = calc.get_sheets()
sheet_1 = sheets[0]
cell_1 = sheet_1.getCellByPosition(1,1)
value = cell_1.setString('Hello World!')

calc.save('./examples/saved/somefile.ods')
calc.save('./examples/saved/somefile.jpg', 'jpg')
# calc.save('./examples/somefile.svg', 'svg')
calc.save('./examples/saved/somefile.pdf', 'pdf')
calc.save('./examples/saved/somefile.png', 'png')
calc.save('./examples/saved/somefile.xls', 'xls')
calc.save('./examples/saved/somefile.xlsx', 'xlsx')

calc.desktop.terminate()


