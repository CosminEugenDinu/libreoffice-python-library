# testing
import sys
sys.path.append('.')
from libreoffice_py import document
from calendar import Calendar
from datetime import date
  
today = date.today()

calc = document.Calc()
sheet = calc.get_sheets()[0]

month = 4
cal = Calendar(firstweekday=0)
monthdates = [
    f'{d[0]}.{d[1]}.{d[2]}' for d in cal.itermonthdays4(today.year, month)
    if d[3] not in (5,6) and d[1] == month
]

data = [
    ['Firm', 'BRIGHT', '', '', ''],
    ['Address', 'Street no.', '', '', ''],
    ['Employee', 'Jack Sprot', '', '', ''],
    ['Date', 'Arrive at', 'Sign at arrival', 'Leave at', 'Sign at arrival']
]

for d in monthdates:
    data.append([d, '', '08:00', '', '18:00'])

for line in data:
    print(line)
range = sheet.get_range(0, 0, len(data[0])-1, len(data)-1)

range.set_data(data)
calc.save('./test/testfile.pdf', 'pdf')
calc.desktop.terminate()