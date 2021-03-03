# libreoffice-python-library
Use python to manipulate the LibreOffice linux process. This framework can be used to create formatted text documents, pdfs,  spreadsheets or drawings.

## Prerequisites
- [python3.x](https://docs.python.org/3/using/unix.html)
- [pipenv](https://github.com/pypa/pipenv)
- [LibreOffice](https://gist.github.com/CosminEugenDinu/d584dddfce534f8272ab9f661eb480a5#file-install_libreoffice-sh)

## Install
```bash
pipenv install -e "git+https://github.com/CosminEugenDinu/libreoffice-python-library.git#egg=libreoffice-py"
```

## Setup
- edit [***lo_config.toml***](https://raw.githubusercontent.com/CosminEugenDinu/libreoffice-python-library/main/lo_config.toml) file to match your system configuration (if does not exists, it is created at first run)

## Usage
```py
from libreoffice_py import document

# Create a new spreadsheet (calc) document
calc = document.Calc()
sheets = calc.get_sheets()
sheet_0 = sheets[0]
cell_0 = sheet_0._uno_sheet.getCellByPosition(0,0)
cell_0.setString('Hello World!')

range = sheet_0.get_range(1, 1, 2, 2)
data = [
    ['Hello', 'World'],
    [0, 1]
]
range.set_data(data)

calc.save('./examples/saved/somefile.ods')
calc.save('./examples/saved/somefile.jpg', 'jpg')
calc.save('./examples/saved/somefile.pdf', 'pdf')
calc.save('./examples/saved/somefile.png', 'png')
calc.save('./examples/saved/somefile.xls', 'xls')
calc.save('./examples/saved/somefile.xlsx', 'xlsx')

# don't forget to kill the libreoffice process
calc.desktop.terminate()
```

## License
***libreoffice-python-library*** is distributed under GNU GPL V3. Because you will implicitely also use ***LibreOffice*** source code and binaries, its license Mozilla Public License v.2.0 applies as well. Copies of both are included in this repository.
