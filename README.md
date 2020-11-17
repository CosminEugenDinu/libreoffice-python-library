# libreoffice-python-library
LibreOffice scripting using python

## Prerequisites
- [python3.x](https://docs.python.org/3/using/unix.html)
- [pipenv](https://github.com/pypa/pipenv)
- [LibreOffice](https://gist.github.com/CosminEugenDinu/d584dddfce534f8272ab9f661eb480a5#file-install_libreoffice-sh)

## Install
```bash
pipenv install -e "git+https://github.com/CosminEugenDinu/libreoffice-python-library.git@pypi#egg=libreoffice-py"
```

## Setup
- edit [***lo_config.toml***](https://raw.githubusercontent.com/CosminEugenDinu/libreoffice-python-library/lo_config.toml) file to match your system configuration (if does not exists, it is created at first run)

## Usage
```py
from libreoffice_py import document

# Create a new spreadsheet (calc) document
calc = document.Calc()
sheets = calc.get_sheets()
sheet_1 = sheets[0]
cell_1 = sheet_1.getCellByPosition(1,1)
cell_1.setString('Hello World!')

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
