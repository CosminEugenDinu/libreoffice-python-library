import os
from .libreoffice import LibreOffice
import uno



class Calc:
    def __init__(self):
        libreoffice = LibreOffice()
        self.desktop = libreoffice.get_desktop()
        PropertyValue = uno.getClass('com.sun.star.beans.PropertyValue')
        inProps = PropertyValue( "Hidden" , 0 , True, 0 ),
        # https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html
        self.document = self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, inProps )

    class Range:
        def __init__(self, _uno_range):
            self._uno_range = _uno_range
        def set_data(self, data):
            return self._uno_range.setData(data) 

    class Sheet:
        def __init__(self, _uno_sheet):
            self._uno_sheet = _uno_sheet
        
        def get_range(self, nLeft, nTop, nRight, nBottom):
            """
            Returns a sub-range of cells within the range.

            Parameters
                nLeft	is the column index of the first cell inside the range.
                nTop	is the row index of the first cell inside the range.
                nRight	is the column index of the last cell inside the range.
                nBottom	is the row index of the last cell inside the range.
                nSheet	is the sheet index of the sheet inside the document.
            Returns
                the specified cell range.
            Exceptions
                com::sun::star::lang::IndexOutOfBoundsException	if an index is outside the dimensions of this range.
            """

            _uno_range = self._uno_sheet.getCellRangeByPosition(nLeft, nTop, nRight, nBottom)
            return Calc.Range(_uno_range) 

    class Sheets:
        def __init__(self, document):
            self._uno_sheets = document.getSheets()

        def __getitem__(self, index):
            sheet = Calc.Sheet(self._uno_sheets[index])
            return sheet
    

            

    def get_sheets(self):
        """
        Returns the collection of sheets in the document.
        """
        sheets = Calc.Sheets(self.document)
        return sheets

    def save(self, path, filetype=None):

        IOException = uno.getClass('com.sun.star.io.IOException')

        # UNO requires absolute paths
        url = uno.systemPathToFileUrl(os.path.abspath(path))

        # Filters used when saving document.
        # https://github.com/LibreOffice/core/tree/330df37c7e2af0564bcd2de1f171bed4befcc074/filter/source/config/fragments/filters
        filetypes = dict(
            jpg='calc_jpg_Export',
            pdf='calc_pdf_Export',
            png='calc_png_Export',
            svg='calc_svg_Export',
            xls='Calc MS Excel 2007 XML',
            xlsx='Calc MS Excel 2007 XML Template',
        )

        if filetype:
            filter = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
            filter.Name = 'FilterName'
            filter.Value = filetypes[filetype] 
            filters = (filter,)
        else:
            filters = ()

        try:
            self.document.storeToURL(url, filters)
        except IOException as e:
            raise IOError(e.Message)

class Draw:
    def __init__(self):
        pass

class Writer:
    def __init__(self):
        pass
    def save(self):
        pass



