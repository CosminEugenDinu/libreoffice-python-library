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
        self._document = self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, inProps )

    def get_sheets(self):
        sheets = self._document.getSheets()
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
            self._document.storeToURL(url, filters)
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



