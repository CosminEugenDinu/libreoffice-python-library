from .libreoffice import LibreOffice


class Calc:
    def __init__(self):
        libreoffice = LibreOffice()
        desktop = libreoffice.get_desktop()
        print(dir(desktop))
        desktop.terminate()

class Draw:
    def __init__(self):
        pass

class Writer:
    def __init__(self):
        pass


    # def start_lo_service(self):
    # def newCalc(self):
        # inProps = PropertyValue( "Hidden" , 0 , True, 0 ),
        # self.newdoc = desktop.loadComponentFromURL(
        #     "private:factory/scalc", "_blank", 0, inProps )
        # print(dir(self.newdoc))
        # pass

    def save(self):
        pass
        # print(self.name, 'successfully saved') 
        # cwd = systemPathToFileUrl( getcwd() )
        # destFile = absolutize( cwd, systemPathToFileUrl(outputfile) )



