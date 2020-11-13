import sys 
from .config import config
from .libreoffice_core import uno
from .libreoffice_process import LOprocess 

# from com.sun.star.beans import PropertyValue

class Desktop:
    def __init__(self):
        lo_proc = LOprocess()
        uno_ctx = lo_proc.connect()
        uno_smgr = uno_ctx.ServiceManager
        desktop = uno_smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", uno_ctx )
        if not desktop:
            raise Exception("Failed to create OpenOffice desktop on port %s" % self.port)

    
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




