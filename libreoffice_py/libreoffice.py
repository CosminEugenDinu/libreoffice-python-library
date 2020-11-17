from .libreoffice_process import LOprocess 

# from com.sun.star.beans import PropertyValue

class LibreOffice:
    def __init__(self):
        self._lo_proc = LOprocess()
        self._uno_ctx = self._lo_proc.connect()
        self._uno_smgr = self._uno_ctx.ServiceManager
        self._desktop = self._uno_smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self._uno_ctx )
        if not self._desktop:
            raise Exception("Failed to create LibreOffice desktop")

    def get_desktop(self):
        return self._desktop

    def terminate(self, desktop):
        self._lo_proc.terminate(desktop)


    