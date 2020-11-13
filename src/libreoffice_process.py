import time
import subprocess
from .config import config
from .libreoffice_core import uno
from com.sun.star.connection import NoConnectException

from tools.messages import Messages

# from .libreoffice_core import ProtertyValue
# from com.sun.star.beans import PropertyValue

class LOprocess:
    def __init__(self,
                 connection=config['libreoffice']['connection']):

        # command flags array
        self.flags = config['libreoffice']['flags']

        # libreoffice binary location
        self.libreoffice_bin = connection['binary_location']
        # # UNO socket connection settings
        self.host = connection['host']
        self.port = connection['port']
        # will become content of '--accept=' flag, like f"--accept='{accept_open}'"
        # something like "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"
        self.accept_open = connection['accept_open']
        # this is used when connetion to running libreoffice process
        # something like "uno:socket,host=%s,port=%s,tcpNoDalay=1;urp;StarOffice.ComponentContext"
        self.connection_url = connection['connection_url']
        self.timeout = connection['timeout']

        # process ref of running libreoffice set from startup method
        self.loproc = None
        self.desktop = None

    def startup(self, port=None):
        """ Starts libreoffice process.
        """
        accept_open = self.accept_open % (self.host, (port or self.port))
        # "--accept=%s" => %s is replaced with accept_open
        # stringify command flags
        start_command_flags = ' '.join(self.flags) % accept_open

        # command used to start libreoffice with open socket
        start_command = f"{self.libreoffice_bin} {start_command_flags}"

        # loproc = Popen(start_command.split())
        # self.loproc = subprocess.Popen(start_command.split(), close_fds=True)
        # lo_proc = subprocess.Popen(start_command.split())
        # "--accept='pipe,name=somepipename;urp;StarOffice.ComponentContext'"
        lo_proc = subprocess.Popen(start_command, shell=True)
        return lo_proc

    def connect(self, host=None, port=None):
        
        local_ctx = uno.getComponentContext()
        smgr_local = local_ctx.ServiceManager
        resolver = smgr_local.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        url = self.connection_url % (self.host, (port or self.port))
        add_msg, get_msgs = Messages('warning')
        add_msg('first mess')
        add_msg('sec meg')
        add_msg('first_mess')
        messages = get_msgs()
        print(messages)
        return

        uno_ctx = None
        # try to resolve connection
        try:
            # url = "uno:socket,host=localhost,port=2002,tcpNoDalay=1;urp;StarOffice.ComponentContext"
            # url = "uno:pipe,name=somepipename;urp;StarOffice.ComponentContext"
            uno_ctx = resolver.resolve(url)
        except NoConnectException as nce:
            # launch libreoffice process
            lo_proc = self.startup()

            timeout = 0
            while timeout < self.timeout:
                # Is it already/still running?
                retcode = lo_proc.poll()
                raise Exception(lo_proc.pid)
                if retcode == 81:
                    # info(3, "Caught exit code 81 (new installation). Restarting listener.")
                    return self.connect()
                    break
                elif retcode is not None:
                    # info(3, "Process %s (pid=%s) exited with %s." % (office.binary, ooproc.pid, retcode))
                    raise
                try:
                    uno_ctx = resolver.resolve(url)
                    break
                except NoConnectException:
                    time.sleep(0.5)
                    timeout += 0.5
                except:
                    raise
            else:
                print('hot here......................................... ')
                raise
        except Exception as e:
            raise
            # error("Launch of %s failed.\n%s" % (office.binary, e))
        return uno_ctx

    def shutdown(self):
        print('I\'m  OFF now!')
