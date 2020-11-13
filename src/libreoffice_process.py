import time
import subprocess
from .config import config
from .libreoffice_core import uno
from com.sun.star.connection import NoConnectException

now = time.time

from tools.messages import Messages

# from .libreoffice_core import ProtertyValue
from com.sun.star.beans import PropertyValue

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
        self.messages = []

    def startup(self, port=None):
        """ Starts libreoffice process.
        """
        add_msg, get_msgs = Messages('info')

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
        add_msg(f'{now()} Starting libreoffice process {start_command}')
        lo_proc = subprocess.Popen(start_command, shell=True)
        add_msg(f'{now()} Process shell_pid={lo_proc.pid} started.')

        # I could'n connect to process if opended it like this:
        # lo_proc = subprocess.Popen(start_command.split())
        # add_msg(f'{now()} Process pid={lo_proc.pid} started.')

        self.messages.append(get_msgs())
        return lo_proc

    def connect(self, host=None, port=None):
        add_msg, get_msgs = Messages('info')

        local_ctx = uno.getComponentContext()
        smgr_local = local_ctx.ServiceManager
        resolver = smgr_local.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        url = self.connection_url % (self.host, (port or self.port))

        uno_ctx = None
        # try to resolve connection
        try:
            # url = "uno:socket,host=localhost,port=2002,tcpNoDalay=1;urp;StarOffice.ComponentContext"
            # url = "uno:pipe,name=somepipename;urp;StarOffice.ComponentContext"
            uno_ctx = resolver.resolve(url)
            add_msg(f'{now()} Connected to libreoffice process via {url}')
        except NoConnectException as nce:
            add_msg(f'{now()} Got: {nce}')
            # launch libreoffice process
            lo_proc = self.startup()

            timeout = 0
            while timeout < self.timeout:
                # Is it already/still running?
                retcode = lo_proc.poll()
                if retcode == 81:
                    add_msg(f'{now()} Caught exit code 81 (new installation of libreoffice ?).')
                    # self.connect()
                    break
                elif retcode is not None:
                    add_msg(f'{now()} Process pid={lo_proc.pid} exited with {retcode}.')
                    raise
                try:
                    uno_ctx = resolver.resolve(url)
                    add_msg(f'{now()} Connected to libreoffice process via {url}')
                    break
                except NoConnectException:
                    time.sleep(0.5)
                    timeout += 0.5
                except:
                    raise
            else:
                raise
        except Exception:
            raise
        self.messages.append(get_msgs())
        print(self.messages)
        return uno_ctx

    def shutdown(self, desktop):
        add_msg, get_msgs = Messages('info')
        desktop.terminate()
        add_msg(f'{now()} Desktop {1} + lo_process {2} terminated!')
        print('I\'m  OFF now!')
