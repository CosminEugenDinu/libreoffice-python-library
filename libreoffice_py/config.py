import toml


def _get_lo_config():
    lo_config_default = """
    [libreoffice]
    # To find this location, open python3 interactive, >>> import uno; print(uno.__file__)
    python_uno_location = "/usr/lib/python3/dist-packages/"

    # soffice.bin should be run with these flags
    flags = ["--headless", "--invisible", "--nocrashreport", "--nodefault", "--nofirststartwizard", "--nologo", "--norestore", "--accept='%s'"]

    [libreoffice.connection]
    binary_location = "/usr/lib/libreoffice/program/soffice.bin"
    # UNO socket connection settings
    host = "localhost"
    port = 2002
    # content of '--accept=' flag for starting soffice.bin with open connection
    accept_open = "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"
    # this is used to connect running libreoffice process with opened socket
    connection_url = "uno:socket,host=%s,port=%s,tcpNoDalay=1;urp;StarOffice.ComponentContext"
    # how much time to wait for connecting
    timeout = 3
    """
    # load connection settings from Config toml file
    try:
        config = toml.load('./lo_config.toml')
    except FileNotFoundError as e:
        # creating default libreoffice config file
        print("Creating file ./lo_config.toml")
        with open('./lo_config.toml', 'w') as lo_config:
            lo_config.write(lo_config_default)
    
    return toml.loads(lo_config_default)



config = _get_lo_config()