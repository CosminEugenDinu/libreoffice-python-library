#!/usr/bin/env bash

# chmod +x open_soffice_socket.sh
# ./open_soffice_socket.sh

# soffice --headless --nologo --nofirststartwizard --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"

# soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless
# soffice --accept='socket,host=localhost,port=8100;urp;StarOffice.Service'


soffice_bin=/usr/lib/libreoffice/program/soffice.bin

$soffice_bin --headless --nologo --nofirststartwizard --norestore --nodefault --accept='socket,host=localhost,port=2002;urp;StarOffice.Service'


