# $language = "Python"
# $interface = "1.0"

import re
import datetime

protocol_dict = {
    "109": "SSH2",
    "98": "Telnet",
    "91": "RDP",
    "130": "FTP",
    "140": "SFTP",
}

def check_session_existence(session_path):
    try:
        objTest = crt.OpenSessionConfiguration(session_path)
        return True
    except Exception:
        return False

def import_mobaXterm_file():
    session_counter = 0
    session_timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    mobaXterm_file = crt.Dialog.FileOpenDialog(
        "Please select MobaXterm File to be imported.",
        "Open",
        "",
        "MobaXterm Files (*.mxtsessions)|*.mxtsessions||")
    if mobaXterm_file == "":
        return
    with open(mobaXterm_file, 'r') as input_file:
        stripped = (line.strip() for line in input_file)
        filter_lines = filter(None, [re.sub(r"\[Bookmarks.*|ImgNum=.*", r"", i) for i in stripped])
        for info in filter_lines:
            if re.search('SubRep=.*', info):
                folder_name = info.split("=")[-1].replace("\\", "/")
            else:
                percentage_split = info.split("%")
                session_name = percentage_split[0].split("#")[0].rstrip("= ")
                protocol_code = percentage_split[0].split("#")[1]
                protocol = protocol_dict[protocol_code]
                hostname = percentage_split[1]
                port = int(percentage_split[2])

                session_path = "MobaXtermSessions/" + folder_name + "/" + session_name

                if check_session_existence(session_path):
                    session_path += "_imported_{0}".format(session_timestamp)

                objConfig = crt.OpenSessionConfiguration("Default")
                objConfig.SetOption("Protocol Name", protocol)
                objConfig.Save(session_path)

                objConfig = crt.OpenSessionConfiguration(session_path)
                objConfig.SetOption("Hostname", hostname)

                objConfig.Save()
                session_counter += 1
    crt.Dialog.MessageBox("Process completed. Imported {0} sessions".format(session_counter), "Import process info")


import_mobaXterm_file()
