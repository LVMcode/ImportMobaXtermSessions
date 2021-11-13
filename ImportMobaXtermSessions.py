# $language = "Python"
# $interface = "1.0"

import re
import datetime

protocol_dict = {
    "22": "SSH2",
    "23": "Telnet"
}

def check_session_existence(session_path):
    try:
        objTest = crt.OpenSessionConfiguration(session_path)
        return True
    except Exception:
        return False

def import_mobaXterm_file():
    session_counter = 0
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
        lines = (line.split(" ") for line in filter_lines if line)
        for info in lines:
            if re.search('SubRep=\w+', info[0]):
                folder_name = info[-1].split("\\")[-1]
            else:
                first_element = info[0]
                session_name = first_element.split("=")[0]
                percentage_split = first_element.split("%")
                hostname = percentage_split[1]
                port = percentage_split[2]
                protocol = protocol_dict[port]

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
