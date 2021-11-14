# $language = "Python"
# $interface = "1.0"

import re
import datetime
import string

#######################################################################################
# GLOBAL OPTIONS
# Set overwrite_existing_sessions to True if you wish to overwrite existing sessions
# in case of full path name coincidences
overwrite_existing_sessions = False

# default_protocol and default_port define protocol and port values in case an imported
# session is not included in the supported protocols
# Currently supported protocols are: SSH2, Telnet and RDF
default_protocol = "SSH2"
default_port = 22

# Set the root import folder for your imported sessions
root_import_folder = "MobaXtermSessions"

# Set has_trailing to False if you don't wish to have the _imported_timestamp trailing in
# your root import folder name
has_trailing = True

# Set has_session_log to True if you wish to create a log file for your sessions
# Use default_session_log_format to set the log format of your imported sessions. You
# will be able to change it later in the Select Log File prompt window. This option has
# no effect if has_session_log is set to False
has_session_log = False
default_session_log_format = "session_%S_%Y%M%D.log"
#######################################################################################

protocol_dict = {
    "109": "SSH2",
    "98": "Telnet",
    "91": "RDP"
}


def check_session_existence(session_path):
    try:
        objTest = crt.OpenSessionConfiguration(session_path)
        return True
    except Exception:
        return False


def format_filename(s):
    valid_chars = "-_.%() {0}{1}".format(string.ascii_letters, string.digits)
    filename = ''.join(c for c in s if c in valid_chars).strip()
    return re.sub(r"(/|-|_|\.|%|\(|\))+$", r".log", filename)


def import_mobaXterm_file():
    session_counter = 0
    session_timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    timestamp_trailing = "_imported_{0}".format(session_timestamp) if has_trailing else ""
    mobaXterm_file = crt.Dialog.FileOpenDialog(
        "Please select MobaXterm File to be imported.",
        "Open",
        "",
        "MobaXterm Files (*.mxtsessions)|*.mxtsessions||")
    if mobaXterm_file == "":
        return
    if has_session_log:
        formatted_default_session_log_format = format_filename(default_session_log_format)
        logs_path = crt.Dialog.FileSaveDialog(
            "Select Log File",
            "Save",
            "{0}".format(formatted_default_session_log_format),
            "Log Files (*.log)|*.log||")
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
                hostname = percentage_split[1]
                port = int(percentage_split[2])

                try:
                    protocol = protocol_dict[protocol_code]
                except KeyError:
                    protocol = default_protocol
                    port = default_port

                raw_session_path = "/".join([root_import_folder + timestamp_trailing, folder_name, session_name])
                session_path = re.sub(r"//+", r"/", raw_session_path).strip("/")

                if check_session_existence(session_path) and not overwrite_existing_sessions:
                    session_path += "_imported_{0}".format(session_timestamp)

                objConfig = crt.OpenSessionConfiguration("Default")
                objConfig.SetOption("Protocol Name", protocol)
                objConfig.Save(session_path)

                objConfig = crt.OpenSessionConfiguration(session_path)

                objConfig.SetOption("Hostname", hostname)

                if protocol == "SSH2":
                    objConfig.SetOption("[SSH2] Port", port)
                elif protocol == "Telnet":
                    objConfig.SetOption("Port", port)
                elif protocol == "RDP":
                    objConfig.SetOption("Port", port)

                if has_session_log and logs_path:
                    objConfig.SetOption("Log Filename V2", "{0}".format(logs_path))

                objConfig.Save()
                session_counter += 1
    crt.Dialog.MessageBox("Process completed. Imported {0} sessions".format(session_counter), "Import process info")


import_mobaXterm_file()
