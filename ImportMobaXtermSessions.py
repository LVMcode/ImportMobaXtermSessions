# $language = "Python"
# $interface = "1.0"

import re
import datetime
import string

########################################################################################################
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

# Set has_personalized_session_name to True if you wish to have a personalized session_name
# rather than the imported session_name
has_personalized_session_name = False
# Your personalized_session_name will be created with the strings inside personalized_session_name_list
# joined with personalized_session_name_separator
# Valid personalized_session_name_separator strings are: "_", "-", "+", " ", "", ...
# personalized_session_name_list is a list of strings that contains keyword strings and
# static strings
# personalized_session_name_list keyword strings:
# "session_name"  --->   imported session_name
# "hostname"      --->   imported hostname
# "port"          --->   imported port
# "protocol"      --->   imported protocol
# personalized_session_name_list must include at least one of this keyword strings: "session_name"
# or "hostname"
# personalized_session_name_list static strings are any strings included in the list that
# aren't a keyword string
personalized_session_name_separator = "_"
personalized_session_name_list = ["session_name", "hostname"]
# For instance:
# personalized_session_name_separator = "--"
# personalized_session_name_list = ["mySession", "hostname", "session_name", "port", "was imported"]
# personalized_session_name could be "mySession--10.0.0.3--SW01--22--was imported"
########################################################################################################

protocol_dict = {
    "109": "SSH2",
    "98": "Telnet",
    "91": "RDP"
}


def check_personalized_session_name(personalized_session_name_list):
    if "session_name" in personalized_session_name_list or "hostname" in personalized_session_name_list:
        return True
    else:
        return False


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

                if has_personalized_session_name:
                    if check_personalized_session_name(personalized_session_name_list):
                        session_params_dict = {"session_name": session_name, "hostname": hostname, "port": str(port), "protocol": protocol}
                        personalized_list = []
                        for i in personalized_session_name_list:
                            element = session_params_dict[i] if i in session_params_dict.keys() else i
                            personalized_list.append(element)
                        session_name = "{0}".format(personalized_session_name_separator).join(personalized_list) \
                            if len(personalized_session_name_list) > 1 else personalized_list[0]
                        session_name = re.sub(r"(\\|/)+", r"", session_name)
                    else:
                        crt.Dialog.MessageBox("personalized_session_name_list must contain session_name or hostname "
                                              "keyword strings.\n "
                                              "personalized_session_name_list = {0}.\n"
                                              "Check your script GLOBAL OPTIONS section".format(personalized_session_name_list),
                                              "Error")
                        return

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
