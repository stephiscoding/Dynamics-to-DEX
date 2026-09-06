import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
import sys
import json

months = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# this .json file contains key-value pairs to convert CRM event names to short event codes
with open("event_codes.json", 'r') as f:
    event_codes = json.load(f)

# this fuction populates the events dict - data structure below:
#{Month:
#   {Case:
#	    [
#           OutletID, [
#		        (session, date, client ID),
#               (session, date, client ID),
#               ...
#           ]
#	    ]
#   }
#}
# the function interates through each row of the CRM export, grabbing the data it needs (date of session, the CRM case name, the ID of the client), and converting the CRM Case names into short-form case names (e.g. "IPPS Joondalup 2025-2026" becomes "IPPS-JOON")
# these case names then have a month and year appended to them (e.g. IPPS-JOON becomes IPPS-JOON-2601)
# it then places all of this data in the above "events" dictionary.
def process_excel(file_name):
    with pd.ExcelFile(file_name) as file:
        entries = pd.read_excel(file, file.sheet_names[0])

    events =  {}
    counters = {}
    for i in range(1,13):
        counters[i] = 1
    
    for index, row in entries.iterrows():
        # get the month for the entry
        session_date = row["Start Time (Event Session) (Event Session)"]
        month = datetime.date(session_date).month

        # convert the event to the event code, prompting the user for a case code if it does not yet exist
        # essentially, we want to let the user create whatever code they want for an event. as the same event might exist under seperate names in CRM, we let them add duplicates with a warning.
        related_event = row["Related Event"]
        new_code = ""
        while True:
            if related_event in event_codes:
                break
            elif new_code == "":
                first_new_code = input(f"{related_event} does not have a corresponding event code. Please enter one (e.g. IPPS-ARM): ")
                if first_new_code not in event_codes.values():
                    event_codes[related_event] = first_new_code
                    break
                new_code = first_new_code
            elif new_code in event_codes.values():
                new_code = input(f"{new_code} is already in use. Please enter one, or press enter to ignore: ")
                if new_code == "":
                    event_codes[related_event] = first_new_code
                    break
                elif new_code not in event_codes.values():
                    event_codes[related_event] = new_code
                    break
        
        # append the year and month to the event code
        case = event_codes[related_event] + "-" + datetime.strftime(session_date, "%y%m")

        # some more ugly code
        if month not in events:
            events[month] = {}
        if case not in events[month]:
            events[month][case] = [str(row["Outlet ID (Related Event) (Event)"]), []]
        
        events[month][case][1].append((
            f"IPPS-{datetime.strftime(session_date, "%y%m")}{counters[month]:03d}", 
            datetime.strftime(session_date, "%Y-%m-%d"), # converts raw excel date to required format for DEX
            str(row["Client ID (Member) (Client)"])
        ))
        counters[month] += 1
    # save the json file with the updated key-value pairs
    with open("event_codes.json", "w") as f:
        json.dump(event_codes, f, indent=4)
    return events

# DEPRECATED
def search_xml_for_value(search_value, root):
    for client in root.iter("ClientId"):
        if client.text == search_value:
            return True
    return False

# Pretty print XML with line breaks between adjacent elements
def prettify(elem, level=0):
    indent = "  "
    i = "\n" + level * indent
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + indent
        for idx, child in enumerate(elem):
            prettify(child, level + 1)
            if not child.tail or not child.tail.strip():
                # Add extra newline between specific adjacent tags
                if idx + 1 < len(elem):
                    next_tag = elem[idx + 1].tag
                    child.tail = i + indent
                else:
                    child.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

# if this function is not run on the XML output, DEX will not accept it. 
def fix_first_line(file):
    with open(file, 'r') as f:
        xml = f.read()
    xml_fixed = xml.replace(" encoding='utf-8'", "")
    with open(file, 'w') as f:
        f.write(xml_fixed)

print("Digesting Excel Spreadsheet...")
cases = process_excel(sys.argv[1])

# create an XML file for each month in the CRM export
for month in cases:
    root = ET.Element("DEXFileUpload", {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xmlns:xsd": "http://www.w3.org/2001/XMLSchema"
    })
    print(f"Creating XML file for {months[month]}...")
    case_count = 0
    print("Creating Cases...")
    # Create and populate Cases element
    # each case should contain all of the unique clients who attended the sessions
    cases_elem = ET.SubElement(root, "Cases")
    for case in cases[month]:
        case_elem = ET.SubElement(cases_elem, "Case")
        ET.SubElement(case_elem, "CaseId").text = str(case)
        ET.SubElement(case_elem, "OutletActivityId").text = cases[month][case][0]
        ET.SubElement(case_elem, "TotalNumberOfUnidentifiedClients").text = "0"
        case_count += 1
        
        # begin adding clients to the case
        client_case_count = 0
        seen_clients = set()
        case_clients_elem = ET.SubElement(case_elem, "CaseClients")
        for session, date, client in cases[month][case][1]:
            # if the client is not already in of the case, add them. otherwise, go to the next client
            if client not in seen_clients:
                seen_clients.add(client)
                case_client_elem = ET.SubElement(case_clients_elem, "CaseClient")
                ET.SubElement(case_client_elem, "ClientId").text = client
                ET.SubElement(case_client_elem, "ReferralSourceCode").text = "INTERNAL"
                reasons_elem = ET.SubElement(case_client_elem, "ReasonsForAssistance")
                reason_elem = ET.SubElement(reasons_elem, "ReasonForAssistance")
                ET.SubElement(reason_elem, "ReasonForAssistanceCode").text = "FAMILY"
                ET.SubElement(reason_elem, "IsPrimary").text = "true"
                client_case_count += 1
        print(f"Added {client_case_count} unique clients to {case}.")
    print(f"Created {case_count} cases.")

    print("Creating Sessions...")

    case_count = 0
    session_client_count = 0

    # Add Sessions
    sessions_elem = ET.SubElement(root, "Sessions")
    for case in cases[month]:
        session_client_count = 0
        for session, date, client in cases[month][case][1]:
            session_elem = ET.SubElement(sessions_elem, "Session")
            ET.SubElement(session_elem, "SessionId").text = session
            ET.SubElement(session_elem, "CaseId").text = case
            ET.SubElement(session_elem, "SessionDate").text = date
            ET.SubElement(session_elem, "ServiceTypeId").text = "12"
            ET.SubElement(session_elem, "TotalNumberOfUnidentifiedClients").text = "0"
            ET.SubElement(session_elem, "InterpreterPresent").text = "false"
            ET.SubElement(session_elem, "ServiceSettingCode").text = "ORGOUTLETOFFICE"
            session_clients_elem = ET.SubElement(session_elem, "SessionClients")
            session_client_elem = ET.SubElement(session_clients_elem, "SessionClient")
            ET.SubElement(session_client_elem, "ClientId").text = client
            ET.SubElement(session_client_elem, "ParticipationCode").text = "CLIENT"
            session_client_count += 1
        print(f"Created {session_client_count} sessions for {case}")

    print("Cleaning up XML...")
    prettify(root)
    tree = ET.ElementTree(root)
    tree.write(f"{months[month]} cases.xml", encoding="utf-8", xml_declaration=True)
    fix_first_line(f"{months[month]} cases.xml")
    print(f"Created file '{months[month]} cases.xml'!")
    
