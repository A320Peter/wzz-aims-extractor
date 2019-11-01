import requests
import base64
import hashlib
import sys, getopt
from bs4 import BeautifulSoup 
import getpass
import unicodedata
import calendar
import time
from datetime import datetime
from datetime import date, timedelta
import math
import re
from os import system, name 
import csv
import os


class Crewmember(object):
    name: str
    role: str

class Sector(object):
    date: int       
    flightnum: str     
    from_: str     
    to: str    
    off: str          
    on: str             
    total: str
    reg: str            
    type_: str
    otherpilot: str
    PIC_name: str
    PIC: str
    COPLT: str
    pf_to_day: str
    pf_to_night: str
    pf_ldg_day: str
    pf_ldg_night: str
    SIM_time: str
    SIM_note: str
    err_flag: str


# ANSI character codes for colors
class bcolors:
    HEADER = '\033[1;32;40m'
    OKBLUE = '\033[1;36;40m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'



# handling exceptions
def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    print("")
    print(">>>>> ERROR <<<<<")
    print("")
    print("An error has occured. Please try to download your roster for different")
    print("time periods in this year or last year. Try with 10 days or more.") 
    print("")
    print("Please copy paste and send the alien error message to the developer mentioning") 
    print("that the roster download is possible for different periods or not possible at all.")
    print("")
    input("Press enter to exit.")
    sys.exit(-1)


# define our clear function 
def clear(): 
  
    # for windows 
    if name == 'nt': 
        _ = system('cls') 
  
    # for mac and linux(here, os.name is 'posix') 
    else: 
        _ = system('clear')

# validate date entry string
def validate(date_text) -> bool:
    try:
        datetime.strptime(date_text, '%d/%m/%Y')
        return True
    except:
        return False


def _check_response(r: requests.Response, *args, **kwargs) -> None:
    # Checks the response from a request
    print(".", end="", flush=True)



# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 0, length = 100, fill = '█'):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = '\r', flush=True)
    # Print New Line on Complete
    if iteration == total: 
        print()


# =============================================================
# Version check upon starup
# return False if there is no new version
# =============================================================

def new_version(version: float) -> bool:


    session = requests.Session()
    session.headers.update({
        "User-Agent":
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:64.0) "
        "Gecko/20100101 Firefox/64.0"})
    print("")
    print("Checking version please wait...")

    r = session.get("https://raw.githubusercontent.com/A320Peter/wzz-aims-extractor/master/version", timeout=20)
  
    if float(r.text) > version:
        return True
    else:
        return False



# =============================================================
# Connects to AIMS creating a global session
# =============================================================

def connect(username:str, password:str, start_date:str) -> bool:

    global _session, _aims_url, _full_roster_link
    _session = requests.Session()
    _session.hooks['response'].append(_check_response)
    _session.headers.update({
        "User-Agent":
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:64.0) "
        "Gecko/20100101 Firefox/64.0"})
    print("Connecting..", end="", flush=True)

    # send login request
    r = _session.post("https://ecrew.wizzair.com/wtouch/wtouch.exe/verify",
               {"Crew_Id": base64.b64encode(username.encode()),
                "Crm": hashlib.md5(password.encode()).hexdigest()},
                      timeout=20)

    print(".", end="", flush=True)
    
    # security
    del password
    
    if r.text.find("Login was unsuccessful") != -1:
        print(". [FAIL] Invalid User/pass or authentication error.")
        return False
    else:
        print(". [SUCCESSFUL] Please wait...", end="", flush=True)


    # check if there is any changes
    r = _session.get("https://ecrew.wizzair.com/wtouch/perinfo.exe/index", timeout=20)
    if r.text.find("You have changes") != -1:
        print("\n=== ERROR ===")
        print("You have changes.")
        print("Disconnected.")
        return False

    return True





# ================================================================================================
# Gets the roster between two dates and returns HTML code as text list. One roster is one element.
# ================================================================================================
def download_roster(start_date:str, end_date:str) -> [str]:

    # array
    roster_array = []
    roster_array.clear()

    # calculate the number of request to be made. AIMS limit: 32 days / roster
    d0 = datetime.strptime(start_date, "%d/%m/%Y")
    d1 = datetime.strptime(end_date, "%d/%m/%Y")
    delta = d1 - d0
    num_of_requests = math.ceil(delta.days / 32)

    #print(num_of_requests)

    actual_date = d0
    printProgressBar(1, num_of_requests+1, prefix = 'Downloading Roster...', suffix = ' ', length = 50)
    for i in range(1, num_of_requests+1):
        if i == num_of_requests:
            last = d1 - actual_date # we only need a few days in the last round, not 32
            days = str(last.days + 1)
        else:
            days = "32"

        date = actual_date.strftime('%d/%m/%Y')  # "07/01/2018"   dd/mm/yyyy
        #print(date)
        time_format = "2"  # 2 = UTC
        
        r = _session.get("https://ecrew.wizzair.com/wtouch/perinfo.exe/crwsche?cal1=" + date + "&nDays=" + days + "&times_format=" + time_format, timeout=20)
        crwsche_page = BeautifulSoup(r.text, 'html.parser')
        roster_link = crwsche_page.frame.next_sibling.next_sibling.get("src")
        _full_roster_link = "https://ecrew.wizzair.com" + roster_link
        
        #print(_full_roster_link)
    
        r = _session.get(_full_roster_link, timeout=20)
        roster_array.append(date + r.text) # we encode the start date before html code for easier roster parsing

        actual_date = actual_date + timedelta(days=32)

        printProgressBar(i, num_of_requests, prefix = 'Downloading Roster...', suffix = ' ', length = 50)

    return roster_array



# =======================================================================
# Gets the HTML logbook between two dates and returns HTML code as string
# =======================================================================
def download_logbook(start_date:str, end_date:str) -> str:
    
    printProgressBar(1, 3, prefix = 'Downloading Logbook...', suffix = ' ', length = 50)

    r = _session.get("https://ecrew.wizzair.com/wtouch/perinfo.exe/pilotlogbook?cal1=" + start_date + "&cal2=" + end_date, timeout=180)
    log_page = BeautifulSoup(r.text, 'html.parser')
    logbook_link = log_page.frame.next_sibling.next_sibling.get("src")
    _full_logbook_link = "https://ecrew.wizzair.com" + logbook_link

    printProgressBar(2, 3, prefix = 'Downloading Logbook...', suffix = ' ', length = 50)

    r = _session.get(_full_logbook_link, timeout=20)

    printProgressBar(3, 3, prefix = 'Downloading Logbook...', suffix = ' ', length = 50)

    return r.text






# =============================================================
# Parsing the HTML logbook and returns a Sector type list
# =============================================================
def parse_logbook(html_doc:str) -> [Sector]:

    html_doc.replace("\n</tr><tr class=", "\n<tr class=")
    soup = BeautifulSoup(html_doc, 'html.parser')

    logbook_table = soup.find_all("div", attrs={'style': re.compile(r'.font-family:Arial.')})

    logbook = [Sector]
    logbook.clear()

    # normalize strings
    for i in range(0, len(logbook_table)):
        if logbook_table[i].string != None:
            logbook_table[i].string = unicodedata.normalize('NFKD', logbook_table[i].string) # remove strange characters :)
            logbook_table[i].string = re.sub(' +', ' ', logbook_table[i].string)
            logbook_table[i].string = logbook_table[i].string.strip() # remove whitespace
            #print(logbook_table[i].string)

    for i in range(3, len(logbook_table)):
        if (logbook_table[i-1].string.find("Night") != -1) and \
            (logbook_table[i].string.find("/") != -1) and \
            (logbook_table[i].string != ""): # new table is starting

            pointer = i + 666
            if pointer > len(logbook_table):
                pointer = len(logbook_table)-18

            for j in range(i, pointer, 18):
                if logbook_table[j].string.find("Totals") != -1: #end of logbook table break anyway
                    break
                logbook_entry = Sector()
                logbook_entry.date = calendar.timegm(time.strptime(logbook_table[j].string, '%d/%m/%y'))
                logbook_entry.from_ = logbook_table[j+1].string
                logbook_entry.to = logbook_table[j+3].string
                logbook_entry.off = logbook_table[j+2].string
                logbook_entry.on = logbook_table[j+4].string
                logbook_entry.type_ = logbook_table[j+5].string
                logbook_entry.reg = logbook_table[j+6].string
                logbook_entry.total = logbook_table[j+7].string
                logbook_entry.PIC_name = logbook_table[j+8].string
                
                if logbook_table[j+9].string == "":
                    logbook_entry.pf_ldg_day = "0"
                else:
                    logbook_entry.pf_ldg_day = logbook_table[j+9].string
                
                if logbook_table[j+10].string == "":
                    logbook_entry.pf_ldg_night = "0"
                else:
                    logbook_entry.pf_ldg_night = logbook_table[j+10].string

                logbook_entry.PIC = logbook_table[j+11].string
                logbook_entry.COPLT = logbook_table[j+12].string

                if logbook_table[j+16].string == "":
                    logbook_entry.pf_to_day = "0"
                else:
                    logbook_entry.pf_to_day = logbook_table[j+16].string

                if logbook_table[j+17].string == "":
                    logbook_entry.pf_to_night = "0"
                else:
                    logbook_entry.pf_to_night = logbook_table[j+17].string

                logbook_entry.SIM_time = logbook_table[j+14].string
                logbook_entry.SIM_note = logbook_table[j+15].string

                if logbook_entry.SIM_time != "":
                    logbook_entry.from_ = ""
                    logbook_entry.to = ""

                # time correction 0:37 -> 00:37
                if len(logbook_entry.off) < 5:
                    logbook_entry.off = "0" + logbook_entry.off
                if len(logbook_entry.on) < 5:
                    logbook_entry.on = "0" + logbook_entry.on

                #empty unknown fields
                logbook_entry.flightnum = ""
                logbook_entry.otherpilot = ""
                logbook_entry.err_flag = ""


                logbook.append(logbook_entry)
                del logbook_entry

    return logbook







# =============================================================
# Parsing the HTML roster and returns a Sector type list
# =============================================================
def parse_roster(html_doc:str, start_date:str) -> [Sector]:

    html_doc.replace("\n</tr><tr class=", "\n<tr class=")
    soup = BeautifulSoup(html_doc, 'html.parser')

    # start date (required to find crew names)
    # UTC date to UTC timestamp
    utc_start_date = calendar.timegm(time.strptime(start_date, '%d/%m/%Y'))

    crew_list_soup = soup.find_all(string=re.compile(">\xa0"))
    temp = ""
    for x in range(len(crew_list_soup)):
        temp = temp + crew_list_soup[x]

    crew_list_string = unicodedata.normalize('NFKD', temp)
    crew_list_string = re.sub("\s\s+" , " ", crew_list_string) # remove double or more whitespaces


    # each roster block is 11px high so we are looking for all of these table cells
    roster_table = soup.find_all("tr", style="height:11px")

    day = 0 # the number of the day (column) we are interested in the roster table (0 = first column)
    AimsDuty = [] # adding roster columns (days) to this list. Each element represents a day.
    AimsDuty.clear()



    # go through each day and each line. Load everything into AimsDuty list
    Err = False
    while day < len(roster_table[0].find_all("td")): # number of columns in the roster table (always 32?)
        row = []
        row.clear()
        for x in range(0, len(roster_table)): # number of rows in the roster table
            temp = roster_table[x].find_all("td")[day].string
            try:
                new_str = unicodedata.normalize('NFKD', temp)    # fix for: https://stackoverflow.com/questions/10993612/python-removing-xa0-from-string
                new_str = re.sub(r"\s+", "", new_str, flags=re.UNICODE) # remove whitespace
            except:
                Err = True
                new_str = ""
            row.append(new_str)
        day += 1
        AimsDuty.append(row)
    if Err:
        print("ERROR: at least one exception has occured when normalizing string. Probably there is no data loss.")

        

    #print(AimsDuty[7][0])
    AimsDay = [Sector]
    AimsDay.clear()

    # it's possible that duties shifting to next day and we need to iterate the loop
    flag = ""

    # Analyze every row one by one, each day and try to guess what kind of duty is that
    for i in range(0, len(AimsDuty)):
        for j in range(0, len(AimsDuty[i])):
            
            if flag != "": # if the duty doesn't fit on one day, there is a broken duty from the previous loop
                if flag == "off":
                    AimsDay[-1].off = AimsDuty[i][j]
                    AimsDay[-1].from_ = AimsDuty[i][j+1]
                    AimsDay[-1].to = AimsDuty[i][j+2]
                    AimsDay[-1].on = AimsDuty[i][j+3]
                if flag == "des":
                    try:
                        AimsDay[-1].to = AimsDuty[i][j]
                        AimsDay[-1].on = AimsDuty[i][j+1]
                    except:
                        print(i)
                        print(j)
                        print(len(AimsDuty[i]))
                flag = ""

            # we expect '2201' or '2201A' flight number formats, then it's a flight duty
            if AimsDuty[i][j][:4].isnumeric():
                if ":" in AimsDuty[i][j+2]: #first duty with check in time
                    off = AimsDuty[i][j+2]
                    on = AimsDuty[i][j+5]
                    dep = AimsDuty[i][j+3]
                    des = AimsDuty[i][j+4]
                else:
                    off = AimsDuty[i][j+1]
                    on = AimsDuty[i][j+4]
                    dep = AimsDuty[i][j+2]
                    des = AimsDuty[i][j+3]


                flightnum = AimsDuty[i][j]
                
                #if destination is too short or too long, it's not valid, probably broken to next day
                if len(des) < 3 or len(des) > 3: # in case the sector is broken to the next day
                    flag = "des"
                elif off.find(":") == -1:
                    flag = "off"

                # in the crew list sometimes there is no '2201A' but only '2201' in case of diversion, etc.
                flightnum_no_suffix = flightnum[:4]

                # get the crew names
                FoundCrew = False
                err = ""
                ts = int(utc_start_date) + (i * 86400)
                duty_date = datetime.utcfromtimestamp(ts).strftime('%d/%m/%Y')
                date_index = crew_list_string.find(duty_date)
                count = 0 
                while FoundCrew == False:

                    count += 1
                    sub_index1 = crew_list_string.find('>', date_index)-2
                    flight_numbers = crew_list_string[date_index:sub_index1]
                    
                    #print("looking for: " + flightnum_no_suffix + " or ALL")
                    #print("found: " + flight_numbers)
                    #print("crew_list_string: " + crew_list_string)
                    #print("count: " + str(count))

                    if flight_numbers.find(flightnum_no_suffix) != -1 or flight_numbers.find("All") != -1:
                        
                        sub_index2_a = crew_list_string.find('>', sub_index1+3)
                        sub_index2_b = crew_list_string.find('/2', sub_index1+3)    
                        if sub_index2_a < sub_index2_b or sub_index2_b == -1: # in case the crew member is the last on the day
                            sub_index2 = sub_index2_a - 3
                        else:
                            sub_index2 = sub_index2_b - 5

                        other_crew = crew_list_string[sub_index1+4:sub_index2]

                        FoundCrew = True

                        # check for redundant same role crew member and put a flag indicating possible mismatch.
                        # it will only affect captains roster as for FOs the PIC name is included in the logbook table
                        search_string = crew_list_string[sub_index1:sub_index2_b]
                        if (search_string.count("SF>") > 1) or (search_string.count("CP>") > 1) or \
                        ((search_string.count("SF>") > 0) and (search_string.count("CP>") > 0)):
                            err = "Verify the other pilot's name on this flight."

                    # in some exceptional cases (like ferry flights on consequtive days with different time zones) in the crew list the flight number might not exist
                    # for the actual date. If this happens after several unsuccessful loops, let's check for the flight number on the prev. day to avoid infinite loop
                    elif count > 10:
                        ts2 = int(utc_start_date) + (i * 86400) - 86400
                        duty_date = datetime.utcfromtimestamp(ts2).strftime('%d/%m/%Y')
                        date_index = crew_list_string.find(duty_date)
                        count = 0


                    else:
                        date_index = crew_list_string.find(duty_date, date_index+1)



                actual_sector = Sector()
                if off == "24:00": # bug fix for AIMS Logbook. 24:00 is the previous day instead of turning the clock to 00:00+1
                    ts = ts + 86400 # + 1 day
                    off = "00:00"
                actual_sector.date = ts
                actual_sector.flightnum = flightnum
                actual_sector.from_ = dep
                actual_sector.to = des
                actual_sector.off = off
                actual_sector.on = on
                actual_sector.reg = ""
                actual_sector.otherpilot = other_crew
                actual_sector.err_flag = err

                if dep.find("*") == -1: # in case of deadheading sector do not add this sector to the list
                    AimsDay.append(actual_sector)
                else:
                    flag = "" # delete the flag if there is any and if it's a deadheading sector

                del actual_sector
                if flag != "":
                    break
    return AimsDay




def delete_last_console_lines(n=1):
    CURSOR_UP_ONE = '\x1b[1A'
    ERASE_LINE = '\x1b[2K'
    for _ in range(n):
        sys.stdout.write(CURSOR_UP_ONE)
        sys.stdout.write(ERASE_LINE)



def main():

    # app version
    version = 0.5

    # clear console
    clear()


    # handling command line arguments
    # -f <filename>     /parsing a local HTML file.
    inputfile = ''
    if len(sys.argv) > 2:
        if sys.argv[1] == "-f":
            inputfile = sys.argv[2]
            if not os.path.isfile('./' + inputfile):
                print("File does not exist.")
                print("")
                input("Press any key to exit.")
                sys.exit(0)
        else:
            print("Available arguments:")
            print("")
            print("-f <filename>")
            print("Parsing a local HTML file.")
            print("")
            input("Press any key to exit.")
            sys.exit(2)

    if inputfile != '':
        print(inputfile + " loaded.")


    if new_version(version):
        print("")
        print(" WiZz AIMS eXtractor v" + str(version))
        print(" =====================================================================")
        print(" | *** This version is out of date. Please download the latest at:   |")
        print(" |                                                                   |")
        print(" | ** Mac or Linux:                                                  |")
        print(" |       https://bit.ly/2IX2Pjv                                      |")
        print(" |                                                                   |")
        print(" | ** Windows:                                                       |")
        print(" |       https://bit.ly/2UYs8sB                                      |")
        print(" |                                                                   |")
        print(" | *** Project website:                                              |")
        print(" |       https://github.com/A320Peter/wzz-aims-extractor             |")
        print(" =====================================================================")
        print(" ")
        input("Press any key to exit.")
        sys.exit(0)


    delete_last_console_lines(1)
    print(" ===========================================================================v"+str(version), flush=True)
    print(" |                       ***  WizzAir AIMS eXtractor  ***                      |")
    print(" ===============================================================================")
    print("")
    print(" This tool is intended to download your AIMS roster and convert into")
    print(" a CSV file that can be imported into Logbook applications like mccPILOTLOG.")
    print("")
    print(" If you are not a WizzAir employee, you should close this application now.")
    print("")
    print(" Your ID and password is not stored anywhere and are encrypted. You should")
    print(" use this application only on your personal computer to avoid violation of")
    print(" company's IT policy.")
    print("")
    print(" The application will terminate in case of active changes in your roster")
    print(" preventing accidentally accepting any notification.")
    print("")
    print("")
    print(" Enter the desired dates for logbook download. Format: DD/MM/YYYY")
    print("")

    # validate date entries
    dt_start = input("             Start date: ")
    while not validate(dt_start):
        delete_last_console_lines(1)
        dt_start = input("   invalid   Start date: ")
    delete_last_console_lines(1)
    print("             Start date: "+dt_start)

    dt_end = input("               End date: ")
    while not validate(dt_end):
        delete_last_console_lines(1)
        dt_end = input("   invalid     End date: ")
    delete_last_console_lines(1)
    print("               End date: "+dt_end)

    while datetime.strptime(dt_end, '%d/%m/%Y') < datetime.strptime(dt_start, '%d/%m/%Y'):
        delete_last_console_lines(1)
        dt_end = input("start < end !  End date: ")
    delete_last_console_lines(1)
    print("               End date: "+dt_end)


    # we don't need authentication if it's a local file
    if inputfile == '':
        print("")
        userID = input("   Enter your ID number: ")
        print("")
        pwd = getpass.getpass('          AIMS password: ')
        delete_last_console_lines(1)
        print('          AIMS password: ******')
        print("")
    else:
        userID = 'template'
        pwd = 'template'


    
    AimsRoster = []
    AimsRoster.clear()
    logbook = []
    logbook.clear()

    if connect(userID, pwd, dt_start) or inputfile != '':


        # we either download a roster or loading it from a file
        if inputfile == '':
            roster_html = download_roster(dt_start, dt_end)
        else:
            roster_html = []
            roster_html.clear()
            with open(inputfile, 'r') as reader:
                roster_html.append(dt_start + reader.read())



        # parse the HTML roster into AimsRoster sectors.
        for x in range(len(roster_html)): 
            AimsRoster.extend(parse_roster(roster_html[x], roster_html[x][:10])) # start date is the first 10 chars
        
        
        # using local file for debug:

        if inputfile != '':
            # Print out all the duties line by line.
            for y in range(len(AimsRoster)):
                print(datetime.utcfromtimestamp(AimsRoster[y].date).strftime('%d/%m/%Y') + "   " + AimsRoster[y].flightnum + "    " + AimsRoster[y].from_ + "->" + AimsRoster[y].to + "    " + AimsRoster[y].off + " - " + AimsRoster[y].on + "    " + AimsRoster[y].reg + "    " + AimsRoster[y].otherpilot)
            print("")
            print("<END OF DUTIES>")
            sys.exit(0)


        # download and parse the logbook
        logbook_html = download_logbook(dt_start, dt_end)
        logbook = parse_logbook(logbook_html)

        # cycle through the logbook and complete all information from AimsRoster list
        for i in range(0, len(logbook)):

            # search for the same record in AimsRoster
            for j in range(0, len(AimsRoster)):
                # we compare all the data we have to minimize mismatch possibility
                if logbook[i].date == AimsRoster[j].date and \
                logbook[i].from_ == AimsRoster[j].from_ and \
                logbook[i].to == AimsRoster[j].to and \
                logbook[i].off == AimsRoster[j].off and \
                logbook[i].on == AimsRoster[j].on:

                    # we found the same entry
                    logbook[i].flightnum = AimsRoster[j].flightnum
                    logbook[i].otherpilot = AimsRoster[j].otherpilot
                    logbook[i].err_flag = AimsRoster[j].err_flag


        """
        #print out all logbook entry
        for i in range(0, len(logbook)):
            print(datetime.utcfromtimestamp(logbook[i].date).strftime('%d/%m/%Y') + "  " + 
                logbook[i].flightnum + "  " +
                logbook[i].from_ + " -> " + 
                logbook[i].to + "    " +
                logbook[i].off + " - " +
                logbook[i].on + "  " +
                logbook[i].type_ + " " +
                logbook[i].reg + "   " + 
                logbook[i].total + "     " +
                logbook[i].PIC_name + "     " + logbook[i].otherpilot + "    " +
                logbook[i].pf_to_day + " " + logbook[i].pf_to_night + " " + logbook[i].pf_ldg_day + " " + logbook[i].pf_ldg_night + "   " +
                logbook[i].SIM_time + "  " + logbook[i].SIM_note)
        """

        # ask for output format
        print("")
        print(" 1)  mccPILOTLOG")
        print(" 2)  LogTen Pro")
        print("")
        output = input(" Choose output format (1 or 2): ")
        while output != "1" and output != "2":
            delete_last_console_lines(1)
            output = input(" Choose output format (1 or 2): ")
        delete_last_console_lines(1)

        delete_last_console_lines(3)
        # ask for flight number format
        print(" 1)  W6 2201")
        print(" 2)   W62201")
        print(" 3)     2201")
        print("")
        fnum_format = input(" Choose flight number format (1 or 2 or 3): ")
        while fnum_format != "1" and fnum_format != "2" and fnum_format != "3":
            delete_last_console_lines(1)
            fnum_format = input(" Choose output format (1 or 2): ")
        print("")

        # Write the logbook
        cautions = 0

        if output == "1":
            logbook_filename = '/Logbook_mccPILOTLOG_'+datetime.today().strftime('%d-%m-%Y_%H.%M.%S')+'.csv' # mccPILOTLOG filename
        else:
            logbook_filename = '/Logbook_LogTenPro_'+datetime.today().strftime('%d-%m-%Y_%H.%M.%S')+'.csv' # LogTen filename

        with open(os.path.expanduser("~/Desktop")+logbook_filename, mode='w') as logbook_file:
            logbook_writer = csv.writer(logbook_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)

            if output == "1":
                logbook_writer.writerow(['MCC_DATE', 'FLIGHTNUMBER', 'AF_DEP', 'AF_ARR', 'TIME_DEP', 'TIME_ARR', 'AC_MODEL', 'AC_REG', 'TIME_TOTAL', 'TIME_PIC', 'TIME_IFR', 'AC_ISSIM', 'CAPACITY', 'TO_DAY', 'TO_NIGHT', 'LDG_DAY', 'LDG_NIGHT', 'PF', 'PILOT1_NAME', 'PILOT2_NAME', 'REMARKS', 'X'])
            else:
                logbook_writer.writerow(['flight_flightDate', 'flight_flightNumber', 'flight_from', 'flight_to', 'flight_actualDepartureTime', 'flight_actualArrivalTime', 'aircraft_secondaryID', 'aircraft_aircraftID', 'flight_totalTime', 'flight_pic', 'flight_ifr', 'flight_simulator', 'flight_picCapacity', 'flight_dayTakeoffs', 'flight_nightTakeoffs', 'flight_dayLandings', 'flight_nightLandings', 'flight_pilotFlyingCapacity', 'flight_selectedCrewPIC', 'flight_selectedCrewSIC', 'flight_remarks', 'flight_flagged'])

            for i in range(0, len(logbook)):
                
                total_time = logbook[i].total

                # am I PIC or FO?
                if output == "1":
                    if logbook[i].PIC_name == "Self" or logbook[i].PIC_name == "":
                        TIME_PIC = logbook[i].total
                        CAPACITY = "PIC"
                        PILOT1 = "Self"
                        PILOT2 = logbook[i].otherpilot # from roster
                    else:
                        TIME_PIC = ""
                        CAPACITY = "Co-Pilot"
                        PILOT1 = logbook[i].PIC_name # from aims logbook
                        PILOT2 = "Self"
                        logbook[i].err_flag = "" # PIC name is accurate so we can delete the warning flag
                else:
                    if logbook[i].PIC_name == "Self" or logbook[i].PIC_name == "":
                        TIME_PIC = logbook[i].total
                        CAPACITY = "1"
                        PILOT1 = "Self"
                        PILOT2 = logbook[i].otherpilot # from roster
                    else:
                        TIME_PIC = ""
                        CAPACITY = "0"
                        PILOT1 = logbook[i].PIC_name # from aims logbook
                        PILOT2 = "Self"
                        logbook[i].err_flag = "" # PIC name is accurate so we can delete the warning flag

                # is it SIM or real time?
                if output == "1":
                    if logbook[i].SIM_note != "":
                        SIM = "True"
                        total_time = logbook[i].SIM_time
                        FNUM = ""
                    else:
                        SIM = "False"
                        FNUM = logbook[i].flightnum
                else:
                    if logbook[i].SIM_note != "":
                        SIM = logbook[i].SIM_time
                        total_time = ""
                        FNUM = "sim"
                    else:
                        SIM = "False"
                        FNUM = logbook[i].flightnum

                # format the flight number
                if fnum_format == "1" and FNUM != "sim":
                    FNUM = "W6 " + FNUM
                elif fnum_format == "2" and FNUM != "sim":
                    FNUM = "W6" + FNUM

                # pilot capacity
                if logbook[i].pf_to_day == "1" or logbook[i].pf_to_night == "1" or logbook[i].pf_ldg_day == "1" or logbook[i].pf_ldg_night == "1":
                    PF_capacity = "1"
                else:
                    PF_capacity = "0"

                # AC model. For LogTen we only use the registration to avoid conflicts
                if output == "1":
                    aircraft_model = "A"+logbook[i].type_
                else:
                    aircraft_model = ""

                # if flight is flagged with error put a marker flag for LogTen Pro
                if logbook[i].err_flag != "":
                    flag_marker = "1"
                else:
                    flag_marker = "0"

                logbook_writer.writerow([datetime.utcfromtimestamp(logbook[i].date).strftime('%d/%m/%Y'),
                 FNUM,
                 logbook[i].from_,
                 logbook[i].to,
                 logbook[i].off,
                 logbook[i].on,
                 aircraft_model,
                 logbook[i].reg,
                 total_time,
                 TIME_PIC,
                 logbook[i].total,
                 SIM,
                 CAPACITY,
                 logbook[i].pf_to_day,
                 logbook[i].pf_to_night,
                 logbook[i].pf_ldg_day,
                 logbook[i].pf_ldg_night,
                 PF_capacity,
                 PILOT1,
                 PILOT2,
                 logbook[i].SIM_note,
                 flag_marker])

                if logbook[i].err_flag != "":
                    if cautions == 0: # we want to print 'cautions' title only once
                        print(" !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                        print(" CAUTIONs:")
                        print(" !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    if output == "1":
                        print(datetime.utcfromtimestamp(logbook[i].date).strftime('%d/%m/%Y') + "  " + FNUM + "  " + logbook[i].from_ + "->" + logbook[i].to + "  CAUTION: " + logbook[i].err_flag)
                    elif output == "2" and cautions == 0:
                        print(" Some flights will be red flagged after imoport. Please verify the other pilot's name on those flights as AIMS data is confusing.")
                    cautions += 1

                # if there is no flight number and not SIM, the flight wasn't fetched from the roster probably there is an error
                if logbook[i].flightnum == "" and SIM == "False":
                    if cautions == 0: # we want to print 'cautions' title only once
                        print("")
                        print("(!) CAUTIONS:")
                        print("")
                    print(datetime.utcfromtimestamp(logbook[i].date).strftime('%d/%m/%Y') + "  ????  " + logbook[i].from_ + "->" + logbook[i].to + "  CAUTION: Please check this flight. Something is not OK. <------------")
                    cautions += 1


        if output == "1":
            print("")
            print(" Note:")
            print(" Night time is not automatically calculated in the generated logbook as")
            print(" AIMS doesn't have this feature. To calculate proper night hours do the following:")
            print(" after importing flights into mccPILOTLOG: select all the imported flights at the same time")
            print(" and tick the box at the bottom panel: 'Re-calculate Night Time' then click on Execute button.")
            print("")
            print("*** COMPLETED with " + str(cautions) + " cautions. ***")
            print("")
            print("Logbook has been saved to: " + os.path.expanduser("~/Desktop")+logbook_filename)
            print("")
        else:
            print("")
            print(" Note:")
            print(" Night time is not automatically calculated in the generated logbook as")
            print(" AIMS doesn't have this feature. To calculate proper night hours do the following:")
            print(" after importing flights into LogTen Pro: select all the imported flights at the same time")
            print(" and select 'Night Time' text field at the right panel. Tap space to autofill.")
            print("")
            print("*** COMPLETED with " + str(cautions) + " cautions. ***")
            print("")
            print("Logbook has been saved to: " + os.path.expanduser("~/Desktop")+logbook_filename)
            print("")

        sys.exit(0)




if __name__ == "__main__":
    sys.excepthook = show_exception_and_exit
    main()
