"""imports Google calendar data and convert and export into format for spreadsheets.

to be used for invoices tax returns etc
Sadly what is missing in the Google calendar expert are the colours
So it might not be able to identify red canceled sessions
or sessions that are yellow maybe
"""
import datetime

now = datetime.datetime.now()
mydate = datetime.datetime(1,1,1)
mydate = mydate.replace(year=2023)
mydate = mydate + datetime.timedelta(weeks=1)
print(mydate)

# make sure you check the file names before running
# it will overwrite existing data so be careful

# the calendar file as exported from Google calendar
calendar_file = open("calendarData.txt", "r")
# the output file which will have the data in a 
# spreadsheet friendly format
output_file = open("calendarXL.csv", "w")

# each event in the calendar
events = []

# while there are other categories, these are the ones we are intersted in
categories = ["name","start","end","stop repeat","exclusions"]

# initially we got the program to skip past the start
# and only run when it sees the first BEGIN:VEVENT
# but in the end it was not needed
"""
def pass_header():
    while True:
        line = calendar_file.readline()
        if line == "BEGIN:VEVENT":
            break
"""

def find_number_start(line):
    for pos, val in enumerate(line):
        if val.isdecimal():
            start_pos = pos
            return pos 

def read_date_time(line):
    """Get one date or date and time value
        
    Assume that the date is the first number on the line
    Assume it is the standard format of 
    4 digit year number
    2 digit month number
    2 digit day date number
    If an all day event it will finish with this
    If it has a time it will follow straight after
    Start with the letter T
    2 digit hour
    2 digit minute
    2 digit second
    """
    # position on the line where I should read the data
    start_pos = find_number_start(line)

    # all dates and times should have numbers
    # so check if no numbers are present
    # when you do get to a number assume it is a date
    if not start_pos:
        print("************** there are no numbers in this line! ************")
        print(line)
        return
    
    date_time={}
    # we assume it is in the format as described above
    # so we know the place of year, month, day, hours, min
    # if not raise an exception
    try:
        date_time["year"] = int(line[start_pos:start_pos+4])
        date_time["month"] = int(line[start_pos+4:start_pos+6])
        date_time["day_in_month"] = int(line[start_pos+6:start_pos+8])
        # all day events dont have a time so will have a shorther length
        if len(line)>=start_pos+10:
            # skip one character in string for the letter T for time
            date_time["hour"] = int(line[start_pos+9:start_pos+11])
            date_time["mins"] = int(line[start_pos+11:start_pos+13])
    except(ValueError):
        print("************* Exception: not a number! ******************")
        print(line)
        date_time = {}
    return date_time

def get_freq(line):
    """
        Get the fequency, ie weekly, yearly etc. 

        It should start with RRULE:FREQ=
        The next item could be WEEKLY, YEARLY, DAILY 
        The rest of the line is the date/ time 
        This function only selects the WEEKLY, YEARLY, DAILY bit
    """
    # start of line is RRULE:FREQ= which is 11 characters
    start_pos = 11
    if len(line) < start_pos:
        print("********** line too short in get_freq *******************")
        return
    start_char = line[start_pos]
    if start_char == "W":
        return "weekly"
    if start_char == "D":
        return "daily"
    if start_char == "Y":
        return "yearly"
    print("******************* do not know the frequency in get_freq *******************")
    return 

def get_one_data():
    """read one item from the calendar

    Can't indetify item by position as it is inconsistant
    so we look for an identifier
    Assume it ends with END:VEVENT
    Look for start date DTSTART
    end date DTEND
    exlusions EXDATE
    when we stop repeating RRULE (also contains frequency)
    the client name SUMMARY
    We ignore everything else
    We keep getting the information
    Add to a dictionary
    At the end return value
    """
    event = {}
    event["exclusions"] = []
    while True:
        line:str = calendar_file.readline()
        if line[:7] == "DTSTART":
            event["start"] = read_date_time(line)
        elif line[:5] == "DTEND":
            event["end"] = read_date_time(line)
        elif line[:6] == "EXDATE":
            event["exclusions"].append(read_date_time(line))
        elif line[:5] == "RRULE":
            # only if the event has finished will it have a stop date
            if find_number_start(line):
                event["stop repeat"] =read_date_time(line)
            # otherwise get it to repeat until current day
            else:
                event["stop repeat"] = datetime.datetime.now
            event["freq"] = get_freq(line)
        elif line[:7] == "SUMMARY":
            event["name"] = line[8:]
        elif line[:10] == "END:VEVENT":
            return event
    

def read_calendar():
    """
        read the whole calendar file
        assume that every event begins with BEGIN:VEVENT
        and the callendar ends with END:VCALENDAR
    """
    global events
    while True:
        line = calendar_file.readline()
        if line == "END:VCALENDAR\n":
            break
        if line == "BEGIN:VEVENT\n":
            events.append(get_one_data())
    calendar_file.close()
    return 

def repeated_events():
    """If repeated, add items to event list for each time its repeated

    We can check if repeated as it should have a stop_repeat item 
    """
    global events
    for event in events:
        if "stop_repeat" in event:
            stop_date_time = event["stop repeat"]

            new_event = []
            if event["freq"] == "weekly":
                pass
                
         





def write_calendar():
    global events
    for event in events:
        for item in event:
            output_file.write(item + ",")
        output_file.write("\n")

def print_events():
    for i in events:
        print(i)
        print("\n")

