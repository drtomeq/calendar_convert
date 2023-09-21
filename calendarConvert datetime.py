"""imports Google calendar data and convert and export into format for spreadsheets.

to be used for invoices tax returns etc
Sadly what is missing in the Google calendar export are the colours
So it might not be able to identify red canceled sessions
or sessions that are yellow maybe
"""
import datetime
# use relativedelta as it allows to add 1 year
from dateutil.relativedelta import relativedelta

# adjust times for BST
# ZoneInfo only works if IANA time database is included in system
# tzdata needed if no database on system
import tzdata
from zoneinfo import ZoneInfo
now = datetime.datetime.now()
mytz = ZoneInfo("Europe/London")

# can't use 0 in datetime so initialise to 1
# make a global version here so can get suggestions 
init_dt = datetime.datetime(1,1,1, tzinfo=mytz)
print(init_dt)
# my_dt = datetime.datetime(1,1,1, tzinfo=datetime.timezone.utc)
# tzdt = datetime.datetime(2023,1,19, tzinfo=ZoneInfo("Europe/London"))


# the calendar file as exported from Google calendar
# change to the appropriate file name
# make sure you check the file names before running
# it will overwrite existing data so be careful
calendar_file = open("londonmathstuitionAug2023.txt", "r")
# the output file which will have the data in a 
# spreadsheet friendly format
output_file = open("calendar2023Aug.csv", "w")

# each event in the calendar, a global variable
events = []

# while there are other categories, these are the ones we are intersted in
# categories = ["name","start","end","stop repeat","exclusions"]

def find_number_start(line):
    '''Find the first character that is a number
    
    used to find where datetimes are in the source text'''
    for pos, val in enumerate(line):
        if val.isdecimal():
            start_pos = pos
            return pos 

def DSTadjust(dt: datetime):
    '''Change for daylight saving time
    
    this was a longer function, but it was uneccesary 
    as time zones can correct for seasonal changes'''

    '''
    print("before", dt)
    tzdt = dt.replace(tzinfo=ZoneInfo("Europe/London"))
    tzdt = tzdt.replace(hour=((tzdt.hour + (tzdt.dst()).seconds//3600))%24)
    tzdt = dt.replace(tzinfo=datetime.timezone.utc)
    '''
    tzdt = dt.astimezone(mytz)
    # print("after", tzdt)
    return tzdt

def read_date_time(line):
    """Get one date or date and time value
        
    Assume that the date is the first number on the line
    Assume it is the standard format of 
        * 4 digit year number
        * 2 digit month number
        * 2 digit day date number
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

    global init_dt
    my_dt = init_dt

    # we assume it is in the format as described above
    # so we know the place of year, month, day, hours, min
    # if not raise an exception
    try:
        my_dt = my_dt.replace(year=int(line[start_pos:start_pos+4]))
        my_dt = my_dt.replace(month=int(line[start_pos+4:start_pos+6]))
        my_dt = my_dt.replace(day=int(line[start_pos+6:start_pos+8]))

        # all day events dont have a time so will have a shorther length
        if len(line)>=start_pos+10:
            # skip one character in string for the letter T for time
            my_dt = my_dt.replace(hour=int(line[start_pos+9:start_pos+11]))
            my_dt = my_dt.replace(minute=int(line[start_pos+11:start_pos+13]))

    except(ValueError):
        print("************* Exception: not a datetime number! ******************")
        print(line)
        datetime = {}

    my_dt = DSTadjust(my_dt)

    return my_dt

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
    Look for start date DTSTART
    Assume it ends with END:VEVENT
    end date DTEND
    exlusions EXDATE
    when we stop repeating RRULE (also contains frequency)
    the task or client name SUMMARY
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
            # if there is a number assume it is the date
            if find_number_start(line):
                event["stop repeat"] =read_date_time(line)
            # otherwise get it to repeat until current day
            else:
                event["stop repeat"] = datetime.datetime.now()
                event["stop repeat"].replace(tzinfo=datetime.timezone.utc)
            event["freq"] = get_freq(line)
        elif line[:7] == "SUMMARY":
            event["name"] = line[8:]
        elif line[:10] == "END:VEVENT":
            return event
    

def read_calendar():
    """read the whole calendar file and add to events
    
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

def repeated_events():
    """If repeated, add items to event list for each time its repeated

    We can check if repeated as it should have a stop_repeat item 
    """

    # initially I just added to the event list
    # but it resulted in multiple copies
    # so make a separate list
    extra_events = []

    global init_dt
    my_dt = init_dt
    global events
    for event in events:
        # all repeated events should have a stop repeat
        # if it didn't, it should have been assigned now earlier
        try: 
        # we do equality to raise exception if stop repeat does not exists
            event["stop repeat"] = event["stop repeat"]
        except(KeyError):
            # if there is no stop repeat, skip it go to next event
            continue
        stop_dt = event["stop repeat"]
        new_event = event.copy()
        if event["freq"] == "weekly":
            delta = datetime.timedelta(weeks=1)
        elif event["freq"] == "yearly":
            # timedelta does not allow for year changes 
            # so use relativedelta method to add a year
            delta = relativedelta(years=1)
        elif event["freq"] == "daily":
            delta = datetime.timedelta(days=1)
        else:
            print("******* not the right frequency ********")
            print("freq = ", event["freq"])
            continue

        while True:
            new_event["start"] = new_event["start"] + delta 
            new_event["end"] = new_event["end"] + delta

            if (new_event["end"]).timestamp() > stop_dt.timestamp():
                break
            if new_event["start"] in new_event["exclusions"]:
                new_event["exclusions"].remove(new_event["start"]) 
            else:
                extra_events.append(new_event.copy())
    events.extend(extra_events)

# replace commas with slash to work in csv format
def replace_commas(in_str:str)->str:
    if "," not in in_str:
        return in_str
    out_str=""
    for character in in_str:
        if character == ",":
            in_str = in_str.replace(character, "/") 
    return in_str


def write_calendar():
    output_file.write("name,start,length\n")
    global events
    for event in events:
        if type(event) == dict: 
            if len(event["name"])>1 and type(event["name"])==str:
              event["name"] = replace_commas(event["name"])
              # go to penultimate char as last is a new line
              output_file.write(event["name"][:-1] + ",")
            else:
                print("********** wrong name type ********************")
        else:
            print("*************** event is not a dict *****************")
            print(event)
            continue
        output_file.write(str(event["start"]) + ",")
        output_file.write(str(event["end"]-event["start"]) + ",")
        output_file.write("\n")

def write_unformated():
    global events
    for event in events:
        for item in event:
            output_file.write(item + " : " + str(event[item])+",")
        output_file.write("\n")


def print_events():
    for i in events:
        print(i)
        print("\n")

read_calendar()
repeated_events()
# write_unformated()
write_calendar()

# print(events)
