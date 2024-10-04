# %%
import json
# load setting file(json)
with open("calendar.json", "r", encoding="utf-8") as f:
    dic = json.load(f)

# "id":"xxxxx@group.calendar.google.com"
calendar_id = dic["id"]

# "tools": {"zoom": "zoom.us","webex": "webex.com", "teams": "teams.microsoft.com"}
# currently not in use
tools = dic["tools"]

# omit
# currently not in use
if "omit" in dic:
    omit=dic["omit"]
else:
    omit=[]

if "weeks" in dic:
    weeks_ahead=dic["weeks"]
else:
    weeks_ahead=4

#for k in tools:
#    print(k,tools[k])
#print("omit",omit)
print("weeks ahead",weeks_ahead)

# %%
# Collect Outlook Appointment Items in term
import win32com.client
import datetime
import csv
from tqdm import tqdm

# load dayoff(holiday) list
dayoff_list = []
with open('syukujitsu.csv', newline='') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
        str_date = row[0]
        if not str_date.startswith('#'):
            dt = datetime.datetime.strptime(str_date,'%Y/%m/%d')
            dayoff_list.append(dt.date())

# Search dayoff list
def isDayOn(date):
    for dayoff in dayoff_list:
        if date == dayoff:
            return False
    return True

# is the appointment onlie meeting?
def isOnline(item):
    for k, v in tools.items():
        if v in item.Body:
            for sub in omit:
                if sub in item.Subject:
                    return False, None
            return True, k
    return False, None

# Outlook appointment search term:From.
st = datetime.date.today() - datetime.timedelta(days=1)
# Outlook appointment search term:To.
ed = datetime.date.today() + datetime.timedelta(weeks=weeks_ahead)


# objects for Outlook
app = win32com.client.Dispatch("Outlook.Application")
root = app.Session.DefaultStore.GetRootFolder()
ns = app.GetNamespace("MAPI")
cal = ns.GetDefaultFolder(9)

# filter string for specific term
filterStr = '[Start]>="{0}" and [Start]<"{1}"'.format(
    st.strftime("%Y/%m/%d"), ed.strftime("%Y/%m/%d")
)
print("filter", filterStr)

# Sort appointments and divide reccurcible ones.
appointments = cal.Items
appointments.sort("[Start]")
appointments.IncludeRecurrences = True
restricted = appointments.Restrict(filterStr)
#print("from", st, "to", ed)
#counter = 0
events = []
print('restricted length:',restricted.Count)
# Access to Outlook appointment
for item in tqdm(restricted):
    # Read each item's body and find tools string in it. If they has, that's what you need
    if isDayOn(item.Start.date()) and not "private" in item.Subject:
        #ret, prefix = isOnline(item)
        ret = True
        #countStr = "{:02}".format(counter)
        if ret:
            summary = item.subject
            desc = "EntryId={}".format(item.EntryId)
            # Build a JSON data for a google calendar event.
            event = {
                "summary": summary,
                "description": desc,
                "start": {
                    "dateTime": item.Start.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "Japan",
                },
                "end": {
                    "dateTime": item.End.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": "Japan",
                },
            }
            # colorId by busy status
            if item.busyStatus == 3:
                event["colorId"]='2'
            elif item.busyStatus == 1:
                event["colorId"]='8'
            events.append(event)

            #counter += 1
#print("Count",counter)

# %%
# file save
with open('events.json', 'wt') as f:
    json.dump(events, f)
print("save done",len(events))
