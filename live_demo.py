from __future__ import print_function
import datetime
import dateutil.parser
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import socket
from geolite2 import geolite2

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
# Date from and date too
date_min = datetime.datetime(2018, 1, 1, 0,0,0,0)
date_max = datetime.datetime(2019, 1, 1)


def main():
    """Extract meetings and contact from google calendar from the specified date range
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))



    #now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    now = date_min.isoformat() + 'Z' # 'Z' indicates UTC time
    then = date_max.isoformat() + 'Z'


    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        timeMax=then,
                                        maxResults=10000, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    # events contains the event data
    total_time = datetime.timedelta(0)
    attendees = {}



    outputfile = open("meetings.csv", "w")
    email_addresses_file = open("emails.tsv", "w")
    companies_file = open("companies.tsv", "w")
 
    if not events:
        print('No upcoming events found.')

    # for each event dump the details into the output file
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))
        parsedStart = dateutil.parser.parse(start)
        parsedEnd = dateutil.parser.parse(end)
        duration = parsedEnd - parsedStart

        total_time = total_time + duration
        
        outputfile.write(str(duration).encode('utf8'))
        outputfile.write("\t")

        if "summary" in event:
            meeting_summary = event['summary'].encode('utf8')
        else :
            meeting_summary = "none"
        outputfile.write(meeting_summary)

        outputfile.write("\t")

        if "htmlLink" in event:
            meeting_htmlLink = str(event['htmlLink']).encode('utf8')
        else :
            meeting_htmlLink = "none"
        outputfile.write(meeting_htmlLink)
        outputfile.write("\t")

        if "conferenceData" in event:
            meeting_conferenceData = str(event['conferenceData']).encode('utf8')
        else :
            meeting_conferenceData = "none"
        outputfile.write(meeting_conferenceData)
        outputfile.write("\t")
        
        email_string = ""
        if "attendees" in event:
            for person in event['attendees']:
                if "email" in person:
                    email_string = str(person['email']).encode('utf8')
                    #print(email_string)
                    if attendees.get(email_string) is None :
                        attendees[email_string] = duration
                    else :
                        attendees[email_string] = attendees[email_string] + duration

        outputfile.write(email_string)
        outputfile.write("\t")

        if "organizer" in event:
            meeting_organizer = str(event['organizer']).encode('utf8')
        else :
            meeting_organizer = "none"

        outputfile.write(meeting_organizer)
        outputfile.write("\t")

        if "location" in event:
            meeting_location = event['location'].encode('utf8')
        else :
            meeting_location = "none"

        outputfile.write(meeting_location)
        #outputfile.write(str(event).encode('utf8'))
        
        outputfile.write("\n")
    
    



    companies = {}
    
    # for each of the attendees, put their email address domain name into the companies file
    for chap in sorted(attendees, key=attendees.get, reverse=True):
        email_addresses_file.write((str(chap) + "\t" + str(attendees[chap])) + "\n")
        domain = chap.split("@")[1]
        if companies.get(domain) is None :
            companies[domain] = 1
        else :
            companies[domain] = companies[domain] + 1
    
    for company in sorted(companies, key=companies.get, reverse=True):
        companies_file.write("\t" + str(company) + "\t" + str(companies[company]) +  "\n")
        try:
            getip(str(company))
        except socket.error as msg:
            print("{0} [could not resolve]".format(str(company).strip())) 
            if len(str(company)) > 2:
                subdomain = str(company).split('.', 1)[1]
                try:
                    getip(subdomain)
                except:
                    continue
    
    email_addresses_file.close()
    outputfile.close() 
    companies_file.close()

    print("\tNumber of meetings\t\t\t" + str(len(events)))
    print("\tTotal time:\t\t\t\t"+str(total_time))
    print("\tNumber of unique attendees:\t\t" + str(len(attendees)))
    print("\tNumber of unique companies (domains):\t" + str(len(companies)) + "\n")
 

def getip(domain_str):
    ip = socket.gethostbyname(domain_str.strip())
    reader = geolite2.reader()      
    output = reader.get(ip)
    if "country" in output :
        result = output['country']['iso_code']
        print("{0} [{1}]: {2}".format(domain_str.strip(), ip, result))


if __name__ == '__main__':
    main()
