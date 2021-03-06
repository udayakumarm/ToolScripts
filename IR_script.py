# Version 0.1 - 
# Feature - 
import jira.client
from jira.client import JIRA
import collections
import time
from dateutil import parser
import datetime, calendar
from time import gmtime, strftime
import mailbox
import xlrd, xlwt
from xlrd import open_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

# for Logging
import logging
LOG_FILENAME = 'JIRA.log'
logging.basicConfig(filename=LOG_FILENAME,level=logging.DEBUG)
logging.debug(time.strftime("%Y-%m-%d  %H-%M-%S"))

# Counter for number of defect whose IR is about to Miss as per Current datw
counter=0

# JIRA connection Logic
try:
    logging.debug("Connecting to JIRA")
    jira_options = {'server': 'https://jira.efi.com'}
    jira = JIRA(options=jira_options, basic_auth=('avinasku', ''))
    logging.debug("Connection established with JIRA")

except Exception as e:
    logging.debug("Failed to connect to JIRA: %s" % e)

# Created list to store issue in Filter
filterStore=list()

# To check the Priority of the defect
setPriority = ["P1","P2","P3"]

# Store OEM name from for Particular defect
storeOEMName=""

# OEM supported
OEM=["Ricoh","Canon","Sharp","KDC","Riso","KMBT","KMBTM","Xerox","OKI"]

#Database to Map OEM
book = open_workbook("C:\Python37\Project.xlsx")

# Mail sending logic.
def sendMail(Partner):
    global counter
    Sub = " : ".join([ID, Desc])
    messageBody = "Hi Team,\n\nPlease reproduce the issue and update IR\n\nThanks\nAvinash"
    msg = MIMEText(messageBody)
    msg['Subject'] = "(Important-IR Missing) " + Sub
    msg["From"] = 'avinash.kumar1@efi.com'
    recipientsTo = [Addr]
    msg["To"] = ", ".join(recipientsTo)
    mail = smtplib.SMTP('smtp.office365.com', 587)
    mail.ehlo()
    mail.starttls()
    mail.login('avinash.kumar1@efi.com', '1731sH!@#')
    mail.sendmail('avinash.kumar1@efi.com', recipientsTo, msg.as_string())
    counter += 1
    mail.close()
    return counter

# Here we are checking OEM/Priority/Age and calling function to send mail based IR SLA
def parseIssue(JiraDetails):
    if(storeOEMName=='Ricoh' or 'KMBTM'):
        if( setPriority[0] ==Priority  and Age>2):
            print(issue.fields.summary)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr}
            sendMail(Partner)
        elif(setPriority[1] == Priority and Age>4):
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr}
            sendMail(Partner)
        elif(setPriority[2] == Priority and Age>6):
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr}
            sendMail(Partner)
    elif(storeOEMName == 'Canon' or 'Sharp' or 'KDC' or 'Riso' or 'KMBT' or 'OKI' or 'Xerox'):
        if (setPriority[0] == Priority and Age > 5):
            print(issue.fields.summary)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr}
            sendMail(Partner)
        elif (setPriority[1] == Priority and Age > 10):
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr}
            sendMail(Partner)
        elif (setPriority[2] == Priority and Age > 15):
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr}
            sendMail(Partner)
    return

#Parsing the defect from JIRA Filter

for i in jira.search_issues('filter=64702',startAt=0, maxResults=1000):
    issue = jira.issue(i.key)
    issueCreatedDate= issue.fields.created
    issueCreatedDate=((issueCreatedDate[0:-18]))
    issueCreatedDate=datetime.datetime.strptime(issueCreatedDate, '%Y-%m-%d').date()
    TodaysDate=datetime.datetime.strptime((datetime.datetime.today().strftime('%Y-%m-%d')), '%Y-%m-%d').date()
    delta=TodaysDate-issueCreatedDate
    Age=delta.days
    logging.debug("FIT10%s"%issue.id)
    logging.debug("Age of the Defect is %s" %Age)
    #print(issue.fields.summary)
    for sheet in book.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == issue.fields.project.name:
                    colidx = colidx - 1
                    storeOEMName = sheet.cell(rowidx, colidx).value
                    logging.debug("OEM=%s"%storeOEMName)
                    #print(storeOEMName, cell.value)
                    ID = str(issue.key)
                    Desc = str(issue.fields.summary)
                    Priority=str(issue.fields.priority)
                    print(issue.fields.summary)
                    if (OEM[0] == storeOEMName):
                        Addr="avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[1] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[2] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails = {'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner': storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[3] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[4] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[5] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[6] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[7] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)
                    elif (OEM[8] == storeOEMName):
                        Addr = "avinash.kumar1@efi.com, avinash.kumar1@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr }
                        parseIssue(JiraDetails)

    filterStore.append(i.key)

logging.debug("Number of defect in all channel which are about to Miss the IR SLA as per current date %s" %counter)
print(filterStore)
