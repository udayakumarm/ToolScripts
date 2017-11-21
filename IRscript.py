#Version 0.01

#	Version 0.02 (Udaya kumar, 16 Nov 2017)
#	1. Changed Username & password as input feilds from user.
#	2. Now the script can be run from any client PC.
#   3. Masking the urlib3 comments
#   Version 0.03 (Udaya kumar, 21 Nov 2017) 
#	1. Masking urllib3 in the logs - typo corrected.
#	2. Changed the email credentials as input to the script.
#	3. Changed logging strings
#   4. Changed the Project.xlsx path to current project path

import jira.client
from jira.client import JIRA
import collections
import time
from dateutil import parser
import datetime, calendar
from time import gmtime, strftime
import mailbox
import xlrd
from xlrd import open_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import getpass

# for Logging
import logging
LOG_FILENAME = 'JIRA.log'
logging.basicConfig(filename=LOG_FILENAME,level=logging.DEBUG)
logging.getLogger("urllib3").setLevel(logging.WARNING)

# Counter for number of defect whose IR is about to Miss as per Current date
counter=0
DefectCount=0
RicohCount=0
CanonCount=0
SharpCount=0
KDCCount=0
RisoCount=0
KMBTCount=0
XeroxCount=0
OKICount=0

# Username & Password
username = input ('Login for JIRA:')
password = getpass.getpass('Password for JIRA:')

# Email Username & Password
email_user = input ('Login for email:')
email_pwd = getpass.getpass('Password for email:')

# JIRA connection Logic
try:
    jira_options = {'server': 'https://jira.efi.com'}
    jira = JIRA(options=jira_options, basic_auth=(username, password))

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
book = open_workbook("Project.xlsx")

# Mail sending logic.
def sendMail(Partner):
    global counter
    FITID= str("FIT10%s"%issue.id)
    Sub = " : ".join([FITID, Desc])
    Due=str("IR Due date is %s"%Partner["IRDueDate"])
    messageBody = "Hi Team,\n\nPlease reproduce the issue and provide IR\n" +"\n" +Due + "\n\nThanks\nAvinash\n\nNote:-Auto Generated mail Please Do Not Respond"
    msg = MIMEText(messageBody)
    msg['Subject'] = "(Important-IR Missing) " + Sub
    msg["From"] = email_user
    recipientsTo = [Addr]
    msg["To"] = ", ".join(recipientsTo)
    mail = smtplib.SMTP('smtp.office365.com', 587)
    mail.ehlo()
    mail.starttls()
    mail.login(email_user, email_pwd)
    mail.sendmail(email_user, recipientsTo, msg.as_string())
    counter += 1
    mail.close()
    return counter

# Here we are checking OEM/Priority/Age and calling function to send mail based IR SLA
def parseIssue(JiraDetails):
    if(storeOEMName=='Ricoh' or 'KMBTM'):
        if( setPriority[0] ==Priority  and Age>2):
            print(issue.fields.summary)
            DueDate= issueCreatedDate + datetime.timedelta(days=4)
            print(DueDate)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr, 'IRDueDate': DueDate}
            sendMail(Partner)
        elif(setPriority[1] == Priority and Age>4):
            DueDate = issueCreatedDate + datetime.timedelta(days=9)
            print(DueDate)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr, 'IRDueDate': DueDate}
            sendMail(Partner)
        elif(setPriority[2] == Priority and Age>6):
            DueDate = issueCreatedDate + datetime.timedelta(days=12)
            print(DueDate)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr, 'IRDueDate': DueDate}
            sendMail(Partner)
    elif(storeOEMName == 'Canon' or 'Sharp' or 'KDC' or 'Riso' or 'KMBT' or 'OKI' or 'Xerox'):
        if (setPriority[0] == Priority and Age > 5):
            DueDate = issueCreatedDate + datetime.timedelta(days=12)
            print(DueDate)
            print(issue.fields.summary)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr, 'IRDueDate': DueDate}
            sendMail(Partner)
        elif (setPriority[1] == Priority and Age > 10):
            DueDate = issueCreatedDate + datetime.timedelta(days=19)
            print(DueDate)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr, 'IRDueDate': DueDate}
            sendMail(Partner)
        elif (setPriority[2] == Priority and Age > 15):
            DueDate = issueCreatedDate + datetime.timedelta(days=25)
            print(DueDate)
            Partner = {'Defect_ID': ID, 'Summary': Desc, 'MailID': Addr, 'IRDueDate': DueDate}
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
    DefectCount+=1
    #print(issue.fields.summary)
    for sheet in book.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == issue.fields.project.name:
                    colidx = colidx - 1
                    storeOEMName = sheet.cell(rowidx, colidx).value
                    #print(storeOEMName, cell.value)
                    ID = str(issue.key)
                    Desc = str(issue.fields.summary)
                    Priority=str(issue.fields.priority)
                    print(issue.fields.summary)
                    if (OEM[0] == storeOEMName):
                        RicohCount+=1
                        #Addr = email_user #"avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr="Ricoh_Sustaining@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr,'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[1] == storeOEMName):
                        CanonCount+=1
                        #Addr = email_user #"avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "CANON_SUSTAINING@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[2] == storeOEMName):
                        SharpCount+=1
                        #Addr = email_user #"avinash.kumar1@efi.com"
                        #Addr = "avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "Sharp_Sustaining@efi.com"
                        JiraDetails = {'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner': storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate}
                        parseIssue(JiraDetails)
                    elif (OEM[3] == storeOEMName):
                        KDCCount+=1
                        #Addr = email_user #"avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "IDC_KDC_Sharp_Sust_Report@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[4] == storeOEMName):
                        RisoCount+=1
                        #Addr = email_user # "avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "RISO_Sustaining@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[5] == storeOEMName):
                        KMBTCount+=1
                        #Addr = email_user # "avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "KM_Sustaining@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[6] == storeOEMName):
                        KMBTCount+=1
                        #Addr = email_user #"avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "KM_Sustaining@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[7] == storeOEMName):
                        XeroxCount+=1
                        #Addr = email_user #"avinash.kumar1@efi.com"
						#Addr = email_user
                        Addr = "Xerox_Sustaining_Engg@efi.com"
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
                    elif (OEM[8] == storeOEMName):
                        OKICount+=1
                        Addr = email_user #"avinash.kumar1@efi.com"
						#Addr = email_user
                        JiraDetails ={'Defect_ID': ID, 'Summary': Desc, 'urgency': Priority, 'Partner':storeOEMName,'Duration': Age, 'MailID': Addr, 'createdDate': issueCreatedDate }
                        parseIssue(JiraDetails)
    filterStore.append(i.key)
logging.debug("****************************************************************************************************************************************")
logging.debug(time.strftime("%Y-%m-%d  %H-%M-%S\n"))
logging.debug("Number of defects from all partner channel for which IR is not entered as of today is %s" %DefectCount)
logging.debug("Ricoh IR Pending= 	%s" %RicohCount)
logging.debug("Canon IR Pending= 	%s" %CanonCount)
logging.debug("Sharp IR Pending= 	%s" %SharpCount)
logging.debug("KDC IR Pending= 		%s" %KDCCount)
logging.debug("RISO IR Pending= 	%s" %RisoCount)
logging.debug("KMBT IR Pending= 	%s" %KMBTCount)
logging.debug("Xerox IR Pending= 	%s" %XeroxCount)
logging.debug("OKI IR Pending= 		%s" %OKICount)
logging.debug("Number of defects from all partner channels which are about to miss the IR SLA as of today %s" %counter)
logging.debug("****************************************************************************************************************************************\n")
print(filterStore)


