from __future__ import print_function
import requests
from bs4 import BeautifulSoup
import time
import os
import datetime
from datetime import date,  timedelta
from pathlib import Path
import smtplib
from email.message import EmailMessage
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import PySimpleGUI as sg
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
import pickle
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from rich.traceback import install
install(show_locals=True)
# basic set up


lm_and_lpt = {"NLB" : ["hannah.pham@algo.edu.vn", "annie.ha@algo.edu.vn"],
              "NVH" : ["finn.tran@algo.edu.vn", "tuananh.phan@algo.edu.vn"],
              "VCP" : ["mharvir.john@algo.edu.vn", "jolin.tran@algo.edu.vn"],
              "PXL" : ["mharvir.john@algo.edu.vn", "ashley.bui@algo.edu.vn"]}



# lm_and_lpt = {"NLB" : ["finn.tran@algo.edu.vn", ""],
#               "NVH" : ["finn.tran@algo.edu.vn", ""],
#               "VCP" : ["finn.tran@algo.edu.vn", ""],
#               "PXL" : ["finn.tran@algo.edu.vn", ""]}

GSCOPES = ['https://www.googleapis.com/auth/gmail.send']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
today = date.today()
tomorrow = (date.today() + timedelta(1)).strftime("%d/%m/%Y")
# get Gmail creds
def get_gmail_credentials():
    # Get the credentials for the Gmail API.
    gmail_creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            gmail_creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not gmail_creds or not gmail_creds.valid:
        if gmail_creds and gmail_creds.expired and gmail_creds.refresh_token:
            gmail_creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', GSCOPES)
            gmail_creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(gmail_creds, token)
    return gmail_creds

# Function to send an gmail to LM and LPT
def send_email(service, to, subject, body, cc = None, userEmail = None):
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject
    message.attach(MIMEText(body))

    if cc:
        message['cc'] = cc
    if userEmail:
        message['bcc'] = userEmail


    create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}
    send_message = (service.users().messages().send(userId="me", body=create_message).execute())
    print(F'Message Id: {send_message["id"]}')
    return send_message




# To make a folder and create path
def make_folder():    
    desktop_path = Path(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
    # Check if OneDrive is enabled
    onedrive_path = Path(os.path.expanduser("~\OneDrive"))
    if onedrive_path.exists():
        # OneDrive is enabled, try to find the Desktop folder in OneDrive folder
        desktop_name_en = "Desktop"
        desktop_name_vi = "Máy tính"
        onedrive_desktop_path_en = os.path.join(onedrive_path, desktop_name_en)
        onedrive_desktop_path_vi = os.path.join(onedrive_path, desktop_name_vi)
        if os.path.exists(onedrive_desktop_path_en):
            # Found Desktop folder in OneDrive folder with English name
            desktop_path = Path(onedrive_desktop_path_en)
        elif os.path.exists(onedrive_desktop_path_vi):
            # Found Desktop folder in OneDrive folder with Vietnamese name
            desktop_path = Path(onedrive_desktop_path_vi)
    try:
        new_folder_path = os.path.join(desktop_path, 'CPR_Tracking')
        os.mkdir(new_folder_path)
        
    except:
        pass
    user = os.path.join(new_folder_path, "userinfo.txt")
    if os.path.exists(user):
        pass
    else:      
        user = open(user, 'w', encoding='utf-8')
    return new_folder_path



def get_web_creds():
    if os.path.exists('token.json'):
            web_creds = Credentials.from_authorized_user_file('token.json', SCOPES) 
    if not web_creds or not web_creds.valid:
        if web_creds and web_creds.expired and web_creds.refresh_token:
            web_creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            web_creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(web_creds.to_json())
    return web_creds




def get_html(date, base):
    s = requests.Session()
    DOMAINS = 'https://lms.logika.asia/group/default/schedule'
    para = f'?GroupLessonSearch%5Bstart_time%5D={date}+-+{date}&GroupLessonSearch%5Bnumber%5D=&GroupLessonSearch%5Blesson.title%5D=&GroupLessonSearch%5Bgroup_id%5D=&GroupLessonSearch%5Bgroup.title%5D={base}&GroupLessonSearch%5Bweekday%5D=&GroupLessonSearch%5Bteacher.name%5D=&GroupLessonSearch%5Bgroup.teacher.name%5D=&GroupLessonSearch%5Bis_online%5D=&export=true&name=default&exportType=html'
    header =  {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "en-US,en;q=0.9,vi-VN;q=0.8,vi;q=0.7",
        "cache-control": "max-age=0",
        "sec-ch-ua": "\"Google Chrome\";v=\"113\", \"Chromium\";v=\"113\", \"Not-A.Brand\";v=\"24\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\"",
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "cookie": "_ym_uid=1664966280852805697; intercom-device-id-ufjpx6k3=fd3080c1-5a63-4af6-bd1f-15268d549a7e; userId=30454; _grid_page_size_schedule=35d0980fa38e2255112d0c62698773cab8aa12a81c6735caf172064b5eb6ea47a%3A2%3A%7Bi%3A0%3Bs%3A24%3A%22_grid_page_size_schedule%22%3Bi%3A1%3Bs%3A3%3A%22200%22%3B%7D; _grid_page_size=d3ebbdf9ec9235bfc4ba59572a6d3dd403a563222be148e9c705e91612194e49a%3A2%3A%7Bi%3A0%3Bs%3A15%3A%22_grid_page_size%22%3Bi%3A1%3Bs%3A3%3A%22200%22%3B%7D; _ym_d=1680775536; createdTimestamp=1683002626; accessToken=0237403b60b60d9d3dbdf5b12982ef1f04218ba627b7b4037c12f2eed6540d3d; sidebar-state=collapsed; SERVERID=b600; studentId=1700942; intercom-session-ufjpx6k3=elk4aTZOa3RiOENvUlMxWEFhWXlCaWZvb2t5eis2RGpzeGhSRmlKV3Y4NjlaeU9XNDcwYlMxa0c1OXR3NUIyZi0tQnRlU3EvbFl3ZEdrdFgrOXE4ZWpQdz09--b3809f93ab62ae12412a57106bfbf4e71e83dd02; studentAccessToken=bf1bc6733067610400d89aeefb5b262dc2e513c4dc1420902aab305dd910dfc0; studentCreatedTimestamp=1683633732; _gid=GA1.2.1195310546.1683878976; _ym_isad=1; _backendMainSessionId=9421651814e1835f2f1cab69954c0f5e; SERVERID=b440; _ym_visorc=w; _ga_3QSGZBLTE3=GS1.1.1683884650.35.1.1683885185.0.0.0; _ga=GA1.2.1620969847.1664966279",
        "Referer": f"{DOMAINS}?GroupLessonSearch%5Bstart_time%5D={date}%20-%20{date}&GroupLessonSearch%5Bnumber%5D=&GroupLessonSearch%5Blesson.title%5D=&GroupLessonSearch%5Bgroup_id%5D=&GroupLessonSearch%5Bgroup.title%5D={base}&GroupLessonSearch%5Bweekday%5D=&GroupLessonSearch%5Bteacher.name%5D=&GroupLessonSearch%5Bgroup.teacher.name%5D=&GroupLessonSearch%5Bis_online%5D=",
        "Referrer-Policy": "strict-origin-when-cross-origin"
    }

    end_point = DOMAINS + para
    res = s.get(end_point, headers = header)
    output = res.content.decode('utf-8')
    soup = BeautifulSoup(output, 'html.parser', from_encoding='utf-8')
    soup.find('tbody')
    # get data 
    all_tr = soup.find_all('tr')
    return all_tr

def get_feedback(feedback_need):
    feedback_sheet = {1 : "M36:O45", 4 :"S36:V45" ,8 : "P69:S78", 12 : "P102:S111", 16: "P144:S144", 20 : "P168:S177", 24 : "P201:S210", 28 : "P234:S243", 32 : "P267:S276", 36 : "P300:S309", 38 : "P333:S342" }
    class_report_filled = []
    class_report_not_filled = []
    service = build('sheets', 'v4', credentials=get_web_creds())
    sheet = service.spreadsheets()
    for group_name,value in feedback_need.items():
        id = value[0]
        lesson = value[1]
        link = value[2]
        feed = feedback_sheet[lesson]
        # student = student_sheet[lesson]
        try:
            result2 = sheet.values().get(spreadsheetId=id, range=f'Report!{feed}').execute()
            values2 = result2.get('values', [])
            # print(values2)
            if len(values2) == 0:
                class_report_not_filled.append(f"{group_name} - Have Not Fill 4 Lessons Feedback - For Lesson {lesson}. Date Check: {today}\nHere is the Link to the CPR: {link} \n")
            else: 
                class_report_filled.append(f"{group_name} - Have Fill 4 Lessons Feedback - For Lesson {lesson} \n\n")
        except HttpError:
            time.sleep(60)

    return class_report_filled, class_report_not_filled


def get_CPR(name, base,date):
    new_folder_path = make_folder()
    cpr_status_finished = []
    cpr_status_not_finished = []
    groups = [] 
    cpr_status_finished = []
    cpr_status_not_finished = [] 
    feedback_need = {}
    class_report_not_filled = []
    class_report_filled = []
    email = ''
    email_head = 'Dear Academic Management team,\nThis an AUTOMATIC email is to inform you about the CPRs that have not been completed:\n\n'
    email_end = f"Please can you confirm whether the CPRs have been filled or is there anything I can do to help speed the process along?\nBest {name}"
    try:
        date = datetime.datetime.strptime(date, "%Y-%m-%d").date()
        diff = int((today - date).days)
    except:        
        diff = int((today - date).days)
    date = date.strftime("%Y-%m-%d")
    all_tr = get_html(date,base)
    # all_tr.remove(all_tr[0])
    for tr in all_tr:
        all_td = tr.find_all('td')
        time_of_next_les = all_td[0]
        next_les = all_td[1]
        group_title = all_td[4].find('a')
        all_li = all_td[4].find('p')
        id = str(all_li)[str(all_li).find("https"):]
        link = str(all_li)[str(all_li).find("https"):]
        class_occurrences = all_td[5].find('span')
        teacher_name = all_td[8].string
        if teacher_name is None:
            print(group_title)
        teacher_name = str(teacher_name).split(' ')
        teacher_name = teacher_name[len(teacher_name)-1]
        
        start = id.find("d/")
        end = id.find("/edit")    
        id = id[start+2:end]
        if group_title is None:
            continue
        
        try:
            teacher_name = teacher_name.replace("(", "").replace(")", "")
        except:
            continue


        all_li = all_li.string
        group_details = ({
            'group' : group_title.string,
            'nextLesson': int(next_les.string),
            'Time of Next Lesson': time_of_next_les.string.replace('\xa0', ' '),
            "CPR Link": all_li[all_li.find("https"):],
            "Pecentage": str(class_occurrences.string),
            "teacher_name": str(teacher_name)
            })
        groups.append(group_details)
    # get CPR 
        if int(next_les.string) in [1,4,8,12,16,20,24,28]:
            feedback_need[str(group_title.string)] = [id,int(next_les.string), link]
    for group in groups:
        lessons = []
        id = group["CPR Link"]
        start = id.find("d/")
        end = id.find("/edit")    
        id = id[start+2:end]
        try:
            service = build('sheets', 'v4', credentials=get_web_creds())
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=id, range='Report!AC2:AC40').execute()
            values = result.get('values', [])
            for j in values:
                lessons.append(j[0]) 
            if len(lessons) == 0:
                result = sheet.values().get(spreadsheetId=id, range='Report!AE2:AE40').execute()
                values = result.get('values', [])

                for j in values:
                    lessons.append(j[0]) 
    # store result
            class_report_filled, class_report_not_filled = get_feedback(feedback_need)
        
            if diff == 1: 
                if (lessons[group["nextLesson"]] != '#DIV/0!') or (lessons[group["nextLesson"]] != '#DIV/0!' and group["Pecentage"] == "0%"):
                    cpr_status_finished.append(f"{group['group']} - Lesson {group['nextLesson']} - FILLED - AVG: {lessons[group['nextLesson']]}\n")


                elif lessons[group["nextLesson"]] == '#DIV/0!' and group["Pecentage"] == "0%":
                    cpr_status_finished.append(f"{group['group']} - N/A\n")
                else:
                    cpr_status_not_finished.append(f"Hi teacher {group['teacher_name']}\nMy name is {name} - Programming Tutor at {base} campus.\nI've received a notification from StepHan the CPR checking bot that there has been a delay of CPR scoring for {group['group']} - Lesson {group['nextLesson']}\nIs there anything I can support you to have this fill out by tomorrow {tomorrow} according to Algorithmics policy?\n")
                

                
            else:
                if (lessons[group["nextLesson"]] != '#DIV/0!' and group["Pecentage"] != "0%") or (lessons[group["nextLesson"]] != '#DIV/0!' and group["Pecentage"] == "0%"):
                    cpr_status_finished.append(f"{group['group']} - Lesson {group['nextLesson']} - FILLED - AVG: {lessons[group['nextLesson']]}\n")


                elif lessons[group["nextLesson"]] == '#DIV/0!' and group["Pecentage"] == "0%":
                    cpr_status_finished.append(f"{group['group']} - N/A\n")
                else:
                    cpr_status_not_finished.append(f"{group['group']} - Lesson {group['nextLesson']} - {group['teacher_name']} - Checking Date: {today}.\nLink: {group['CPR Link']}\n\n")
            
        
        except HttpError as err:
            cpr_status_finished.append(f"This {group['group']} don't have CPR!\n")

            
    # Create a new text file in the folder
    new_file_path = os.path.join(new_folder_path, f'{date}_{base}_CPR.txt')
    with open(new_file_path, 'w', encoding='utf-8') as file:
        for item in cpr_status_finished:
            file.write("%s\n" % item)
        for item in class_report_filled:
            file.write("%s\n" % item)
        file.write("<---------------------------------------------------------------------->\n")   
        if diff == 1:
            for item in cpr_status_not_finished:
                file.write("%s\n" % item)
        elif diff != 1 and (len(cpr_status_not_finished) != 0 or len(class_report_not_filled) != 0):
            file.write(email_head)
            email += email_head
            for item in cpr_status_not_finished:
                email += item
                file.write(item)
            for item in class_report_not_filled:
                email += item
                file.write(item)
            file.write(email_end)
            email += email_end
    
    return email


# this will help create file and make a user_info file to save for next use
def input_info():
    new_folder_path = make_folder()
    user_info_path = os.path.join(new_folder_path, "userinfo.txt")
    checking = sg.popup_yes_no("Do you want to do Today CPR Task(Y/N)?")
    with open(user_info_path, 'r', encoding='utf-8') as file:
        
        # print(os.path.getsize(user_info_path))

        if os.path.getsize(user_info_path) > 0:
            if sg.popup_yes_no("Do you want to use last time name and base?") == "Yes":
                lines = file.readlines()
                name, base, userEmail = lines[0].strip(), lines[1].strip(), lines[2].strip()
            else:
                with open(user_info_path, 'w', encoding='utf-8') as file:
                    base = sg.popup_get_text("Enter the Campus to check:")
                    name = sg.popup_get_text("Enter your name:")
                    userEmail = sg.popup_get_text("Enter your email:")
                    file.write(f"{name}\n")
                    file.write(f"{base}\n")
                    file.write(f"{userEmail}")
        else:
            with open(user_info_path, 'w', encoding='utf-8') as file:
                base = sg.popup_get_text("Enter the Campus to check:")
                name = sg.popup_get_text("Enter your name:")
                userEmail = sg.popup_get_text("Enter your email:")
                file.write(f"{name}\n")
                file.write(f"{base}\n")
                file.write(f"{userEmail}")
    return checking, name, base, userEmail





def main():
    email = ''
    layout = [[sg.Text('Loading...', font=('Helvetica', 20))]]
    window = sg.Window('Loading', layout, no_titlebar=True, finalize=True)
    checking,name,base,userEmail = input_info()
    if checking == "Yes":
        for i in range(1,3):
            date = today - timedelta(days=i)
            email = get_CPR(name,base,date)
    else:
        date = sg.popup_get_text("Enter the date to check (yyyy-mm-dd):")    
        email = get_CPR(name,base,date)
    window.close()
    sg.popup_yes_no(f"CPR Checking is Done please check your desktop for the info!")
    print(email)
    # Create a Gmail API service client
    service = build('gmail', 'v1', credentials=get_gmail_credentials())
    # SMTP server details
    if email != '' and sg.popup_yes_no("Do you want to send an Email to Report to your LM?") == "Yes":    
        # Email details
        
        to = lm_and_lpt[base][0]
        cc = lm_and_lpt[base][1]
        # bcc = lm_and_lpt[base][2]
        subject = 'CPR CHECKING REPORT'
        body = email
        # Send the email
        send_email(service, to, subject, body, cc, userEmail)

if __name__ == '__main__':
    main()