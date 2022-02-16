
#Import Related Library
import imaplib, email
import os, time, json
from random import choice
from webbrowser import get
import pyautogui
import subprocess

#Defualt Value

utility = {
        'mail':['meaebill@mea.or.th','mwatax@mwa.co.th'],
        'invoice':['MEA e-Invoice','e-Invoice'],
        'receipt':['e- Receipt/e-Tax Invoice','e-RECEIPT/TAX INVOICE'],
        'keep_fdir':['ค่าไฟฟ้า - การไฟฟ้านครหลวง','ค่าน้ำประปา - การประปานครหลวง'],
        'keep_folder':['Electric','Water']
        }# 0 = Electric // 1 = Water
    
utility_cat = {'ค่าไฟฟ้า': 0,'ค่าน้ำประปา': 1}
utility_type = {'ใบแจ้งหนี้': 'invoice','ใบกำกับภาษี': 'receipt'}

mail_id = r'C:\Users\admin\Python_Files\Python_Notebook\Email_PY\credential_mail.json'

download_dir = r'E:\5.Keeping_Online_Doc\Utility_Doc' #keeping files location


def get_mail(mail_location):
    up = open(mail_location)
    updata = json.load(up)
    email_user  = updata['mail']
    email_pw = updata['password']
    up.close()
        
    return email_user, email_pw

#Action Python
user, pw = get_mail(mail_id)

mail = imaplib.IMAP4_SSL('imap.gmail.com',993)
mail.login(user,pw)
print('Login Successfully')

time.sleep(2)

ch_cat = pyautogui.confirm(text='กรุณเลือกชนิดของบิลที่ต้องการ Download', title='Choice', buttons=['ค่าไฟฟ้า', 'ค่าน้ำประปา'])
choice_cat = int(utility_cat[ch_cat]) #ประเภทของข้อมูลที่ต้องการ download
print(ch_cat)

ch_type = pyautogui.confirm(text='กรุณเลือกชนิดของบิลที่ต้องการ Download', title='Choice', buttons=['ใบแจ้งหนี้', 'ใบกำกับภาษี'])
choice_type = utility_type[ch_type] #ชนิดของข้อมูลที่ต้องการ download
print(choice_type)

mailnum = pyautogui.prompt(text='จำนวนเมล์ที่ต้องการ Download - กรุณากรอกเป็นตัวเลข', title='Numbers#' , default='1')
print(mailnum)

def filter_mail(mail, cat_choice, type_choice):

    mail.select('Inbox') #search where to search before actual search

    mail_from = utility['mail'][cat_choice] #Category of mail - Eletric or Water
    mail_head = utility[type_choice][cat_choice] 
    search_info = '(FROM "' + mail_from +'" SUBJECT "' + mail_head +'")'

    result, data = mail.search(None, search_info)

    return result, data

def get_nummails(mail, m_num, mail_data):

    ids = mail_data[0] # data is a list.
    id_list = ids.split() # ids is a space separated string

    #Identify number of mails want to download 
    #popup - number of mails to be downloaded
    num = -int(m_num) #Make number negative since want the last message filtered

    set_mea = id_list[num:] # get the latest

    mea_data = []
    for msg in set_mea:
        data = mail.fetch(msg, "(RFC822)")[1] # fetch the email body (RFC822) for the given ID
        raw_email = data[0][1]
        mea_data.append(raw_email)
        # here's the body, which is raw text of the whole email
        # including headers and alternate payloads

    mea_msg = []
    for mail in mea_data:
        raw_string = mail.decode('utf-8')
        email_msg = email.message_from_string(raw_string)
        mea_msg.append(email_msg)
        
    return mea_data, mea_msg

def get_files(message):

    count = 0
    for msg in message: 
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
        
    fileName = part.get_filename()
    
    if bool(fileName):
        filePath = os.path.join(download_dir,utility['keep_fdir'][choice_cat],fileName)
        if not os.path.isfile(filePath) :
            fp = open(filePath, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
    count+=1

    pyautogui.alert(text=f'Download Invoice--> {count} files completed', title='Notification', button='OK')

def pop_directory(cat_choice):

    from datetime import date
    today = date.today()
    yr = today.strftime("%Y") #Year - yyyy
    mth = today.strftime("%m") #month - mm
    period = str(yr) + str(mth)
    to_folder = period + "_" + utility['keep_folder'][cat_choice] #folder 

    open_dir = os.path.join(download_dir,utility['keep_fdir'][cat_choice],to_folder)
    print(open_dir)

    try:
        os.mkdir(open_dir)
    except:
        pass

    pop = [fr'explorer /select, "{open_dir}"']

    subprocess.Popen(pop)

time.sleep(5)

result, data = filter_mail(mail,choice_cat,choice_type)

mea_data, mea_msg = get_nummails(mail, mailnum, data)

get_files(mea_msg)

time.sleep(5)

#pop_directory(choice_cat)

