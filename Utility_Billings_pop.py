'''Objective : To retrived eletronic documents for accouting and keeping purposes
เอกสารปัจจุบันมีการจัดส่งเป็นรูปแบบอิเล็กโทรนิคเป็นจำนวนมาก ทั้งเพื่อความรวดเร็วและประหยัดทรัพยากรธรรมชาติ 
เพราะฉะนั้น code นี้จึงมีขึ้นเเพื่อการจัดการที่สะดวก และลดเวลาการทำงาน'''

#Import Related Library
import imaplib, email
import os, time
import pyautogui

""" 
Explaination : 
    โค้ดนี้สามารถใช้ได้กับทั้งส่วน การประปา และการไฟฟ้า โดย concept คือรับใบแจ้งหนี้ (เพื่อนำจ่าย) และรับใบกำกับ (หลังชำระเสร็จ)
    ไม่ได้จัดเข้ารูปแบบ class เพื่อให้ใช้งานได้สะดวกขึ้น 
Intention : 
    Download เอกสาร PDF ไปเก็บรักษาใน folder ที่ต้องการ 
"""

#Defualt Value
utility = {
        'mail':['meaebill@mea.or.th','mwatax@mwa.co.th'], #Source mails (sender)
        'invoice':['MEA e-Invoice','e-Invoice'], #Invoice heading include these wording
        'receipt':['e- Receipt/e-Tax Invoice','e-RECEIPT/TAX INVOICE'], #Receipt heading include these wording
        'keep_fdir':['ค่าไฟฟ้า - การไฟฟ้านครหลวง','ค่าน้ำประปา - การประปานครหลวง'], #Main folder for keeping 
        'keep_folder':['Electric','Water'] #Internal folder for keeping document
        } # 0 = Electric // 1 = Water
    
utility_cat = {'ค่าไฟฟ้า': 0,'ค่าน้ำประปา': 1}
utility_type = {'ใบแจ้งหนี้': 'invoice','ใบกำกับภาษี': 'receipt'}

mail_id = r'C:\Users\admin\desktop\login_mail.json' #mail location - keep in json form <-- (need to replace)

download_dir = r'E:\Utility_Doc' #keeping files location <-- (need to replace)


def get_mail(mail_json): #Get information for mail login
    #mail_json : as full path to json file
    import json
    up = open(mail_json)
    updata = json.load(up)
    email_user  = updata['mail']
    email_pw = updata['password']
    up.close() #close json file
        
    return email_user, email_pw

#Step 1 : Login Mail
#Using Gmail - can't set 2 step verification and need to set less secure option in mail prior to using

user, pw = get_mail(mail_id)

mail = imaplib.IMAP4_SSL('imap.gmail.com',993)
mail.login(user,pw)
print('Login Successfully')

time.sleep(2)

#Step 2 : Use pyautogui - pop up to ask for which transactions to take
    #Category - ค่าไฟฟ้า หรือ ประปา
ch_cat = pyautogui.confirm(text='กรุณเลือกชนิดของบิลที่ต้องการ Download', title='Choice', buttons=['ค่าไฟฟ้า', 'ค่าน้ำประปา'])
choice_cat = int(utility_cat[ch_cat]) #ประเภทของข้อมูลที่ต้องการ download
print(f'กำลังดึงข้อมูลเกี่ยวกับค่าใช้จ่ายส่วน : {ch_cat}\n')

    #Type - ประเภทของเอกสารที่ต้องการซื้อ 
ch_type = pyautogui.confirm(text='กรุณเลือกชนิดของบิลที่ต้องการ Download', title='Choice', buttons=['ใบแจ้งหนี้', 'ใบกำกับภาษี'])
choice_type = utility_type[ch_type] #ชนิดของข้อมูลที่ต้องการ download
print(f'ข้อมูลที่ดึง เป็นข้อมูลประเภท : {choice_type}\n')

    #Number - จำนวนเอกสาร / เมล์ที่ต้องการ download
mailnum = pyautogui.prompt(text='จำนวนเมล์ที่ต้องการ Download - กรุณากรอกเป็นตัวเลข', title='Numbers#' , default='1')
print(f'ดึงเมล์เป็นจำนวน : {mailnum} เมล์')

def filter_mail(mail, cat_choice, type_choice): #Get specific data from mail - using search

    mail.select('Inbox') #search where to search before actual search

    mail_from = utility['mail'][cat_choice] #Category of mail - Eletric or Water
    mail_head = utility[type_choice][cat_choice] 
    search_info = '(FROM "' + mail_from +'" SUBJECT "' + mail_head +'")'

    result, data = mail.search(None, search_info)

    return result, data

def get_nummails(mail, m_num, mail_data): #Get only wanted mails - sort mail by created time - get the lastest one

    ids = mail_data[0] # data is a list.
    id_list = ids.split() # ids is a space separated string

    #Identify number of mails want to download 
    num = -int(m_num) #Make number negative since want the last message filtered
    print(num)

    set_mea = id_list[num:] # get the latest

    mea_data = []
    for msg in set_mea:
        data = mail.fetch(msg, "(RFC822)")[1] # fetch the email body (RFC822) for the given ID
        raw_email = data[0][1]
        mea_data.append(raw_email)
        # here's the body, which is raw text of the whole email
        # including headers and alternate payloads

    mea_msg = []
    for mails in mea_data:
        raw_string = mails.decode('utf-8')
        email_msg = email.message_from_string(raw_string)
        mea_msg.append(email_msg)
        
    return mea_data, mea_msg

def get_files(message, cat_choice): #Download PDF from mails selected
    count = 0
    for msg in message: 
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
        
            fileName = part.get_filename()
    
            if bool(fileName):
                filePath = os.path.join(download_dir,utility['keep_fdir'][cat_choice],fileName)
                if not os.path.isfile(filePath) :
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
        count+=1
    
        print(f'Download --> {fileName} completed')

    pyautogui.alert(text=f'Download Invoice--> {count} files completed', title='Notification', button='OK') #Pop up to notify completion

def pop_directory(cat_choice): #Pop up location on computure - save time to go seek for files

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

    pop = fr'explorer /select, "{open_dir}"'
    import subprocess
    subprocess.Popen(pop) #หน้าต่างจะ pop up directory ที่เราสร้างไว้ เพื่อให้ไปแก้ไขชื่อไฟล์ก่อน นำเข้าไปไว้ใน folder ที่สร้างเตรียมไว้

time.sleep(5)

result, data = filter_mail(mail, choice_cat,choice_type) #filter mails

mea_data, mea_msg = get_nummails(mail, mailnum, data) #get mail message

get_files(mea_msg, choice_cat) #Get pdf

time.sleep(5)

pop_directory(choice_cat)

