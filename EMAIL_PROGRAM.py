import imaplib
import email
import time
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import pandas
import warnings
import pwinput
warnings.simplefilter(action='ignore', category=FutureWarning)


#keywords in message for auto response:
global keywords
keywords = {
            "test"
            }



#message to send
mass_email = 'Thank you for your email. We appreciate your interest in contacting us. For further information, please feel free to call us at (***) ***-****. Our staff would be more than happy to answer any questions you may have.'


#email information
print('THIS VERSION HAS BEEN BUILT FOR OUTLOOK --- for **** --- 3_2023\n')
user = input("Enter your email address: ")
nickname = user.split('@')
sender = user

# Prompt user to enter the second input
if nickname[1] == 'outlook.com':
    password = pwinput.pwinput(prompt='Enter your Outlook password: ', mask='*')
elif nickname[1] == 'gmail.com':
    password = pwinput.pwinput(prompt='Enter your gmail password: ', mask='*')
else:
    print("Email type not accepted")


#gmail or outlook
if nickname[1] == 'outlook.com':
    mail = imaplib.IMAP4_SSL('outlook.office365.com', '993')
    session = smtplib.SMTP('smtp.office365.com', '587')
    session.ehlo()
    if session.has_extn('STARTTLS'):
        print('(starting TLS)')
        session.starttls()
        session.ehlo()  # Reidentify ourselves over TLS connection.
    else:
        print('(no STARTTLS)')
elif nickname[1] == 'gmail.com':
    mail = imaplib.IMAP4_SSL('imap.gmail.com', '993')
    session = smtplib.SMTP('smtp.gmail.com', '587')
    session.ehlo()
    if session.has_extn('STARTTLS'):
        print('(starting TLS)')
        session.starttls()
        session.ehlo()  # Reidentify ourselves over TLS connection.
    else:
        print('(no STARTTLS)')
else:
    print("Email type not accepted")


print(user)
#login to email accounts
mail.login(user, password)
print("Logged In")



mail.select("Inbox") # connect to inbox.
_, selected_mails = mail.search(None, "ALL")



def main__():
    df = pandas.DataFrame()
    for num in selected_mails[0].split():
        _, data = mail.fetch(num , '(RFC822)')
        _, bytes_data = data[0]

        #convert the byte data to message
        email_message = email.message_from_bytes(bytes_data)
        print("\n===========================================")

        #access data
        #print("Subject: ",email_message["subject"])
       # print("To:", email_message["to"])
       # print("From: ",email_message["from"])
        #print("Date: ",email_message["date"])
        
        #check if email is equal to todays date
        #check and format todays date from emails
        split_date = email_message["date"].split()
        email_date = str(split_date[1]) + "_"+ str(split_date[2]) + "_" + str(split_date[3])
        

        
        #todays actual date
        today_date = time.strftime("%d_%b_%Y")
        if email_date == today_date:
            mid_df, truefalse, resp = send_email(email_message)
            if truefalse == True and resp == True:
                print("\nAUTOMATED MESSAGE SENT...")
            else:
                df = df.append(mid_df)
            

                    
        else: 
            print("No more emails today...")
            break
            
    message_to_self(df)        
    return df
        
        
def message_contents(x):
    message = x    
    for part in message.walk():
        if part.get_content_type()=="text/plain" or part.get_content_type()=="text/html":
            message = part.get_payload(decode=True)
            #print("Message: \n", message.decode())
            print("==========================================\n")
            return message.decode()

def filter_switch():
   
    toplevel = Toplevel()
    root = tk.Tk()
    root.title("Scrolltext Widget")
    tk.Label(root,text='---Made by WB---',font=("Times New Roman", 12))\
        .grid(row=0,column=0)
    myScrollTextWidget = scrolledtext.ScrolledText(root,wrap=tk.WORD,width=800,height=40,font=("Times New Roman",20))
    myScrollTextWidget.grid(row=1,column=1)
    def printToConsole():
        print(myScrollTextWidget.get("1.0","end-1c"))
    #Buttons
    myButton = tk.Button(root,text="Print to console!",command=printToConsole).grid(row=2,column=1)
    myScrollTextWidget.insert(tk.INSERT, 'THIS VERSION HAS BEEN BUILT TO WORK WITH OUTLOOK\n\nTHERE ARE BUTTONS ON THE LEFT SIDE OF THE SCREEN: \n---ON THIS SCREEN, YOU CAN SELECT ALL EMAILS OR FILTERED EMAILS \n---ON THE NEXT SCREENS, YOU WILL APPROVE SENDING OUT AN EMAIL \n\nIF YOU APPROVE TO SEND AN EMAIL THIS PROGRAM WILL IMMEDIATLY SEND IT \n---THE CURRENT EMAIL MESSAGE IS: \n\n      Thank you for your email.\n      We appreciate your interest in contacting us.\n      For further information, please feel free to call us at (757) 595-3327.\n      Our staff would be more than happy to answer any questions you may have') 
    scrolW=30  
    scrolH=2  
    scr=scrolledtext.ScrolledText(toplevel, width=scrolW, height=scrolH, wrap=tk.WORD)  

    # Create a BooleanVar for the user's response
    var = BooleanVar()

    # Create the "Yes" and "No" buttons and set their commands to set the BooleanVar and close the popup
    button_yes = Button(root, text="See all emails", command=lambda: (var.set(True), toplevel.destroy()))
    button_yes.place(x=10, y=200)
    button_no = Button(root, text="See filtered emails", command=lambda: (var.set(False), toplevel.destroy()))
    button_no.place(x=10, y=250)

    # Wait for the popup to close and return the BooleanVar's value
    w = 1600 # width for the Tk root
    h = 900 # height for the Tk root

    # get screen width and height
    ws = root.winfo_screenwidth() # width of the screen
    hs = root.winfo_screenheight() # height of the screen

    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)

    # set the dimensions of the screen 
    # and where it is placed
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    toplevel.wait_window(toplevel)
    return var.get()
    
def checker_all(em):
    msg = message_contents(em)
    words = msg.split()

    for word in words:
        if mass_email in msg:
            send_em=False
        else:
            send_em=True
            break


                
    return send_em, msg
    
def checker_keywords(em):
    msg = message_contents(em)
    words = msg.split()

    for word in words:
        if mass_email in msg:
            send_em=False
        else:
            if word in keywords:
                send_em=True
                break
            else:
                send_em=False
                
    return send_em, msg
                       
def send_email(em_message):
    final_df = pandas.DataFrame(columns = 
        ['1',
        '2', 
        '3', 
        '4'
        ])
    checker_em = em_message
    if keyword_filter == True:
        TF, msg_data = checker_all(checker_em)
    else:
        TF, msg_data = checker_keywords(checker_em)
    #name
    em_name = em_message["from"]
    em_name = em_name.split()
    em_name = em_name[0:2]
    emname = (str(em_name[0]))
    #subject
    em_sub = em_message["subject"]
    emsub = str(em_sub)
    #email
    em_email = em_message["from"]
    em_email = em_email.split()
    ememail = str(em_email[2:])
    ememail = ememail.replace("<", "")
    ememail = ememail.replace(">", "")
    ememail = ememail.replace("[", "")
    ememail = ememail.replace("]", "")
    ememail = ememail.replace("'", "")
    ememail = str(ememail)
    #run popup
    response = popup(checker_em, msg_data)
    
    if TF == True and response == True and str(ememail) != str(sender):
        default_message(emname, emsub, ememail, checker_em)
    else:
        print('EMAIL FROM: ', em_message["from"],
                ' ABOUT: ', em_message["subject"],
                ' NEEDS TO BE RESPONDED TO MANUALLY')
        final_df = final_df.append({
        '1': 'EMAIL FROM: ',
        '2': em_message["from"],
        '3': ' ABOUT: ',
        '4': em_message["subject"]
            }, ignore_index=True)  
    
    return final_df, TF, response

        
    
def popup(data, msgdata):
    
    From_Subject = str(data["from"])
    Message_Information = str(msgdata)
    toplevel = Toplevel()
    root = tk.Tk()
    root.title("Scrolltext Widget")
    tk.Label(root,text='---Made by WB--- Name and Email: ' + From_Subject,font=("Times New Roman", 12))\
        .grid(row=0,column=0)
    #Define ScrollTextWidget
    #wrap keyword used to wrap around text
    myScrollTextWidget = scrolledtext.ScrolledText(root,wrap=tk.WORD,width=800,height=40,font=("Times New Roman",15))
    myScrollTextWidget.grid(row=1,column=1)
    def printToConsole():
        print(myScrollTextWidget.get("1.0","end-1c"))
    #Buttons
    myButton = tk.Button(root,text="Print to console!",command=printToConsole).grid(row=2,column=1)
    myScrollTextWidget.insert(tk.INSERT, Message_Information)
    scrolW=30  
    scrolH=2  
    scr=scrolledtext.ScrolledText(toplevel, width=scrolW, height=scrolH, wrap=tk.WORD)  

    # Create a BooleanVar for the user's response
    var = BooleanVar()

    # Create the "Yes" and "No" buttons and set their commands to set the BooleanVar and close the popup
    button_yes = Button(root, text="Send the Email", command=lambda: (var.set(True), toplevel.destroy()))
    button_yes.place(x=70, y=200)
    button_no = Button(root, text="Do Not Send the Email", command=lambda: (var.set(False), toplevel.destroy()))
    button_no.place(x=70, y=250)

    # Wait for the popup to close and return the BooleanVar's value
    w = 1600 # width for the Tk root
    h = 900 # height for the Tk root

    # get screen width and height
    ws = root.winfo_screenwidth() # width of the screen
    hs = root.winfo_screenheight() # height of the screen

    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)

    # set the dimensions of the screen 
    # and where it is placed
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    toplevel.wait_window(toplevel)
    return var.get()
    
def default_message(name, subject, ememail, emmsg):
    msg = MIMEMultipart()
    msg['From'] = str(sender)
    msg['To'] = str(ememail)
    msg['Subject'] = ('Response to: ' + subject)
    message = ('Hello there ' + name + ', \n \n' + mass_email + '\n \n' + 'Best Regards, \nWTB Staff and Company')
    msg.attach(MIMEText(message))

    #create session

    session.login(user, password)

    #sendmail
    session.sendmail(str(sender), str(ememail), str(msg))
    print('EMAIL FROM: ', emmsg["from"],
            ' ABOUT: ', emmsg["subject"],
            ' HAS BEEN SENT DEFAULT MESSAGE')
            
            
def message_to_self(dataa):
    msg = MIMEMultipart()
    msg['From'] = str(sender)
    msg['To'] = str(sender)
    msg['Subject'] = ('EMAILS TO RESPOND TO:')
    message = (str(dataa))
    msg.attach(MIMEText(message))

    #create session

    session.login(user, password)

    #sendmail
    session.sendmail(str(sender), str(sender), str(msg))

            
            
keyword_filter = filter_switch()

final_message = main__()
print('\n===========================================\n===========================================\n',final_message, '\n\n----EMAIL OF MANUAL RESPONSES SENT TO SELF----\n')


