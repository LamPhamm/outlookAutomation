import win32com.client as win32
import datetime
import re
import colorama
from colorama import Fore, Back, Style
colorama.init(autoreset=True)

def main():
    intro()

def intro():
    user_input=1
    while user_input==1 or user_input==2:
        #Ask the user for input
        user_input=int(input(Fore.BLUE+"Enter 1 to see email or 2 to send email or different number to quit: "))
        if user_input==1:
            number_of_messsages=int(input(Fore.BLUE+"Enter the number of messages you want to see: "))
            see_email(number_of_messsages)
        elif user_input==2:
            email=input(Fore.BLUE+"Enter the email: ")
            subject = input(Fore.BLUE+"Enter the subject: ")
            message = input(Fore.BLUE+"Enter the message: ")
            send_email(email,subject,message)

    print("Thank you for using me!")

def see_email(number_of_messages):
    #Create an outlook session
    outlook = win32.gencache.EnsureDispatch('outlook.Application')
    mapi = outlook.GetNamespace("MAPI")

    #The inbox folder corresponds to number 6
    inbox = mapi.GetDefaultFolder(6)
    #Get all of the messages
    messages=inbox.Items
    messages.Sort("[ReceivedTime]", True) #Sort the messages in order from most recent to least recent

    #Iterating the message items and reading emails
    for i,message in enumerate(messages):
        #Get the information from the message
        subject=message.Subject
        date=message.ReceivedTime.date()
        time=message.ReceivedTime.time()
        #Split the body message into a list to get rid of new line characters
        body_message=message.Body
        body_message=re.split('[\n \r]',body_message)

        #Create the datetime
        date=str(date).split("-")
        date=datetime.datetime(int(date[0]),int(date[1]),int(date[2]))

        #Get the email address from each message if possible
        if message.SenderEmailType=='EX':
            try:
                sender=message.Sender.GetExchangeUser().PrimarySmtpAddress
            except:
                sender="Unknown@gmail.com"
        else:
            sender=message.SenderEmailAddress

        #Display the messages
        print(i+1)
        print(Fore.LIGHTGREEN_EX+'\tEmail: ',Fore.LIGHTMAGENTA_EX+sender)
        print(Fore.LIGHTGREEN_EX+'\tDate: ',Fore.LIGHTMAGENTA_EX+date.strftime("%x"),Fore.LIGHTMAGENTA_EX+str(time))
        print(Fore.LIGHTGREEN_EX+'\tSubject: ',Fore.LIGHTMAGENTA_EX+ subject)
        print(Fore.LIGHTGREEN_EX+'\tMessage: ',Fore.LIGHTMAGENTA_EX+' '.join(body_message))

        #Only read up to the number of messages specified by the user
        if i+1==number_of_messages:
            break

def send_email(email,subject,message):
    #Create an outlook session
    outlook = win32.gencache.EnsureDispatch('outlook.Application')
    mapi = outlook.GetNamespace("MAPI")

    #Create an item object for the mail
    mail = outlook.CreateItem(0)

    #These are the attributes of the message
    mail.To=email.lstrip("34m")
    mail.Subject=subject
    mail.Body=message

    #Send the email
    mail.Send()
    print(Fore.BLUE+"Message was sent!")

main()