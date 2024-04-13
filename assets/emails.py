import os
import pandas as pd
import win32com.client
import pythoncom


def email_reply(description: str, subject: str, display: str):
    """
    Introduction
    ------------
        - email_reply search all over the Outlook inbox messages for a subject defined by the user through the parameter.
        - By the time it finds it, enters the email, displays the reply Outlook screen, puts all recipients and carbon copies emails, places a standard body including a signature, and then attaches the last Excel File from the excel_reports folder.

    Keep in mind
    ------------
        - Subject must exist in the Outlook Inbox messages and an Excel file in the excel_reports folder.
    
    Parameters
    ----------
        - subject: String, no default values. Introducing the subject name of the email.
        - display: String, no default values. Selecting whether to visualize the email before sending it or autosend.
    """
    
    # Excel folder for file attachment
    path = os.getcwd().replace('\\', '/') + '/files'
    # Signature folder for signature attachment
    signature = os.getcwd().replace('\\', '/') + '/signature'

    # If files folder does not exist, it will create a new one
    if not os.path.exists(path):
        os.mkdir(path)

    else:
        # If either the description, subject, display parameters are empty and no files were found within the files folder program will not initialize
        if description != "" and subject != "" and display != "" and os.listdir(os.getcwd().replace('\\', '/') + '/files') != []:
            #Calling CoInitialize
            pythoncom.CoInitialize()
            
            # Initializes Outlook
            Outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")

            # Redirects to Inbox
            inbox = Outlook.GetDefaultFolder(6)
            
            # Checks all elements in Inbox
            messages = inbox.Items

            # Obtains only last message or most current lookup
            message = messages.GetLast()

            # Loops through all messages, opens and select on reply to the indicated by the user
            for message in messages:
                # If the messahe is found and there is a file within the files folder
                if message.Subject == subject and os.listdir(path) != []:
                    # Select on Reply All in the email found
                    reply = message.ReplyAll()
                    # Places an standard body and attaches the last Excel File from the excel_reports folder
                    reply.HTMLBody = f"""
                        <html>
                            <head></head>
                                <body>
                                    <p>Hello there,
                                    <br>
                                    <br>
                                    {description}
                                    <br>
                                    <br>
                                    I will be attentive to your comments, have an excellent day!
                                    <br>
                                    <br>
                                    Best Regards,
                                    <br>
                                    </p>
                                </body>
                        </html>""" + reply.HTMLBody
                    reply.Attachments.Add(f'{path}/{str(os.listdir(path)[-1])}')
                    
                    # Would you prefer to check the contents of the email before sending it?
                    if display == 'Y' or display == 'y':
                        reply.Display()
                    
                    # Sending email straightforwardly
                    elif display == 'N' or display == 'n':
                        reply.Send()

        # In case of empty parameters
        else:
            print("Enter a subject or place a file within the folder.")
    

def email_send(subject: str, priority: int, description: str, display: str):
    """
    Introduction
    ------------
        - email_send uses an Excel file for placing all the emails through a for loop, within the To whitespace.
        - The emails included in the Emails.xlsx file will be receiving an email with an standard format.
        - The standard format uses a HTML syntax.

    Keep in mind
    ------------
        - Subject must exist in the Outlook Inbox messages and an Excel file in the excel_reports folder.
        - This process should be executed (preferably) before exporting files into a specific Colombia Audiences pre-defined path. 
        - Importance: 2 is in case the priority of the email is High and Display can use 'Y', 'y'for showing email before sending it.
    
    Parameters
    ----------
        - subject: String, no default values. Introducing the subject name of the email.
        - priority: Integer, no default values. Introducing the email priority.
        - description: String, no default values. 
        - Display: String, no default values. Indicating in this option if the user rather prefers checking the email before sending it or sending it directly.
    """
        
    # Reading Excel file
    df = pd.read_excel(f'{os.getcwd().replace("\\", "/")}/emails/Emails.xlsx', sheet_name = "Emails", engine = 'openpyxl')

    # Excel folder for file attachment
    path = os.getcwd().replace('\\', '/') + '/files'

    # Signature path
    signature = os.getcwd().replace('\\', '/') + '/signature'

    # If files folder does not exist, it will create a new one
    if not os.path.exists(path):
        os.mkdir(path)

    # If signature folder does not exist, it will create a new one
    elif not os.path.exists(signature):
        os.mkdir(signature)

    # Signature file
    signature_file = signature + '/' + os.listdir(os.getcwd().replace('\\', '/') + '/signature/')[0]

    if os.listdir(path) != [] and not df.empty and os.listdir(signature) != []:
        #Calling CoInitialize
        pythoncom.CoInitialize()
        # Initializes Outlook
        Outlook = win32com.client.Dispatch('outlook.application')

        # Iterating through all values from the column Email to be included within the To whitespace
        for emails in df['Email']:
            # Splitting name by period and switching string to Capital
            name = emails.partition('.')[0].capitalize()
            # Selecting on New Email
            mail = Outlook.CreateItem(0)
            # Placing all emails from the Excel file in the To whitespace
            mail.To = emails
            # Subject defined by the user
            mail.Subject = subject
            # HTML syntax about the content of how the email will look like
            mail.HTMLBody = f"""                       
                <html>
                    <head>
                    </head>
                    <body>
                        <main>
                            <p>Hello {name}, 
                            <br>
                            <br>
                            {description}
                            <br>
                            <br>
                            We greatly appreciate from your time and attention,
                            <br>
                            <br>
                            Regards,
                            <br>
                            <br>
                            <img src = "file:{signature_file}" width = 25%>
                            </p>
                        </main>
                    </body>
                </html>"""
            
            # Attaching files
            mail.Attachments.Add(path + '/' + str(os.listdir(path)[-1]))
            
            # Defining priority
            mail.Importance = priority

            # Would you prefer to check the contents of the email before sending it?
            if display == 'Y' or display == 'y':
                mail.Display()

            # Sending email straightforwardly
            elif display == 'N' or display == 'n':
                mail.Send()

    else:
        # Empty options
        print("Either the Emails Excel file is empty, no signatures were included or no files were attached in the folder.")