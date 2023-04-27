import win32com.client as win32
import pandas as pd
import datetime

# Define today's date
today = datetime.date.today().strftime('%m/%d/%Y')

# Read the CSV file
df = pd.read_csv('test4251.csv')

# Set up outlook
outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
sent_folder = namespace.GetDefaultFolder(5)
inbox_folder = namespace.GetDefaultFolder(6)

# main program
for index, row in df.iterrows():
    label = row['label']
    email_address = row['email_address']
    date = row['date']
    name = row['name']
    title = row['title']
    event_name = row['event_name']
    event_month = row['event_month']
    event_date = row['event_date']
    event_location = row['event_location']
    sent_date = row['sent_date']
    replied_date = row['replied_date']
    action = row['action']

    # filter out rows that don't have auto label or today's date
    if label != 'auto' and date != today:
        continue
        
    # scan inbox for email received from recipient in the last 30 days
    inbox_items = inbox_folder.Items.Restrict("[SenderEmailAddress] = '{}'".format(email_address))
    if len(list(inbox_items)) > 0:
        last_reply_date = max([item.ReceivedTime.date() for item in inbox_items])
        if (datetime.date.today() - last_reply_date).days < 30:
            df.at[index, 'replied_date'] = last_reply_date.strftime('%m/%d/%Y') # add email received date to csv
            df.at[index, 'label'] = 'respond' # change label from auto to respond
            df.at[index, 'action'] = 'skip' # change action to skip this row for future steps
    else:
        df.at[index, 'action'] = 'send' # change action to send email to recipient in this row
        next
        
    # scan sent folder for email sent to recipient in the last 30 days
    sent_items = sent_folder.Items.Restrict("[To] = '{}'".format(email_address))
    if len(list(sent_items)) > 0:
        last_sent_date = max([item.SentOn.date() for item in sent_items])
        if (datetime.date.today() - last_sent_date).days < 30:
            df.at[index, 'sent_date'] = last_sent_date.strftime('%m/%d/%Y') # add email sent date to csv
            df.at[index, 'date'] = (datetime.datetime.today() + datetime.timedelta(days=7)).strftime('%m/%d/%Y') # change next date these actions should be repeated for this row 
            df.at[index, 'action'] = 'skip' # change action to skip this row for future steps
        else:
            df.at[index, 'action'] = 'send' # change action to send email to recipient in this row

# second main program for recipients with respond label      
for index, row in df.iterrows():
    email_address = row['email_address']
    label = row['label']
    date = row['date']
    name = row['name']
    title = row['title']
    event_name = row['event_name']
    event_month = row['event_month']
    event_date = row['event_date']
    event_location = row['event_location']
    sent_date = row['sent_date']
    replied_date = row['replied_date']
    action = row['action']

    # filter out rows that don't have respond label or today's date
    if label != 'respond' and date != today:
        continue

        # Search sent folder for emails sent to recipient in the last 30 days
    sent_items = sent_folder.Items.Restrict("[To] = '{}'".format(email_address))

    if len(list(sent_items)) > 0:
        last_sent_date = max([item.SentOn.date() for item in sent_items])
        if (datetime.date.today() - last_sent_date).days < 30:
            df.at[index, 'sent_date'] = last_sent_date.strftime('%m/%d/%Y') # add email sent date to csv
            df.at[index, 'date'] = (datetime.datetime.today() + datetime.timedelta(days=7)).strftime('%m/%d/%Y') # change next date these actions should be repeated for this row 
            df.at[index, 'action'] = 'skip' # change action to skip this row for future steps
        else:
            df.at[index, 'action'] = 'send' # change action to send email to recipient in this row
            
            
df.to_csv('test4251.csv', index=False) # write changes to csv       

print(df)




            
