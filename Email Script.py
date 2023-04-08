import win32com.client
import pandas as pd
import datetime

# Set up Outlook application and folders
outlook = win32com.client.Dispatch('Outlook.Application')
inbox_folder = outlook.GetNamespace('MAPI').GetDefaultFolder(6)
sent_folder = outlook.GetNamespace('MAPI').GetDefaultFolder(5)

# Define email template with placeholders for variables
email_template = """
Hello {name},

My name is Matt Wassmer and I represent NxtGen Events, a provider of high-performance audio-visual services nationwide. With over 20 years of experience, we have been delivering flawless on-site AV support at venues across the country, surpassing in-house AV pricing by 20-30%. Our clients consider us an integral part of their team, and we have been helping them produce successful events for over two decades.

I am writing to inquire about the RFP list for the {event_name} occurring on {event_date}. By partnering with us, we can help you save budget dollars while ensuring worry-free, world-class service. We take pride in building long-lasting relationships with our clients, as evidenced by the positive feedback we have received from clients nationwide.

Please find below a few examples of what our clients have said about us:

{client_feedback}

Thank you for considering NxtGen Events.

Matt Wassmer 
Vice President/Managing Partner
NxtgenEvents, Powered by ASAV
Nationwide Event Technology Services
Office/Mobile: 901-238-9074
Fax: 866-331-3962
https://nxtgenevents.com/
"""

# Define today's date
today = datetime.date.today().strftime('%Y-%m-%d')

# Read data from CSV file
df = pd.read_csv('emails.csv')

# Define image path
image_path = 'C:\\path\\to\\logo.png'

# Define send email fucntion
def send_email(to, body, subject):
    mail_item = outlook.CreateItem(0)
    mail_item.To = to
    mail_item.Subject = subject
    mail_item.HTMLBody = body + "<br><br><img src='" + image_path + "'>"
    mail_item.Send()

# Main program loop
for index, row in df.iterrows():
    label = row['label']
    date = row['date']
    name = row['name']
    email_address = row['email_address']
    title = row['title']
    event_name = row['event_name']
    event_date = row['event_date']
    event_location = row['event_location']
    sent_date = row['sent_date']
    replied_date = row['replied_date']

    # Skip rows that do not meet the criteria
    if label != 'auto' or date != today:
        continue

    # Search sent folder for previous emails to recipient
    sent_items = sent_folder.Items.Restrict("[To] = '{}'".format(email_address))

    # Check if recipient has been sent an email in the last 30 days
    if len(list(sent_items)) > 0:
        last_sent_date = max([item.SentOn.date() for item in sent_items])
        if (datetime.date.today() - last_sent_date).days < 30:
            # Update CSV file with sent date and date variable
            df.at[index, 'sent_date'] = last_sent_date.strftime('%Y-%m-%d')
            df.at[index, 'date'] = (datetime.datetime.today() + datetime.timedelta(days=7)).strftime('%Y-%m-%d')
            continue

    # Search inbox for previous emails from recipient
    inbox_items = inbox_folder.Items.Restrict("[From] = '{}'".format(email_address))

    # Check if recipient has replied in the last 30 days
    if len(list(inbox_items)) > 0:
        last_reply_date = max([item.ReceivedTime.date() for item in inbox_items])
        if (datetime.date.today() - last_reply_date).days < 30:
            # Update CSV file with replied date and label variable
            df.at[index, 'replied_date'] = last_reply_date.strftime('%Y-%m-%d')
            df.at[index, 'label'] = 'update'
            continue

    # Create email body from template
    client_feedback = """\
    “It is with much enthusiasm that I recommend the services of NxtgenEvents.  They know our programs like the back of their hands, anticipate our needs, are excellent in trouble shooting onsite and competitively priced. The crew is very professional and experts in their field.  In fact, since they excel in their services, it’s the least challenging portion of my job. Again, I highly recommend NxtgenEvents and their highly experienced and professional team”.
    Bridgette Brigham, Senior Conference Manager, Technical Association of the Paper & Pulp Industry (TAPPI)
     
    “I wanted to let you know that your team has done an exceptional job this week!  They have been so on top of it and I think this conference ran more smoothly than any conference we’ve ever done. My team here and I are very grateful!”  Lyn Sholl, CMP, Executive Director, American Filtration & Separations Society, (AFS).
      
    “As usual, I am filled with awe and wonder at how well your team performs.  It was so enjoyable to work with Matt and Tiago. They perpetuate the rumors of your company’s professionalism and reliability! I look forward to the “next one”. 
    Carrie Winchman, Sr. Association Manager, New England Development Research Association (NEDRA). 
     
    “We highly recommend working with NxtgenEvents! We’ve worked with Bill and his team for both of our conferences in 2018 and 2021. As we shifted to a hybrid conference in 2021, the NxtgenEvents team guided us through this new-to-us territory seamlessly, from creating our virtual platform and prepping our virtual speakers to pulling off a custom, hybrid networking experience, which was no small task. We appreciate the close-knit feel of his team and the on-site staff, and we look forward to working with them for our future conferences.”
    Melody Kitchens, Director of Marketing and Events, Crossroads: An Artia Solutions Conference"""

    body = email_template.format(name=name, event_name=event_name, event_date=event_date, 
                                 event_location=event_location, client_feedback=client_feedback)

    subject = "AV RFP for the {} in {} on {}".format(event_name, event_location, event_date)

    # Send email
    send_email(email_address, body, subject)

    # Update CSV file with sent date and label variable
    df.at[index, 'sent_date'] = today

# Save updated CSV file
df.to_csv('emails.csv', index=False)

