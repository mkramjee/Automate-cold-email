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

# Define image path 
image_path = 'C:/Users/Matt Wassmer/OneDrive - NxtgenEvents, Powered by ASAV\Desktop\LOGO.png'

# Define email template with placeholders for variables
email_template = """
<p>Hello {name},</p>

<p>My name is Matt Wassmer and I represent NxtGen Events, a provider of high-performance audio-visual services nationwide. With over 20 years of experience, we have been delivering flawless on-site AV support at venues across the country, surpassing in-house AV pricing by 20-30%. Our clients consider us an integral part of their team, and we have been helping them produce successful events for over two decades.</p>

<p>I am writing to inquire about the RFP list for the {event_name}. By partnering with us, we can help you save budget dollars while ensuring worry-free, world-class service. We take pride in building long-lasting relationships with our clients, as evidenced by the positive feedback we have received from clients nationwide.</p>

<p>Please find below a few examples of what our clients have said about us:</p>

<p>&ldquo;It is with much enthusiasm that I recommend the services of NxtgenEvents. They know our programs like the back of their hands, anticipate our needs, are excellent in trouble shooting onsite and competitively priced. The crew is very professional and experts in their field. In fact, since they excel in their services, it&rsquo;s the least challenging portion of my job. Again, I highly recommend NxtgenEvents and their highly experienced and professional team&rdquo;.<br>
- Bridgette Brigham, Senior Conference Manager, Technical Association of the Paper &amp; Pulp Industry (TAPPI)<br>
 <br>
&ldquo;I wanted to let you know that your team has done an exceptional job this week! They have been so on top of it and I think this conference ran more smoothly than any conference we&rsquo;ve ever done. My team here and I are very grateful!&rdquo; <br>
- Lyn Sholl, CMP, Executive Director, American Filtration &amp; Separations Society, (AFS).<br>
 <br>
&ldquo;As usual, I am filled with awe and wonder at how well your team performs. It was so enjoyable to work with Matt and Tiago. They perpetuate the rumors of your company&rsquo;s professionalism and reliability! I look forward to the &ldquo;next one&rdquo;. <br>
- Carrie Winchman, Sr. Association Manager, Association of Fundraising Professionals (AFP)<br>
 <br>
&ldquo;We highly recommend working with NxtgenEvents! We&rsquo;ve worked with Bill and his team for both of our conferences in 2018 and 2021. As we shifted to a hybrid conference in 2021, the NxtgenEvents team guided us through this new-to-us territory seamlessly, from creating our virtual platform and prepping our virtual speakers to pulling off a custom, hybrid networking experience, which was no small task. We appreciate the close-knit feel of his team and the on-site staff, and we look forward to working with them for our future conferences.&rdquo;<br>
- Melody Kitchens, Director of Marketing and Events, Crossroads: An Artia Solutions Conference</p>

<p>Thank you for considering NxtGen Events.</p>

<p>Best Regards, </p>

<p>Matt Wassmer <br>
Vice President/Managing Partner <br>
NxtgenEvents, Powered by ASAV <br>
Nationwide Event Technology Services <br>
Office/Mobile: 901-238-9074 <br>
Fax: 866-331-3962 <br>
https://nxtgenevents.com/</p>

<p></p>

"""

# Define send email fucntion
def send_email(to, body, subject):
    mail = outlook.CreateItem(0)
    mail.To = '{}'.format(email_address)
    mail.Subject = "AV RFP for {} - {}".format(event_name, event_location)
    mail.HTMLBody = body + "<br><br><img src='" + image_path + "'>"
    mail.Send()

# main program
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

    # filter out rows where labels don't match auto, today, and send then move to next step
    if date != today and label != 'auto' and action != 'send':
        continue
    
    else:
        to = '{}'.format(email_address) # define to field
        body = email_template.format(name=name, event_name=event_name) # define message body
        subject = "AV RFP for {} - {}".format(event_name, event_location) # define subject field
        
        send_email (to, body, subject) # call send email function
        
        df.at[index, 'sent_date'] = today # change sent date in csv to today's date
        df.at[index, 'date'] = (datetime.datetime.today() + datetime.timedelta(days=30)).strftime('%m/%d/%Y') # change the next time this row should be looped through to 30 days from today
        df.at[index, 'action'] = 'sent' # change action to sent 


df.to_csv('test4251.csv', index=False) # write changes to csv

print(df)














    
