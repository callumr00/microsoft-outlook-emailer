# Standard Library Modules
import os

# Third Party Modules
import pandas as pd
import win32com.client
from tqdm import tqdm

file_name = 'recipients.csv'
df = pd.read_csv(os.path.join(os.getcwd(), file_name))

def clean_data(df):
    '''
    Returns a DataFrame after removing all rows that are duplicates,
    have null values or have an invalid email address

    Arguments:
    df - DataFrame to be cleaned
    '''
    
    df = df.drop_duplicates(keep='first')

    df = df.dropna(how='any')

    df = df[df['email'].str.contains('@')]

    df = df.reset_index()
    df = df.drop('index', axis=1)

    print(len(df))

    # df.to_csv('cleaned.csv', index=False)

    return df


df = clean_data(df)

outlook = win32com.client.Dispatch('Outlook.Application')

# Read and store text file data to later construct the email
with open('Email Style.txt', encoding='utf-8') as file:
    email_style = file.read()

with open('Email Introduction.txt', encoding='utf-8') as file:
    email_introduction = file.read()

with open('Category 1.txt', encoding='utf-8') as file:
    category_one_details = file.read()

with open('Category 2.txt', encoding='utf-8') as file:
    category_two_details = file.read()

with open('Email Signature.txt', encoding='utf-8') as file:
    email_signature = file.read()


def send_email(recipient_category, recipient_name, recipient_email):
    '''
    Creates an email, customizes the subject, body and attachments and then
    sends to recipient 

    Arguments:
    recipient_category - Category of recipient (Retailer / Delegate / 
                         Developer / Franchise Partner)
    recipient_email    - Email of recipient
    recipient_name     - Name of recipient
    '''

    email_size = 0x0
    email = outlook.CreateItem(email_size)

    email.Subject = 'Email Subject'
    email.To = recipient_email

    # Set email_details
    if recipient_category == '1':
        email_details = category_one_details

    if recipient_category == '2':
        email_details = category_two_details


    email.HTMLBody = f'''
        <html>
        {email_style}
        <body>
        <p>Dear {recipient_name},</p>
        {email_introduction}
        {email_details}
        {email_signature}
        </body>
        </html>
        '''

    # email.Attachments.Add()

    email.Display()

    # email.Send()


def mass_email(batch_size=500, last_email=None):
    '''
    Repeatedly calls send_email() for each row in the DataFrame, stopping
    after either reaching the end of the DataFrame or the end of the batch

    Arguments:
    batch_size - The size of a single batch (default=500)
    last_email - The last email sent in the previous batch (default=None)
    '''

    # Set the starting row to the row after the last email sent
    if last_email == None:
        starting_row = 0
    else:
        starting_row = df[df['email'] == last_email].index[0] + 1


    # Call send_email() for each row in the DataFrame, up to [batch_size] rows
    print(f"\nFirst email: {df.iloc[starting_row]['email']}\n")

    if batch_size > len(df[starting_row:]):
        for i in tqdm(range(starting_row, starting_row + len(df[starting_row:]))):
            print(f"{df.iloc[i]['category']}, {df.iloc[i]['name']}, {df.iloc[i]['email']}")
            send_email(
                recipient_category = df.iloc[i]['category'],
                recipient_name = df.iloc[i]['name'],
                recipient_email = df.iloc[i]['email']
            )
        print(f"\nLast email: {df.iloc[starting_row + len(df[starting_row:]) - 1]['email']}\n")
    else:
        for i in tqdm(range(starting_row, starting_row + batch_size)):
            send_email(
                recipient_category = df.iloc[i]['category'],
                recipient_name = df.iloc[i]['name'],
                recipient_email = df.iloc[i]['email']
            )
        print(f"\nLast email: {df.iloc[starting_row + batch_size - 1]['email']}\n")


mass_email(
    batch_size=500,
    last_email=None
)