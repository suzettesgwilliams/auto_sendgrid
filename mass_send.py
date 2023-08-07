import os, sys
import csv
import pandas as pd
sys.path.append(os.path.dirname(__file__))

import logging
from python_http_client import exceptions 
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, BatchId

from openpyxl import load_workbook

from my_secrets import SendGridSecrets

east_coast_count = 0
midwest_coast_count = 0
west_coast_count = 0

def get_batch_id(id_filename: str):
    api_key = SendGridSecrets.api_key
    sendgrid_client = SendGridAPIClient(api_key)
    try:
        response = sendgrid_client.client.mail.batch.post()
    except exceptions.BadRequestsError as e:
        print(e.body)
        exit()
  
    batch_id = response.to_dict["batch_id"]
    if batch_id is None:
        raise (f"No batch ID generated")

    with open(id_filename, "w") as f:
        print(batch_id, file=f, end="")

    return batch_id

def format_email():
    with open('emails.csv', 'r') as csv_file:
        reader = csv.reader(csv_file)
        next(reader) # skips header row
        # reads csv file and pulls out important data into arrays
        for row in reader:
            emails = row[0]
            names = row[1]
            templates = row[2]
            unixtimes = row[3]
            batch_ids = row[4]
        # creates email instance for this specific recipient to schedule send    
        send_mass(emails, names, templates, unixtimes, batch_ids)
            # if sent > 24 then dont send
        # update spreadsheet +1 with pandas to indicate sent and date
        # update counters for east, midwest, and west coast based on csv info

def send_mass(
    emails,
    names,
    templates,
    unixtimes,
    batch_ids,
):
    # call send single on every row in the arrays
    for (email, name, template, unixtime, batch_id) in [emails, names, templates, unixtimes, batch_ids]:
        send_single(email, name, template, unixtime, batch_id)


def send_single(
    email: str,
    name: str,
    template: str = "",
    unixtime: int = None,
    batch_id: str = None,
):

    # determines what template in sendgrid to use for this email
    if template == "hello":
        template_id = SendGridSecrets.hello_template_id
    elif template == "user_update":
        template_id = SendGridSecrets.user_update_template_id
    elif template == "user_feedback":
        template_id = SendGridSecrets.user_feedback_template_id
    else:
        raise Exception(f"Unknown template {template}")

    if template in [
        "hello",]:
        from_email = SendGridSecrets.hello_from_email
    elif template in ["hello","user_update"]:
        from_email = SendGridSecrets.hello_from_email

    # creates message
    message = Mail(
        from_email=from_email,
        to_emails=email,
    )
    message.dynamic_template_data = {
        "name": name,
    }
    message.template_id = template_id
    if batch_id is not None:  
        message.batch_id = BatchId(batch_id)
    # determines what time to send email
    if unixtime:
        message.send_at = unixtime

    api_key = SendGridSecrets.api_key
    sendgrid_client = SendGridAPIClient(api_key)

    # attempts to contact server to tell what msg to send and when
    try:
        response = sendgrid_client.send(message)
    except exceptions.BadRequestsError as e:
        print(e.body)
        exit()

def cancel_batch(batch_id: str):
    data = {"batch_id": batch_id, "status": "cancel"}

    api_key = SendGridSecrets.api_key
    sendgrid_client = SendGridAPIClient(api_key)
    try:
        response = sendgrid_client.client.user.scheduled_sends.post(request_body=data)
    except exceptions.BadRequestsError as e:
        print(e.body)
        exit()
   
    new_id = response.to_dict["batch_id"]
    new_status = response.to_dict["status"]
    if batch_id != new_id or new_status != "cancel":
        raise (f"Cancellation unsuccessful")

def update_sheet():
    wb = load_workbook('foobar.xlsx')

    # select specific sheet
    ws = wb['SheetX']

    # read the sheet into a pandas dataframe
    df = pd.read_excel('foobar.xlsx', sheet_name='SheetX')

    # modify dataframe (most likely gonna be the sent count + date)
    df['column_name'] = new_values

    # write dataframe back to sheet
    df.to_excel(ws, index=False, header=True)

    wb.save('foobar.xlsx')



def main():
    print('Send test email to ("hello", "hello@email.org")?')
    answer = input("Continue? [y/n]:")

    if answer != "y":
        print("Exiting.")
        exit()
    for email, name in [("hello@email.org", "Giácomo")]:
        send_single(email, name, template="user_update")


if __name__ == "__main__":
    main()
