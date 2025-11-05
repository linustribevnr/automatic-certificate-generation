# Automated Certificate Generator & Emailer

## Usage instructions:

> 📝 **NOTE:** Please follow the below instructions properly to avoid any errors.

- Install the necessary Python packages by running `pip install -r requirements.txt` command in the command line, in the root directory. Use Python version 3 or greater else you need to run `pip install json smtplib ssl email` also.

- In the root directory, create a `.env` file and add the following environment variables, which are needed for sending emails.
```
email="USE_YOUR_EMAIL_FOR_TESTING"
email_app_password="USE_APP_PASSWORD"
email_host="smtp.gmail.com"
email_port="465"
```
In the above environment variables, make sure you replace values of `email` and `email_app_password`.

- Inside the `data` folder, make sure to keep the following 2 files:
  - Participant's excel sheet (make sure it has a `.xlsx` or `.xls` extension).
  - Certificate template (make sure it has a `.png` extension).

- In `data/input_data.json` file, mandatorily fill in the following with the correct details:
```json
{
    "participants_excel_file_name": "FILE_NAME_WITH_EXTENSION",
    "email_subject": "EMAIL_SUBJECT_GOES_HERE",
    "email_body": "EMAIL_BODY_GOES_HERE",
    "certificate_template_filename": "TEMPLATE_FILE_NAME_WITH_PNG_EXT",
    "img_link":"LINK_TO_IMAGE_IN_EMAIL",
    "print_date": false
}
```
For the email body, seperate new lines with a `<br/>`. Note that in the `participants_excel_file_name` field, the filename should be the one present inside `data` folder and only specify the file name with extension, not the file path. Similar rule applies for `certificate_template_filename` field.

- The participants excel file must have atleast `Name` and `Email` columns. Note that these column names are *case-sensitive*.

## MUST DO THIS FOR GOOD FORMATTING:
- we need to add a sample name in the excel file, that is very large and just enough to fill the whole name placeholder line. If this large name is formatted well, then all other names will be formatted good(to the center).

## Generating Certificates & Sending Emails:

- Run `main_script.py` file using the following command in the command line at the root directory: `python main_script.py`.

- A window with the certificate template will appear, do the following:
  - Double click for the first time. The place where you double clicked will indicate the **top left corner** for the NAME on the certificate.
  - Double click for the second time, if you want to print date on certificate. The place where you double clicked will indicate the **top left corner** for the DATE on the certificate.
  - Now, hit `ENTER`.

- The script will then generate the certificates and send emails. For the first generated certificate, a new window will open to show that. Make sure you close it to continue generating the remaining certificates.

- Open the excel sheet and you will now see two new columns: `Email Sent` and `Remarks`. To run the script again with the same excel sheet, delete these columns and follow the same above steps. If `Email Sent` column is present and is set to `Yes`, email wont be sent to that mail id. Also, make sure to delete the `certificates` folder while re-running to avoid any naming conflicts.

- Before running the main script, **make sure the excel sheet is not open in your system**.

## Steps to generate app password for gmail:

- Follow this link: [Create and use app passwords](https://support.google.com/mail/answer/185833?hl=en-GB).
