import os
import requests 
import win32com.client
import re
import boto3
from botocore.exceptions import NoCredentialsError 
import webbrowser
import configparser as config
import logging
import datetime
import win32timezone
from sendmail import send_mail

#url = "https://s3.us-west-2.amazonaws.com/data.prod.sortly.com/400902cb56115ead1b0eac5a/csvs/963ec423dbe5be4f6ce07340cbeb8bae0570b675.original.csv?response-content-disposition=attachment&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=AKIAINAWTSTHEUFHL3VA%2F20230907%2Fus-west-2%2Fs3%2Faws4_request&X-Amz-Date=20230907T062606Z&X-Amz-Expires=86400&X-Amz-SignedHeaders=host&X-Amz-Signature=56a8abdf6f4b1fa2f8bc622a100f595df453cfdd6407b2ed75e6a144eb9d8b7b"  # Replace with the URL of the file you want to download

# downloaded_file = "../downloadedfile"
# os.makedirs(downloaded_file, exist_ok=True)

config_out= config.ConfigParser()
config_out.read ("Config.ini")
current_directory = os.getcwd()
current_dir = current_directory.replace("\\","/")
log_filename = config_out.get('log_name','download_log')
log_toemail = current_dir + log_filename
filename =str( datetime.date.today() ) + "-downloaded-file.csv"
save_path =  filename     # Replace with the desired local file path and name
email_to =  config_out.get('email_config','log_mailto')
ccmail_to = config_out.get('email_config','log_ccmailto')
url=""
logging.basicConfig(
        filename=log_toemail,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
        )

def download_file(url):
    try:
      
        response = requests.get(url)
        #logging.info(f"{response.text}")
        # Check for any HTTP errors
        #response.raise_for_status()
        if response.status_code ==200:
            with open(save_path, "wb") as file:
                file.write(response.content)
        else:
            logging.info(f"Download error:  {response.text}")
            send_mail(email_to,"Autoupload error/info log","Attached is the autoupload log file.",log_toemail,ccmail_to)     
    except requests.exceptions.RequestException as e:
        print(f"Error: {e}")
        logging.info(f"Downloading Email attachedment error log: {e} ")
        send_mail(email_to,"Autoupload error/info log","Attached is the autoupload log file.",log_toemail,ccmail_to)     
    except requests.exceptions.HTTPError as e:
        print(f"HTTP error occurred: {e}")
        logging.info(f"Downloading Email attachedment error log: HTTP error occurred: {e} ")
        send_mail(email_to,"Autoupload error/info log","Attached is the autoupload log file.",log_toemail,ccmail_to)      
   
def get_awsfile():
        session = boto3.Session( aws_access_key_id="AKIAINAWTSTHEUFHL3VA",aws_secret_access_key="")
        s3= session.client("s3")
        aws3_url = "https://s3.us-west-2.amazonaws.com/data.prod.sortly.com/400902cb56115ead1b0eac5a/csvs/963ec423dbe5be4f6ce07340cbeb8bae0570b675.original.csv"
        s3_parts = aws3_url.split('/')
        bucket_name =s3_parts[2]
        object_key ='/'.join(s3_parts[3:])    
        local_file_path = "c:\hg\csv_export_report.csv"
        try:
            s3.download_file(bucket_name,object_key,local_file_path)
            print(f"File downloaded on {loal_file_path}")
        except NoCredentialsError:
            print("AWS credentials not found. Please configure your AWS credentials.")
        except Exception as e:
            print(f"Error: {str(e)}") 
            
def my_print(txt):
    print(txt)

msg_template="""Welcome {name} to my website {website}"""

def format_message(my_name="jay",my_website="website"):
    my_msg=msg_template.format(name=my_name,website=my_website)
    
    return my_msg
def base_func(*args,**kwargs):
    print(args,kwargs)


     
def get_link():
    try:
        logging.info (f"Downloading current file from email")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 represents the Inbox folder
        subject_to_filter = "[Subject] = 'CSV Export Report'"
    # Retrieve emails with the specified subject
        messages = inbox.Items.Restrict( subject_to_filter )
        filtered_emails = [email for email in inbox.Items if "CSV Export Report" in email.Subject]
        latest_email = sorted(filtered_emails, key=lambda x: x.ReceivedTime, reverse=True)[0]
        link_pattern = r'https?://\S+'
        specific_words =["original.csv"]
        
        # for email in messages:
        #     print(f"Subject: {email.Subject}")      
        email_body = latest_email.Body
        #pattern = r"https?://\S*?(" + "|".join(map(re.escape, specific_words)) + r")\S*"
        #print (f"this is pattern {pattern}")
        #match = re.search(link_pattern, email_body)
        matches = re.findall(link_pattern, email_body, flags=re.IGNORECASE)
            # if matches:
            #     link_url = match.group(0)
            # else:
            #     print("No link found in the email.")
            #     #exit()
            #print (f"this is matches {matches}")
        for match in matches:
            if match.find("s3.us-west-2.amazonaws.com") !=-1:
                print(f"this is the link {match}")
                download_file(match)
                break
                
                    
            # link_url=""
            # path=  "C:/hg/"
            # file_name = os.path.basename(link_url)  # Extract the file name from the URL
            # save_path = os.path.join(path, file_name)
            # print(f"File '{file_name}' downloaded and saved to '{save_path}'.")
            # response = requests.get(match)

            # if response.status_code == 200:
            #     with open(save_path, 'wb') as file:
            #         file.write(response.content)
            #         print(f"File '{file_name}' downloaded and saved to '{save_path}'.")
            # else:
            #     print(f"Failed to download the file. Status code: {response.status_code}")
        #send_mail(email_to,"Autoupload log","Attached is the autoupload log file.",log_toemail,ccmail_to)
        logging.shutdown()    
        del filtered_emails
        del messages
        del inbox
        del outlook
        del latest_email
        
    except Exception as e:
            print(f"Error: {str(e)}") 
            logging.info(f"Error : {str(e)}")
            send_mail(email_to,"Autoupload error/info log","Attached is the autoupload log file.",log_toemail,ccmail_to)     
           
                
        
            