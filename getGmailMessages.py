# thanks to Giselle Valladares, Applications Support Analyst at Google Workspace. Her stackoverflow post provided the Gmail API interactions. 
from __future__ import print_function

import os.path
import quopri
import re
import pdfkit

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import email
import base64 #add Base64
import requests 
import json

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://mail.google.com/"]


def main():
	
	with open("config.json", "r") as configfile:
		config = json.load(configfile)
		print("Read config file successfully")

	email_sender = config["sender"] # Address from which newsletter was received
	webhook_url = config["webhook_url"] # Create webhook in your server integrations and copy webhook URL here
	output_folder = config["output_folder"] # subdirectory of project root
	post_method = config["post_method"] # HTML, PDF, or Discord posts
	webhookMsgLimit = config["webhookMsgLimit"] # this is the Discord webhook size limit
	notification_role = config["notification_role"] # The Discord role ID formatted for a text post

	# region Authorization
	"""Shows basic usage of the Gmail API."""
	
	creds = None

	# The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
	if os.path.exists("token.json"):
		creds = Credentials.from_authorized_user_file("token.json", SCOPES)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				"credentials.json", SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open("token.json", "w") as token:
			token.write(creds.to_json())

		#Filter and get the IDs of the message I need. 
		#I'm just filtering messages that are unread and come from a provided address.
	
	#endregion

	# Call the Gmail API
	try:
		# region Get messages from Gmail
		service = build("gmail", "v1", credentials=creds)
		search_results = service.users().messages().list(userId="me", q=("from:" + email_sender + " is:unread")).execute()
		result_count = search_results['resultSizeEstimate']
		print(search_results)
		# endregion
	except HttpError as error:
		# TODO(developer) - Handle errors from gmail API.
		print(f'An error occurred while accessing Gmail API: {error}')

	# Post newsletters
	try:
		message_subject = "" # empty by default
		message_id_list = [] # empty array, all the messages ID will be listed here
		thread = [] # empty list, threaded messages for webhook

		if result_count > 0:
			unread_messages = search_results['messages']
			# Forward matching mail messages to webhook
			for message in unread_messages:
				postIndex = 0 # index of posts within thread
				message_id_list.append(message['id'])
				
				# get the body of the message
				message_body, message_html, message_subject = get_message(service, message['id'] )
				
				# split long messages into threads
				if len(message_body) > webhookMsgLimit:
					thread = split_message_body(message_body, webhookMsgLimit)
					intro = thread[0]
				else:
					print('Message is shorter than the character limit, no need to split')
					thread.append(message_body)

				# region Save local files
				# Write html from email message into local file
				local_filename = output_folder + re.sub('[^A-Za-z0-9 ]+', '', message_subject)
				save_html_file(message_html, local_filename + ".html")
				
				# Write pdf from html document into local file
				pdf_config = pdfkit.configuration(wkhtmltopdf=config["wkhtmltopdf"])
				pdfkit.from_file((local_filename + '.html'), (local_filename + '.pdf'), configuration=pdf_config)
				# endregion

				# region Send newsletter to webhook	
				send_webhook("__**" + str.upper(message_subject) + "**__\n\r", webhook_url) # Title is prefixed and suffixed with Discord markup, bold and underlined

				# Send newsletter content to webhook
				if post_method == "html":
					# Send html file as attachment to webhook
					requests.post(webhook_url, files={"file": open(local_filename + ".html", "rb")})
				
				if post_method == "pdf":
					# Send pdf file as attachment to webhook
					requests.post(webhook_url, files={"file": open(local_filename + ".pdf", "rb")})
				
				if post_method == "thread":
					# Send email to webhook as sequence of posts
					for post in thread:
						postIndex = postIndex + 1
						newHeader = "**START POST**\n\r*Post # {index} of {total}*\n\r"
						post = newHeader.format(index = postIndex, total = len(thread)) + post
						send_webhook(post, webhook_url)
				# endregion

				# Mark these mail messages as 'Read'
				service.users().messages().modify(userId="me", id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
				
			# Send notification to specified role
			send_webhook(notification_role, webhook_url)

			print('(thanks!)>  -^^,--,~')

			return message_id_list
		else:
			print('There were 0 results for that search string')
			return ""
	except Exception as error:
		print(f'An error occurred: {error}')

# function to get the body of the message, and decode the message
def get_message(service, msg_id):

	try:
		message_list=service.users().messages().get(userId='me', id=msg_id, format='raw').execute()

		msg_raw = base64.urlsafe_b64decode(message_list['raw'].encode('ASCII'))

		msg_str = email.message_from_bytes(msg_raw)

		msg_subject = msg_str['Subject']

		content_types = msg_str.get_content_maintype()
        
		#how it will work when is a multipart or plain text

		if content_types == 'multipart':
			part1, part2 = msg_str.get_payload()

			print("This is the message body, plaintext:")
			print(part1.get_payload())
			msg_part = quopri.decodestring(part1.get_payload())
			decoded_msg = msg_part.decode("utf-8")

			print("This is the message body, html:")
			print(part2.get_payload())
			msg_part = quopri.decodestring(part2.get_payload())
			decoded_html = msg_part.decode("utf-8")

			return decoded_msg, decoded_html, msg_subject
		else:
			print("This is the message body plain text:")
			# print(msg_str.get_payload())
			return '', quopri.decodestring(msg_str.get_payload()).decode("utf-8"), msg_subject

	except HttpError as error:
		# TODO(developer) - Handle errors from gmail API.
		print(f'An error occurred: {error}')

def split_message_body(message_body, webhook_limit):
	
	thread = []
	
	try:
		print(len(message_body))
		paragraphs = message_body.split("\n") # split fulltext on newlines
		new_post = "" # blank post
		post_script = "\n\r**END POST**\n\r"

		for p in paragraphs:
			if p.__contains__('Unsubscribe '): # Skip footer paragraphs
				continue
			
			# Add paragraph to new_post if it stays under webhook_limit
			if (len(p) + len(new_post) + len(post_script)) < webhook_limit:
				new_post = new_post.__add__("\n\r" + p)
			
			# Stop adding paragraphs to new_post
			if (len(p) + len(new_post) + len(post_script)) >= webhook_limit:
				print('Reached webhook limit on this block:')
				print(new_post)
				print(len(new_post))
				new_post = new_post + post_script
				thread.append(new_post)
				new_post = ""

		return thread
	
	except HttpError as error:
		# TODO - handle errors from splitting message body
		print(f'An error occurred: {error}')
	
def send_webhook(newPost, url):
	
	## Text only:
	data = {
    	"content" : newPost
		}

	result = requests.post(url, json=data)

def save_html_file(html_string, html_filepath):

	# Creating an HTML file 
	Func = open(html_filepath,"w") 

	# Adding input data to the HTML file 
	Func.write(html_string) 

	# Saving the data into the HTML file 
	Func.close()

if __name__ == "__main__":
  main()