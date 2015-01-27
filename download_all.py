# Download ALL attachments from GMAIL
# Make sure you have IMAP enabled in your gmail settings.
# If you are using 2 step verification you may need an APP Password.
# https://support.google.com/accounts/answer/185833

# imap protocol doc: http://tools.ietf.org/html/rfc3501

import re
import email
import hashlib
import getpass
import imaplib
import os
from collections import defaultdict, Counter
import platform


def make_attachment_directory(path):
    full_path = os.path.join(path, 'attachments')
    if 'attachments' not in os.listdir(path):
        os.mkdir(full_path)

    return full_path


def get_gmail_labels(mail_object):
    """
    :param mail_object: --> imap object
    :return: a list of strings representing valid email box labels
    """
    pattern = re.compile(r'.*\".*\"\s+\"(?P<label>.*)\".*')

    response, data = mail_object.list()

    labels = []
    for itm in data:
        search = re.search(pattern, str(itm))
        labels.append(search.group('label'))

    return labels


def get_gmail_messages_with_attachments_by_label(mail_object, label):
    try:
        mail_object.select('\"' + label + '\"')
        # typ, data = mail_object.search(None, 'ALL')
        response, data = mail_object.search(None, '(X-GM-RAW "has:attachment")')
        print("Return [{0}]: Found {1} messages containing attachments under \"{2}\"".format(response,
                                                                                             len(data[0].split()),
                                                                                             label))
    except Exception as e:
        response, data = 'FAIL', ['', ]
        print("Return [{0}]: {1}".format(response, e))

    return data


def gmail_login(username, password):
    """
    Attempt to login to GMail.
    :param username -> string
    :param password -> string
    :return a mail object if success, else None
    """

    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        response, account_details = mail.login(username, password)
    except imaplib.IMAP4.error as e:
        mail = None
        response, account_details = 'FAIL', [str(e), ]

    print("Return [{0}]: {1}".format(response, account_details[0]))

    return mail


def get_gmail_attachments_for_message_ids(mail, message_IDs, filter_set=None):

    # setup caching of names/data gleaned from attachments
    fileNameCounter = Counter()
    fileNameHashes = defaultdict(set)

    # ensure the directory where the files will be stored go into an attachments folder
    directory = make_attachment_directory('.')

    # Fetch the messages w/ attachments and download
    for message_ID in message_IDs:
        response, message_parts = mail.fetch(message_ID, '(RFC822)')
        # print("Return [{0}]: {1}".format(typ, messageParts[0][1]))
        if 'OK' not in response:
            print('Error fetching mail.')

        emailBody = message_parts[0][1]

        # Message decoding for python 3 is bytes, python 2 expects str
        message_from = email.message_from_bytes
        if type(emailBody) == str:
            # python 2 expects str
            message_from = email.message_from_string

        message = message_from(emailBody)
        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()

            if filename is not None:
                if filter_set is not None and not filename.lower().endswith(filter_set):
                    continue
                filename = ''.join(x for x in filename.splitlines())

            if filename:
                payload = part.get_payload(decode=True)
                if payload:
                    x_hash = hashlib.md5(payload).hexdigest()
                    if x_hash in fileNameHashes[filename]:
                        print('\tSkipping duplicate file: {file}'.format(file=filename))
                        continue
                    fileNameCounter[filename] += 1
                    fileStr, fileExtension = os.path.splitext(filename)
                    if fileNameCounter[filename] > 1:
                        new_fileName = '{file}({suffix}){ext}'.format(suffix=fileNameCounter[filename],
                                                                      file=fileStr,
                                                                      ext=fileExtension)
                        # print('\tRenaming and storing: {file} to {new_file}'.format(file=filename,
                                                                                    # new_file=new_fileName))
                    else:
                        new_fileName = filename
                        # print('\tStoring: {file}'.format(file=filename))
                    fileNameHashes[filename].add(x_hash)
                    file_path = os.path.join(directory, new_fileName)
                    if os.path.exists(file_path):
                        print('\tExists in destination: {file}'.format(file=new_fileName))
                        continue
                    try:
                        with open(file_path, 'wb') as fp:
                            fp.write(payload)
                    except Exception as e:
                        print(type(e))
                        print('Could not store: {file} (invalid name or path under {op_sys}.'.format(
                            file=file_path,
                            op_sys=platform.system()))
                else:
                    print('Attachment {file} was returned as type: {ftype} skipping...'.format(file=filename,
                                                                                               ftype=type(payload)))
                    continue

if __name__ == "__main__":

    # Get Gmail login credentials
    username = input('Enter your GMail username: ')
    password = getpass.getpass('Enter your password: ')

    mail = gmail_login(username, password)

    # get message IDs for the Gmail label of interest
    message_IDs = get_gmail_messages_with_attachments_by_label(mail, '[Gmail]/All Mail')
    message_IDs = message_IDs[0].split()

    # reduce what is actually downloaded to extensions you care about
    # - NOTE: if only one item a set is written like this, filters = (".jpg", ) <-- note the comma
    filters = (".jpg", ".gif")

    # filter out files of a specific type
    get_gmail_attachments_for_message_ids(mail, message_IDs, filter_set=filters)

    # close up shop
    mail.close()
    mail.logout()