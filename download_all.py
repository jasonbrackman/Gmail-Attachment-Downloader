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
from collections import defaultdict
import platform


def get_hash(file_to_hash):
    # return unique hash of file
    block_size = 65536
    hasher = hashlib.md5()
    try:
        with open(file_to_hash, 'rb') as handle:
            buf = handle.read(block_size)
            while len(buf) > 0:
                hasher.update(buf)
                buf = handle.read(block_size)
    except IOError as err:
        print(err)
    return hasher.hexdigest()


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
    fileNameCount_dict = defaultdict(int)
    fileNameList_dict = defaultdict(list)

    # ensure the directory where the files will be stored go into an attachments folder
    attachment_directory = make_attachment_directory('.')

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
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            filename = part.get_filename()

            if filename is not None:
                if filter_set is not None and not filename.lower().endswith(filter_set):
                    print(filename)
                    continue
                filename = ''.join(x for x in filename.splitlines())
            if bool(filename):
                filePath = os.path.join(attachment_directory, 'temp.attachment')
                if os.path.isfile(filePath):
                    os.remove(filePath)
                if not os.path.isfile(filePath):
                    try:
                        # print 'Processing: {file}'.format(file=fileName)
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                        x_hash = get_hash(filePath)
                    except Exception as e:
                        print(e)
                        x_hash = get_hash(filePath)

                    if x_hash in fileNameList_dict[filename]:
                        print('\tSkipping duplicate file: {file}'.format(file=filename))

                    else:
                        fileNameCount_dict[filename] += 1
                        fileStr, fileExtension = os.path.splitext(filename)
                        if fileNameCount_dict[filename] > 1:
                            new_fileName = '{file}({suffix}){ext}'.format(suffix=fileNameCount_dict[filename],
                                                                          file=fileStr, ext=fileExtension)
                        else:
                            new_fileName = filename
                        fileNameList_dict[filename].append(x_hash)
                        hash_path = os.path.join(attachment_directory, new_fileName)
                        if not os.path.isfile(hash_path):
                            if new_fileName == filename:
                                print('\tStoring: {file}'.format(file=filename))
                            else:
                                print('\tRenaming and storing: {file} to {new_file}'.format(file=filename,
                                                                                            new_file=new_fileName))
                            try:
                                os.rename(filePath, hash_path)
                            except Exception as e:
                                print(type(e))
                                print(
                                    'Could not store: {file} (invalid filename or path under {op_sys}.'.format(
                                        file=hash_path, op_sys=platform.system()))
                        elif os.path.isfile(hash_path):
                            print('\tExists in destination: {file}'.format(file=new_fileName))
                    if os.path.isfile(filePath):
                        os.remove(filePath)

    # mail.close()
    # mail.logout()

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