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
    blocksize = 65536
    hasher = hashlib.md5()
    try:
        with open(file_to_hash, 'rb') as afile:
            buf = afile.read(blocksize)
            while len(buf) > 0:
                hasher.update(buf)
                buf = afile.read(blocksize)
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


if __name__ == "__main__":

    # Get Gmail login credentials
    username = input('Enter your GMail username: ')
    password = getpass.getpass('Enter your password: ')

    mail = gmail_login(username, password)

    # get message IDs for the Gmail label of interest
    message_IDs = get_gmail_messages_with_attachments_by_label(mail, '[Gmail]/All Mail')

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
        message = email.message_from_bytes(emailBody)
        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            fileName = part.get_filename()
            if fileName is not None:
                fileName = ''.join(x for x in fileName.splitlines())
            if bool(fileName):
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

                    if x_hash in fileNameList_dict[fileName]:
                        print('\tSkipping duplicate file: {file}'.format(file=fileName))

                    else:
                        fileNameCount_dict[fileName] += 1
                        fileStr, fileExtension = os.path.splitext(fileName)
                        if fileNameCount_dict[fileName] > 1:
                            new_fileName = '{file}({suffix}){ext}'.format(suffix=fileNameCount_dict[fileName],
                                                                          file=fileStr, ext=fileExtension)
                        else:
                            new_fileName = fileName
                        fileNameList_dict[fileName].append(x_hash)
                        hash_path = os.path.join(attachment_directory, 'attachments', new_fileName)
                        if not os.path.isfile(hash_path):
                            if new_fileName == fileName:
                                print('\tStoring: {file}'.format(file=fileName))
                            else:
                                print('\tRenaming and storing: {file} to {new_file}'.format(file=fileName,
                                                                                            new_file=new_fileName))
                            try:
                                os.rename(filePath, hash_path)
                            except:
                                print(
                                    'Could not store: {file} it has an invalid file name or path under {op_sys}.'.format(
                                        file=hash_path, op_sys=platform.system()))
                        elif os.path.isfile(hash_path):
                            print('\tExists in destination: {file}'.format(file=new_fileName))
                    if os.path.isfile(filePath):
                        os.remove(filePath)

    # mail.close()
    # mail.logout()
