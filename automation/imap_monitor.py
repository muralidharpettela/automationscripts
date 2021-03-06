import os

import eventlet

imapclient = eventlet.import_patched('imapclient')

import os.path as path
import sys
import traceback
import logging
from logging.handlers import RotatingFileHandler
import configparser
import email
from time import sleep
from datetime import datetime, time
from pathlib import Path
from update_stock.update_products_data_staging import StagingUpdateProducts

# Setup the log handlers to stdout and file.
log = logging.getLogger('imap_monitor')
log.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '%(asctime)s | %(name)s | %(levelname)s | %(message)s'
)
handler_stdout = logging.StreamHandler(sys.stdout)
handler_stdout.setLevel(logging.DEBUG)
handler_stdout.setFormatter(formatter)
log.addHandler(handler_stdout)
handler_file = RotatingFileHandler(
    'imap_monitor.log',
    mode='a',
    maxBytes=1048576,
    backupCount=9,
    encoding='UTF-8',
    delay=True
)
handler_file.setLevel(logging.DEBUG)
handler_file.setFormatter(formatter)
log.addHandler(handler_file)


# TODO: Support SMTP log handling for CRITICAL errors.


def process_email(mail_, download_, log_):
    """Email processing to be done here. mail_ is the Mail object passed to this
    function. download_ is the path where attachments may be downloaded to.
    log_ is the logger object.

    """
    log_.info(mail_['subject'])
    if "BK_Artikeldaten" in mail_['subject']:
        for part in mail_.walk():
            if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
                download_path = os.path.join(download_, part.get_filename())
                open(download_path, 'wb').write(part.get_payload(decode=True))
                file_path = Path.cwd()
                products_json_path = os.path.join(file_path, "update_stock/products.json")
                stock_update = StagingUpdateProducts(download_path, products_json_path)
                stock_update.process()
        print()
    return 'return meaningful result here'


def main():
    log.info('... script started')
    while True:
        # <--- Start of configuration section

        # Read config file - halt script on failure
        try:
            config_file = open('automation/imap_monitor.ini', 'r+')
        except IOError:
            log.critical('configuration file is missing')
            break
        config = configparser.SafeConfigParser()
        config.readfp(config_file)

        # Retrieve IMAP host - halt script if section 'imap' or value
        # missing
        try:
            host = config.get('imap', 'host')
        except configparser.NoSectionError:
            log.critical('no "imap" section in configuration file')
            break
        except configparser.NoOptionError:
            log.critical('no IMAP host specified in configuration file')
            break

        # Retrieve IMAP username - halt script if missing
        try:
            username = config.get('imap', 'username')
        except configparser.NoOptionError:
            log.critical('no IMAP username specified in configuration file')
            break

        # Retrieve IMAP password - halt script if missing
        try:
            password = config.get('imap', 'password')
        except configparser.NoOptionError:
            log.critical('no IMAP password specified in configuration file')
            break

        # Retrieve IMAP SSL setting - warn if missing, halt if not boolean
        try:
            ssl = config.getboolean('imap', 'ssl')
        except configparser.NoOptionError:
            # Default SSL setting to False if missing
            log.warning('no IMAP SSL setting specified in configuration file')
            ssl = False
        except ValueError:
            log.critical('IMAP SSL setting invalid - not boolean')
            break

        # Retrieve IMAP folder to monitor - warn if missing
        try:
            folder = config.get('imap', 'folder')
        except configparser.NoOptionError:
            # Default folder to monitor to 'INBOX' if missing
            log.warning('no IMAP folder specified in configuration file')
            folder = 'INBOX'

        # Retrieve path for downloads - halt if section of value missing
        try:
            download = config.get('path', 'download')
        except configparser.NoSectionError:
            log.critical('no "path" section in configuration')
            break
        except configparser.NoOptionError:
            # If value is None or specified path not existing, warn and default
            # to script path
            log.warn('no download path specified in configuration')
            download = None
        finally:
            download = download if (
                    download and path.exists(download)
            ) else path.abspath(__file__)
        log.info('setting path for email downloads - {0}'.format(download))

        while True:
            # <--- Start of IMAP server connection loop

            # Attempt connection to IMAP server
            log.info('connecting to IMAP server - {0}'.format(host))
            try:
                imap = imapclient.IMAPClient(host, use_uid=True, ssl=ssl)
            except Exception:
                # If connection attempt to IMAP server fails, retry
                etype, evalue = sys.exc_info()[:2]
                estr = traceback.format_exception_only(etype, evalue)
                logstr = 'failed to connect to IMAP server - '
                for each in estr:
                    logstr += '{0}; '.format(each.strip('\n'))
                log.error(logstr)
                sleep(10)
                continue
            log.info('server connection established')

            # Attempt login to IMAP server
            log.info('logging in to IMAP server - {0}'.format(username))
            try:
                result = imap.login(username, password)
                log.info('login successful - {0}'.format(result))
            except Exception:
                # Halt script when login fails
                etype, evalue = sys.exc_info()[:2]
                estr = traceback.format_exception_only(etype, evalue)
                logstr = 'failed to login to IMAP server - '
                for each in estr:
                    logstr += '{0}; '.format(each.strip('\n'))
                log.critical(logstr)
                break

            # Select IMAP folder to monitor
            log.info('selecting IMAP folder - {0}'.format(folder))
            try:
                result = imap.select_folder(folder)
                log.info('folder selected')
            except Exception:
                # Halt script when folder selection fails
                etype, evalue = sys.exc_info()[:2]
                estr = traceback.format_exception_only(etype, evalue)
                logstr = 'failed to select IMAP folder - '
                for each in estr:
                    logstr += '{0}; '.format(each.strip('\n'))
                log.critical(logstr)
                break

            # Retrieve and process all unread messages. Should errors occur due
            # to loss of connection, attempt restablishing connection
            try:
                result = imap.search('UNSEEN')
            except Exception:
                continue
            log.info('{0} unread messages seen - {1}'.format(
                len(result), result
            ))
            for each in result:
                try:
                    result = imap.fetch(each, ['RFC822'])
                except Exception:
                    log.error('failed to fetch email - {0}'.format(each))
                    continue
                mail = email.message_from_bytes(result[each][b'RFC822'])
                try:
                    process_email(mail, download, log)
                    log.info('processing email {0} - {1}'.format(
                        each, mail['subject']
                    ))
                except Exception:
                    log.error('failed to process email {0}'.format(each))
                    raise
                    continue

            while True:
                # <--- Start of mail monitoring loop

                # After all unread emails are cleared on initial login, start
                # monitoring the folder for new email arrivals and process
                # accordingly. Use the IDLE check combined with occassional NOOP
                # to refresh. Should errors occur in this loop (due to loss of
                # connection), return control to IMAP server connection loop to
                # attempt restablishing connection instead of halting script.
                imap.idle()
                # TODO: Remove hard-coded IDLE timeout; place in config file
                result = imap.idle_check(5 * 60)
                if result:
                    imap.idle_done()
                    result = imap.search('UNSEEN')
                    log.info('{0} new unread messages - {1}'.format(
                        len(result), result
                    ))
                    for each in result:
                        fetch = imap.fetch(each, ['RFC822'])
                        mail = email.message_from_bytes(
                            fetch[each][b'RFC822']
                        )
                        try:
                            process_email(mail, download, log)
                            log.info('processing email {0} - {1}'.format(
                                each, mail['subject']
                            ))
                        except Exception:
                            log.error(
                                'failed to process email {0}'.format(each))
                            raise
                            continue
                else:
                    imap.idle_done()
                    imap.noop()
                    log.info('no new messages seen')
                # End of mail monitoring loop --->
                continue

            # End of IMAP server connection loop --->
            break

        # End of configuration section --->
        break
    log.info('script stopped ...')


if __name__ == '__main__':
    main()