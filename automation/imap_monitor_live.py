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
from update_stock.update_products_data_live import LiveUpdateProducts

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

class EmailMonitor:
    def __init__(self, path_ini_file):
        while True:
            # <--- Start of configuration section

            # Read config file - halt script on failure
            try:
                self.config_file = open(path_ini_file, 'r+')
            except IOError:
                log.critical('configuration file is missing')
                break
            config = configparser.ConfigParser()
            config.read_file(self.config_file)

            # Retrieve IMAP host - halt script if section 'imap' or value
            # missing
            try:
                self.host = config.get('imap', 'host')
            except configparser.NoSectionError:
                log.critical('no "imap" section in configuration file')
                break
            except configparser.NoOptionError:
                log.critical('no IMAP host specified in configuration file')
                break

            # Retrieve IMAP username - halt script if missing
            try:
                self.username = config.get('imap', 'username')
            except configparser.NoOptionError:
                log.critical('no IMAP username specified in configuration file')
                break

            # Retrieve IMAP password - halt script if missing
            try:
                self.password = config.get('imap', 'password')
            except configparser.NoOptionError:
                log.critical('no IMAP password specified in configuration file')
                break

            # Retrieve IMAP SSL setting - warn if missing, halt if not boolean
            try:
                self.ssl = config.getboolean('imap', 'ssl')
            except configparser.NoOptionError:
                # Default SSL setting to False if missing
                log.warning('no IMAP SSL setting specified in configuration file')
                self.ssl = False
            except ValueError:
                log.critical('IMAP SSL setting invalid - not boolean')
                break

            # Retrieve IMAP folder to monitor - warn if missing
            try:
                self.folder = config.get('imap', 'folder')
            except configparser.NoOptionError:
                # Default folder to monitor to 'INBOX' if missing
                log.warning('no IMAP folder specified in configuration file')
                folder = 'INBOX'

            # Retrieve path for downloads - halt if section of value missing
            try:
                self.download = config.get('path', 'download')
            except configparser.NoSectionError:
                log.critical('no "path" section in configuration')
                break
            except configparser.NoOptionError:
                # If value is None or specified path not existing, warn and default
                # to script path
                log.warn('no download path specified in configuration')
                self.download = None
            finally:
                self.download = self.download if (
                        self.download and path.exists(self.download)
                ) else path.abspath(__file__)
            log.info('setting path for email downloads - {0}'.format(self.download))
            break

    def process_email(self, mail_, download_, log_):
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
            stock_update = LiveUpdateProducts(download_path, products_json_path)
            stock_update.process()
            print()
        return 'return meaningful result here'

    def establish_connection(self):
        log.info('connecting to IMAP server - {0}'.format(self.host))
        while True:
            try:
                self.imap = imapclient.IMAPClient(self.host, use_uid=True, ssl=self.ssl)
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
            break
        log.info('server connection established')

    def login(self):
        # Attempt login to IMAP server
        log.info('logging in to IMAP server - {0}'.format(self.username))
        while True:
            try:
                self.result = self.imap.login(self.username, self.password)
                log.info('login successful - {0}'.format(self.result))
            except Exception:
                # Halt script when login fails
                etype, evalue = sys.exc_info()[:2]
                estr = traceback.format_exception_only(etype, evalue)
                logstr = 'failed to login to IMAP server - '
                for each in estr:
                    logstr += '{0}; '.format(each.strip('\n'))
                log.critical(logstr)
                break
            break

    def select_folder(self):
        while True:
            # Select IMAP folder to monitor
            log.info('selecting IMAP folder - {0}'.format(self.folder))
            try:
                self.result = self.imap.select_folder(self.folder)
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
            break

    def process_unread_emails(self):
        while True:
            # <--- Start of IMAP server connection loop

            # Attempt connection to IMAP server

            # Retrieve and process all unread messages. Should errors occur due
            # to loss of connection, attempt restablishing connection
            try:
                result = self.imap.search('UNSEEN')
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
            break

    def start_monitoring(self):
        while True:
            # <--- Start of mail monitoring loop

            # After all unread emails are cleared on initial login, start
            # monitoring the folder for new email arrivals and process
            # accordingly. Use the IDLE check combined with occassional NOOP
            # to refresh. Should errors occur in this loop (due to loss of
            # connection), return control to IMAP server connection loop to
            # attempt restablishing connection instead of halting script.
            try:
                self.imap.idle()
                # TODO: Remove hard-coded IDLE timeout; place in config file
                result = self.imap.idle_check(5 * 60)
                if result:
                    self.imap.idle_done()
                    result = self.imap.search('UNSEEN')
                    log.info('{0} new unread messages - {1}'.format(
                        len(result), result
                    ))
                    for each in result:
                        fetch = self.imap.fetch(each, ['RFC822'])
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
                    self.imap.idle_done()
                    self.imap.noop()
                    log.info('no new messages seen')
                # End of mail monitoring loop --->
                continue
            except self.imap.AbortError:
                self.establish_connection()
                self.login()
                self.select_folder()
                log.info('Connection Aborted and re-instantiated again')


def main():
    log.info('... script started')
    email_monitoring = EmailMonitor(path_ini_file="automation/imap_monitor.ini")
    email_monitoring.establish_connection()
    email_monitoring.login()
    email_monitoring.select_folder()
    email_monitoring.process_unread_emails()
    email_monitoring.start_monitoring()
    log.info('script stopped ...')


if __name__ == '__main__':
    main()
