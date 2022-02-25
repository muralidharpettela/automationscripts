import win32com.client
import pythoncom
import re
import os
from update_products_data_staging import StagingUpdateProducts


class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            subject = mail.Subject
            #attachments = mail.Attachments
            # save to current directory
            outputDir = 'D:\data'
            if "BK_Artikeldaten" in subject:
                try:
                    for attachment in mail.Attachments:
                        saved_file_location = os.path.join(outputDir, attachment.FileName)
                        attachment.SaveAsFile(saved_file_location)
                        staging_products_update = StagingUpdateProducts(saved_file_location)
                        staging_products_update.process()
                        print(f"attachment saved")
                except Exception as e:
                    print("Error when saving the attachment:" + str(e))
            else:
                print("Subject match didnt found....")
            # Taking all the "BLAHBLAH" which is enclosed by two "%".
            #command = re.search(r"Test_(.*?)", subject).group(1)

            #print(command) # Or whatever code you wish to execute.



outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

#and then an infinit loop that waits from events.
pythoncom.PumpMessages()