
#!/usr/bin/python
# -*- coding: cp1250  -*-
__author__ = 'Fekete András Demeter'

import win32com.client
import time

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

class OutlookLib:

    def __init__(self, settings={}):
        self.settings = settings

    def get_ReceivedTime(self,msg):
        return msg.ReceivedTime

while 1:

    the_earliest_ReceivedTime=None
    try:
        for folder in mapi.Folders:
                try:
                    the_earliest_ReceivedTime=None
                    for subfolder in folder.Folders:
                            try:
                                count = 0
                                for subsubfolder in subfolder.Folders:

                                    for msg in subsubfolder.Items:  # subsubfolder.Items.Restrict("[Unread] = true")
                                            count += 1
                                            if msg.ReceivedTime != None:
                                                if the_earliest_ReceivedTime == None:
                                                    the_earliest_ReceivedTime = msg.ReceivedTime
                                                else:
                                                    if the_earliest_ReceivedTime > msg.ReceivedTime:
                                                        the_earliest_ReceivedTime = msg.ReceivedTime
                                            else:
                                                pass
                                    if count == 0:
                                        the_earliest_ReceivedTimestr = None
                                    else:
                                        the_earliest_ReceivedTimestr = the_earliest_ReceivedTime.strftime("%Y.%m.%d %H:%M:%S")
                                    print("boxName: ", str(folder.Name), " - folderName: ", str(subfolder.Name), " - subfolderName: ", str(subsubfolder.Name), " - mails piece: ", str(count), " - the earliest mail ReceivedDatetime: ",str(the_earliest_ReceivedTimestr))
                                    count = 0
                                    the_earliest_ReceivedTimestr = None
                            except:
                                the_earliest_ReceivedTimestr = None
                                count = 0
                                pass


                            try:
                                count = 0
                                for msg in subfolder.Items:  # subfolder.Items.Restrict("[Unread] = true")
                                    count += 1
                                    try:

                                        if msg.ReceivedTime!=None:

                                            if the_earliest_ReceivedTime==None:
                                                the_earliest_ReceivedTime = msg.ReceivedTime

                                            else:
                                                if the_earliest_ReceivedTime > msg.ReceivedTime:
                                                    the_earliest_ReceivedTime = msg.ReceivedTime

                                                else:
                                                    pass
                                        else:
                                            pass
                                    except:
                                        pass
                                if count == 0:
                                    the_earliest_ReceivedTimestr = None
                                else:
                                    the_earliest_ReceivedTimestr = the_earliest_ReceivedTime.strftime("%Y.%m.%d %H:%M:%S")
                                print("boxName: ", str(folder.Name), " - folderName: ", str(subfolder.Name), " - mails piece: ", str(count), " - the earliest mail ReceivedDatetime: ", str(the_earliest_ReceivedTimestr))
                                count = 0
                                the_earliest_ReceivedTimestr = None
                            except:
                                the_earliest_ReceivedTimestr = None
                                count = 0
                                pass

                except:
                    pass

        time.sleep(900)

    except Exception as e:
        print(e)
        time.sleep(900)
