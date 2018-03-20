from win32com.client.gencache import EnsureDispatch as Dispatch
import os
import codecs

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

class Oli():
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in range(1,array_size+1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted( self._obj._prop_map_get_.keys() )

# for inx, folder in Oli(mapi.Folders).items():
#     # iterate all Outlook folders (top level)
#     #print "-"*70
#     print(folder.Name)

#     for inx,subfolder in Oli(folder.Folders).items():
#         print("({}) {} => {}".format(inx, subfolder.Name, subfolder))

def mkdir_p(path):
    try:
        os.makedirs(path)
    except:
        pass

def saveEmails(folder, mailPath):
    for inx, mail in Oli(folder.Items).items():
        if mail.Attachments.Count == 0:
            continue

        try:
            name = "{} {} - {}".format(mail.CreationTime.Format("%Y-%m-%d"), mail.Sender.Name, mail.Sender.Address)
            if os.path.exists(mailPath + '/' + name):
                # print("Skip {}".format(name))
                continue

            mkdir_p(mailPath + '/' + name)

            print(name)
            # print(mail.Body)

            path = mailPath + '/' + name

            text_file = codecs.open(path + "/body.txt", "w", "utf-8")
            text_file.write("From: {}\n".format(mail.Sender.Address))
            text_file.write("Subject: {}\n\n".format(mail.Subject))
            text_file.write(mail.Body)
            text_file.close()

            for inx, att in Oli(mail.Attachments).items():
                print("{}: {}".format(name, att.FileName))
                att.SaveAsFile("{}\\{}\\{}".format(mailPath, name, att.FileName))
        except:
            pass
            # break

internResumes = mapi.Folders[2].Folders[2].Folders[1].Folders[1]
resumes = mapi.Folders[2].Folders[2].Folders[1].Folders[2]

saveEmails(internResumes, os.environ['HOME'] + "/NextCloud/ResumeUploads")
saveEmails(internResumes, os.environ['HOME'] + "/NextCloud/InternResumeUploads")

# saveEmails(internResumes, os.getcwd() + "/Nextcloud/ResumeUploads")
# saveEmails(internResumes, os.getcwd() + "/Nextcloud/InternResumeUploads")