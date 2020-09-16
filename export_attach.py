import win32com.client

# connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# main folder and deeper
inbox = outlook.Folders["evgeniy@cg.ru"].Folders["Inbox"].Folders["For export"]
messages = inbox.Items
i = 0

# loop by message in folder
for message in messages:
    for att in message.Attachments:
        att.SaveAsFile('c://work//' + str(i) + '_' + att.FileName)
    i += 1
