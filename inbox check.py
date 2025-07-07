import win32com.client as win32
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
for store in outlook.Folders:
    print(store.Name)