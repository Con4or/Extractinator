import asyncio
import win32com.client 
import re
import psutil
import asyncio

# CHANGELOG
    # v1.11
        # Clarified error throwing if outlook isn't open, line 64-65
    # v1.1
        # Imported asyncio to enable waiting/sleeping
        # Added async def main(), should make program wait 3 seconds so you can enjoy the viewing pleasure of extractinator ascii art
        # Called above function in the def extractinator_splash function on line 52
        # Adjusted the splash art to give an extra space above to more clearly see it
    # v1.0
        # Initial commit version. Functionality to check emails and mark as read.
# END CHANGELOG

# Function to wait for 3 seconds
async def main():
    await asyncio.sleep(3)
    
# Function to check if outlook is open
def check_outlook_open():
    list_process = []
    for pid in psutil.pids():
        p = psutil.Process(pid)
        # Append to the list of process
        list_process.append(p.name())
    # If outlook open then return True
    if 'OUTLOOK.EXE' in list_process:
        return True
    else:
        return False

# Function to search for IP addresses in the email bodies
def get_IPaddy(Item):
    try:
        body = Item.body
        matches = re.findall(r"\b(?:[0-9]{1,3}\[\.\]){3}[0-9]{1,3}\b", body, re.MULTILINE)
        matchy = []
        for match in matches:
            matchy.append(match)
    except:
        print("None")
    return matchy

# Because all hackers need the proper ASCII splash art for their programs
def extractinator_splash():
    print("\n")
    print("      __  __     v1.11")
    print("   ___\ \/ /____ _   __   ____ ____ _ _  __ __  _____ ____ __")
    print("  / ___\  /_  _/  \ /  | / __/_  _/_// |/ //  |/_  _/ _  //  \\")
    print(" / ___//  \/ // _ // _ |/ /_  / // // || // _ | / // // // _ /")
    print("/____//_/\/_//_/\_\_/|_|___/ /_//_//_/|_//_/|_|/_//____//_/\_\\")
    print("\n")

    asyncio.run(main())

try:
    outlook_open = check_outlook_open()
except:
    outlook_open = False

if outlook_open == False:
    print("Outlook is not open. Please start Outlook to Extractinate!")
# If outlook is open, then execute once before finishing
if outlook_open == True:
    outlook = win32com.client.Dispatch("outlook.Application").GetNameSpace("MAPI")
       
    #2 == mailbox, need to then reference each folder within to drill down to destination
    ########################################################################################################################
    # change the values in the parentheses in Line 77 to navigate through folders and then subfolders.                     #
    # If you don't have subfolders within Inbox, or just a single subfolder, get rid of the unneeded .Folders("CHANGE-ME") #
    ########################################################################################################################
    try:
        inbox = outlook.Folders.Item(2).Folders("Inbox").Folders("CHANGE-ME").Folders("CHANGE-ME")
        messages = inbox.Items
        dedupe_msg = []
        for message in messages:
            if message.UnRead == True:
                try:
                    dedupe_msg = dedupe_msg + get_IPaddy(message)
                    message.UnRead = False
                except:
                    continue
        dedupe_msg = list(dict.fromkeys(dedupe_msg))
        extractinator_splash()
        for a in dedupe_msg:
            # re.sub doesn't "see" the brackets without the backslash
            b = re.sub("\[.\]", ".", a)
            print(b)
        # Program gets to the end of unread emails message
        print("\n")
        print("All emails have been.... Extractinated")
                        
    except:
        print("Error Occurred. Could not Extractinate Subfolder")