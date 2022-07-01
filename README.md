# Extractinator
Fun little Python script to open a folder in Outlook and search for IP addresses, and then print them to the terminal.

To use, go to line 76 in the script and change .Folder("CHANGE-ME") to match your inbox-subfolder structure. If you have no subfolders, get rid of them and just keep .Folder("Inbox"). If you only have one subfolder, get rid of any extra .Folder("CHANGE-ME")'s. I have a subfolder within a subfolder. If you have more, add successive ones to it to get to the subfolder you are looking to extract.
