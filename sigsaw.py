# a script borne of spite and frustration. A specific prototype of a generalizable script
import win32com.client
import os
# from datetime import datetime, timedelta

# accessing outlook
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# get messages from drafts folder   
drafts = mapi.GetDefaultFolder(16)
messages = drafts.Items

for message in messages:
    pattied = message.Body

# defining start and end of extraneous signature guck. Ultimately prompt for this

start_bull = "Just published!"
end_bull = "Omaha)"

# getting index of start and end point
start_point = pattied.find(start_bull)
end_point = pattied.find(end_bull) + len(end_bull)

# create substring consisting of the crap to be excised
fullbull = pattied[start_point:end_point]

# excising the crap
depattied = pattied.replace(fullbull, " ")

print(depattied)
