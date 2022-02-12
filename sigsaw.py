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

# write pattied and depattied to text documents

f = open(r"C:\Users\edawson4\Documents\emails\depattied.txt","w")
f.write(depattied)
f.close

g = open(r"C:\Users\edawson4\Documents\emails\pattied.txt","w")
g.write(pattied)
g.close

# produce statistics

words = pattied.split(' ')
total_words = len(words)
real_words = depattied.split(' ')
total_real_words = len(real_words)
cut_words = total_words - total_real_words
percent_bs = round(((cut_words / total_words) * 100), 2)
print(percent_bs, "percent of words in this email chain were unnecessary!\n Of ", total_words, "total words,", cut_words, "were self-aggrandazing email signature bs.\n A version of the exchange with these words removed has been saved to local storage.")


# future change--prompt for file names, etc
