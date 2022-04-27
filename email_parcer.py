
# Code to take in a list of names and email addresses in csv format and spit 
# out the email addresses separated by semicolons so it can be use in microsoft
# Outlook for group mailings.
import pandas        #  For reading and parcing the csv file
import pathlib       #  For finding the silly file on Windows

#file_to_open = pathlib.Path(r'C:\Temp\WomenInPhysics.csv')
file_to_open = pathlib.Path(r'C:\Users\rtlines\OneDrive - BYU-Idaho\BYUI-Synced\Course Materials\Department Chair\Women in Physics\WomenInPhysics.csv')

df = pandas.read_csv(file_to_open)   # read the file into a data frame
# Print out the data frame just to be sure we got the right data
print(df)
print("\n")
# The data is structured as a column of names and a second column of emails
cnt = df.shape[0]      # get the number of rows of email addresses
print(cnt)             # and print it for good measure
print("\n")
temp = ""               # this is a place to put the MS Outlook email set
for i in range (cnt):
   temp = temp + df.Email[i] + "; "  # get each email and add the semicolon
print (temp)                          # print it out for copy and past into outlook

# And that is all for now
