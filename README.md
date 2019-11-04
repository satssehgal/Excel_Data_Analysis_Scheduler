# Excel_Data_Analysis_Scheduler

This light weight utility will take files sitting in a shared drive, append the data to a file of your choosing and copy the shared file to a repoitory where you can store the shared file for archived reasons.

Folder Structure:

Drop Folder: This is where a shared file will be dropped, ie by finance for monthly results
Processed Folder: This is where the file in the dropped folder will be copied to for archive purposes
Main.xlsx: his is the main file the analyst will work on. When the code is ran, it will take the contents of the file fro the drop    folder, append it to main.xlsx and move the file from Drop to Processed Folder.
myutil.py: Update folder path for your liking, update scheduled time. Let the utility run in the background and it will schedule and automate the process for you.
