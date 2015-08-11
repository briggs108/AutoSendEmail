# AutoSendEmail
Automate Email Sending in Excel and Outlook

Author: Ted Briggs

CHANGELOG: N/A
License: N/A
Bugs: None known. Please report problems

Installation: The entire program is contained within the Excel environment. The code is located in the VBA editor within the Developer tab.

Purpose of program: The workbook is used to store employee biographical information. If the information needs to be updated, the program will automatically generate personalized emails populated with each person's biographical information, along with an HTML printed message. The program can also automatically send these emails. 

Instructions: First, fill out relevant biographical information in the appropriate columns. Second, select a row from which you would like to start generating emails.
          Third, execute the Generate_Email macro via the Developer tab. Because the "Send" command is commented out, the macro will automatically generate all of the emails with a brief message and the biographical information in tables.
          You can execute the code on either tab, and the message is slightly different based on which tab you start from. 
          
NOTES: This program is only compatible with Microsoft Excel and Outlook. The "Send" command within the code has been commented out so as not to run. Please uncomment if you would like the emails to automatically send. 
