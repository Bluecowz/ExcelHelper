Hello!
* This program handles spreadsheet manipulation for work. Because it is a pain in the ass.
* Author: Mike Gallant
* Version: 1.2

Handles my excel spreadsheets for work. 

At work we use Excel spreadsheets for work so I designed this to handle the excel manipulation for me. 

Can add times, descriptions of what I have done, and can send it in an email to whomever needs it.

You will need to add 3 directories to the config file. The path of this program, your timesheet archive folder, and the
template folder to create the timesheets.

TODO:
1)Find a better way to store the password. I understand it is an issue storing it in plain text.
2) Send email in another thread to the UI is responsive.
3) Overall efficiencies are needed.

Dependencies:
Openpyxl version 2.4.1

Config file setup:

config.xml

<Settings>
    <MyPath>Path to the current excel spreadsheet</MyPath>
    <ArchivePath>Path to archive</ArchivePath>
    <TemplatePath>Path to archive</TemplatePath>
    <TemplateName>Template name</TemplateName>
    <ActiveXL>The active name</ActiveXL>
    <YourEmail>Your email</YourEmail>
    <Name>Your name</Name>
    <Recipient>Where is it going?</Recipient>
    <STMPServer>STMP server</STMPServer>
    <SMTPPort>Port</SMTPPort>
    <Password>Email Password</Password>
</Settings>


Feedback always welcome!
