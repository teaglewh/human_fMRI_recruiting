# human_fMRI_recruiting
Visual Basic code for sending, receiving, and sorting emails to prospective human fMRI study participants.

Presentation: https://docs.google.com/presentation/d/13c9U_ZmzMkuhkrJpSYftrfgHS2jGorm6hfbAp908nwU/edit?usp=sharing

# Contents (code)

**1. After_initial_interest_email**  
    -- Add new email info to Excel sheet; move the email to folder for initial emails  
    -- Respond to email with request for phone screening (currently lists days/times open for calling)  
**2. After_phone_screening**  
     *If passed phone screening:*  
         -- Move the email to the appropriate folder; categorize the email  
         -- Send emails: one email solicits participant information, the other sends current available appointments  
    *If didn't pass phone screening:*  
         -- Move and categorize email accordingly  
**3. After_receive_ATV_info**  
    -- Import filled-out table of participant information to Excel  
    *If selected appointment is available:*  
         -- Respond to the email  
         -- Move the email to appropriate folder  
    *If selected appointment is not available*  
         -- Send an email asking to choose another appointment time  
         -- Send an email with updated appointment availability  
         -- Move the email to appropriate folder  

# How to use this code: Overview

This code creates several macros. You (the user) can then create shortcut buttons in your email which will automatically run the code when pressed.

In my Outlook application, I have customized the ribbon and added a custom group called "Human fMRI" on the main Home (Mail) tab. In this group, I have added four macro commands/buttons each with their own customized symbols: 
     -- After Initial Email (with a phone book symbol)
     -- Pass Phone Screen (green check)
     -- Ineligible After Phone Screen (red X)
     -- Import Participant ATV Info (save symbol)
You can organize your email differently, use different symbols, etc, but this is how I have done it.

# How to use this code: Instructions for use

The potential study participant will email you expressing their interest in the study. Note that you will need to be connected to the VPN/the network to run this code.

1. Click on the email in your inbox to open it in the main window of your Outlook application.
2. Click on the After Initial Email button. If Outlook asks you whether you want to enable macros, click yes.

The email will disappear from your inbox (it has been labeled and moved to a folder called "Human fMRI: Unscreened" for easy organization; if your email does not already have this folder system, a new folder will be created). The participant will receive an automated email instructing them to call for a phone screening. Additionally, details about the email and sender will be saved into a simple Excel spreadsheet which keeps track of who has expressed interest in the study.

3. The participant will call for the phone screening based on the instructions given in the email. You will conduct the screening over the phone based on the following script: https://docs.google.com/document/d/1LSg-PGYq6A_IKoe_DdwDeL_igCmNNbUdN8sEVZCj3jc/edit?usp=sharing
4. Following the phone conversation, go to the "Human fMRI: Unscreened" folder and select/open the interest email from the person you just phone screened.
5. Depending on whether the person passed the phone screening or not, click on the appropriate button ("Pass Phone Screen" or "Ineligible After Phone Screen").

The email will disappear from the "Human fMRI: Unscreened" folder and move to the "Human fMRI: Phone Screened" folder. It will be labeled depending on whether the person passed the phone screening or not. If they did not pass the phone screening, this is the end of what happens with them. However, if they DID pass the phone screening, two emails will automatically be sent to the person: one email will solicit the ATV patient information as well as their preferred appointment time, the other email will send current available appointments (this pulls dates and times from an Excel sheet, which has to be manually updated once in a blue moon). The participant is instructed to copy/paste a table into their return email, then fill out the table with their information before sending the email.

The participant will hopefully understand and follow these instructions. As of writing this, some checks have been built in, but as you use this, you may find that there are frequent use errors that should have checks built for or FAQs that should be addressed in the email. (Due to covid, I have not actually tested this code on participants) You will receive an email in your inbox that has a table of patient information in it.

6. Open the email. Click on the Import Participant ATV Info button.

The table of information will be copied to an Excel sheet for participant ATV information. You will later reference this Excel document when making the participant's ATV account if they don't already have one. The participant's preferred appointment time will then be checked against the Excel file with the list of open appointments. If the appointment is available still, it will be removed from the list and a confirmation email will be sent to the participant. The email will then be moved to a new folder called [whatever it's called] and labeled. If the appointment is not available, a follow up email will be automatically sent to the participant asking them to choose a different day/time. An updated list of available appointments will also be sent. The email will be labeled accordingly.

Now your only responsibilities are to set everything up in the system (create their ATV account if necessary, create an appointment, etc), send out reminder emails, pay the people, etc. See last section and/or Marianne's instructions for more complete information about the other steps.

# Creating the macros and setting up your email

For these instructions, I'm going to assume that you want to do things the way I did. Obviously feel free to adapt these instructions to fit your own preferences.

Note on vocabulary:
  - Ribbon: The thing at the top with File, Home, Send / Receive, Folder, View, Help, etc. Each of these is a different "tab".
  - Group: In each tab, there are different groups of buttons/commands. For example, in the Home tab, there is the New group which includes buttons like "New Email" and "New Items" (at least that's how it is on mine at the moment).

Add Developer tab
  1. Open your email. Right click on the ribbon, revealing some options. Select "Customize the Ribbon..."
  2. Under the Customize the Classic Ribbon, it should have you customizing the Main Tabs. Add the "Developer" tab, then click OK. Nice work!

Create the macros
  1. Click on your new Developer tab. Click on the Visual Basic button in the Code group. (If you don't see this, you might need to further customize your ribbon to get that shortcut) It will open a VBA window.
  2. Right click on Project1 (left side of the screen, if yours looks anything like mine). Select Insert, then select Module.
  3. Copy/paste the code into the module. You can copy it into multipe modules, or all into the same module. Up to you--whatever you think will be most user-friendly for you.
        3.1. Make sure all of the filepaths are correct. If not, change them in the code to reflect how you've mapped the network drive. Alternatively, you could skip this. Then, if any filepaths are not correct when you test the code with some fake emails, it will crash and ask you if you want to debug. Click "yes I want to debug", then it will open up the VBA editor with the troublesome line of code highlighted.
        3.2. Make sure the text in the automated emails reflects what you want it to say (for example, are the days/times available for phone calls up-to-date? Is this how you will be organizing your phone screenings?) As with the filepaths, I recommend checking this when you test the code with some fake emails. For this step, make sure to run through with various scenarios (send yourself an email and have it "pass" the phone screening, send another email that "fails" the phone screening, have another email that passes everything but selects an appointment time that is already taken, etc). This will allow you to avoid combing through and looking for everything in the code itself. Note that currently I believe things are changed to reflect the current covid situation.
  4. Close the VBA window. It saves automatically from what I can tell.

Create your customized shortcut buttons
  1. Go back to the Home tab. Right click on the ribbon and select "Customize the Ribbon..."
  2. Under the "Customize the Classic Ribbon" side, you should see a box with your different tabs in it. Click on "Home (Mail)", then select "New Group" (it's a button below the box, in between "New Tab" and "Rename"). Click on your new group, then click the "Rename" button. Name it "Human fMRI". Choose a symbol that feels right. I chose the phone symbol. Click OK, returning you to the Customize the Ribbon window.
  3. From the "Choose commands from:" drop-down menu, select Macros. (Re-select your new Human fMRI group on the Customize the Classic Ribbon side if it's not highlighted)
  4. You will see a list of macros. Select Project1.MainUno, click Add. Select Project1.MainDos, click Add. Select Project1.IneligibleAfterPhoneScreen, click Add. Select Project1.MainTres, click Add. They should show up within the Human fMRI custom group.
  5. Select MainUno. Click Rename. Name it "After Initial Email" and choose the phone book symbol to represent the fact that it adds the sender's basic information to the Excel sheet. Or choose a different symbol; I don't care.
  6. Select MainDos. Click Rename. Name it "Pass Phone Screen" and choose the green check symbol for obvious reasons.
  7. Select IneligibleAfterPhoneScreen. Click Rename. Name it "Ineligible After Phone Screen" and choose the red X symbol.
  8. Select MainTres. Click Rename. Name it "Import Participant ATV Info" and choose the save/floppy disk symbol to represent the fact that you are saving all the participant's ATV info.
  9. Click OK to close out of the Customize the Ribbon window. In your Home tab, you should see a new group with all your custom buttons in it.
  10. Send yourself an email and test everything, both so you can see how it works and also to make sure these instructions are correct.

# Outstanding TODOs in the code (to get it to do what is described in this file)

TODO: test the code and figure out this list.

# Outstanding TODOs to complete The Full Vision
Formatting notes:  
 - **Bold text** indicates a button on the ribbon that you create, which you click on to complete the step.
 - [Brackets] indicate something that the other person does, like send you an email: the ball is in their court. You don't have to do anything until their action prompts your next step.
 - {Curly brackets} indicate something that you do, such as making appointments in Robert's schedule. Alas, these are tasks which cannot be automated (as far as I know).
 - *Italics* indicate a macro which causes an action to happen automatically. This will run in the background once prompted and will complete the step without your direct or conscious involvement. Hooray for technology!    

Macros and workflow:
- [X] **After initial email: button**
  - [X] Respond with “thanks for interest pls call any time on Tuesday for eligibility screening”  
  - [X] Add new email info to sheet  
  - [X] Move email to folder for initial emails of interest  
  - Note: Doesn’t send automatically for now in case the email is being forwarded from Marianne    
- [X] **After successful eligibility call: button on initial email**
  - [X] Email draws from available screening sheet and lists the next 10  
  - [X] Sends formatted email asking for ATV info and preferred screening times  
  - [X] Moves initial email to folder for completed phone screening ppl  
- [X] [initial email]  
- [X] {call on Tuesday}  
- [X] [ATV info email]  
- [X] *Note: include button for paid in excel sheet for later button*    
- [ ] **After ATV table email received: button**
  - [X] Adds ATV and appt info to new spreadsheet  
  - [ ] Moves ATV email to folder with like emails (regardless of appt success)  
  - [ ] Checks free appts list to confirm that preferred date/time is there.  
  - [ ] If selected date/time is already taken, email back with updated available times and CC me  
  - [ ] If selected date/time is not already taken, delete it from the excel spreadsheet and proceed:  
  - [ ] Confirmation email (“you will get another email in the next few days with more details, but your appt is for this day”, CC me  
  - Note: Later: make ATV table include top 3 choices and algorithm to the highest available slot  
- [ ] *After I input MRN #: automatic*
  - [ ] If “already ATV” not checked when MRN is entered, open but don’t send template email for attaching the long instructions  
  - [ ] If “already ATV” checked when MRN is entered, auto send simple instructions  
  - [ ] Email Stacy and Allison with name date time mrn #, do this encrypted  
- [ ] {create or confirm ATV profile, retrieve MRN # and input it into excel as well as check whether there was already an ATV account}  
- [ ] {create appt in Roberts schedule}  
- [ ] {pay participant or don’t}  
- [ ] *Wednesday 1pm: automatic email*
  - [ ] Encrypted says to pay participants with participant names. Email is to me only    
- [ ] **After pay participant: button 1**
  - [ ] Deletes email  
  - [ ] Checks “paid” button in excel    
- [ ] **After don’t pay participant: button 2**
  - [ ] Deletes email  
  - [ ] Make participant row highlighted red for no show in excel  
- [ ] {possibly get and attach long instructions in email to participant}  
- [ ] *Auto email*
  - [ ] Auto 1 week and 1  day before appt (based on excel sheet), send reminder emails to participant  
- [X] *Auto: if send email w/ subject, send available appts*  

