@echo off

REM PUTTING THE REPORTS ON THERE PLACE
POWERSHELL .\report_copier.ps1

REM GETTING THE LIST OUT

POWERSHELL .\Email_sent_Not_Using_org.PS1
POWERSHELL .\Email_Sent_OUT.PS1
POWERSHELL .\External_To_org.PS1
POWERSHELL .\Total_Email_Count.PS1
POWERSHELL .\Unauth_Email_Count.PS1
POWERSHELL .\Within_org.PS1
POWERSHELL .\email_recived_from_outside_list.PS1
POWERSHELL .\emails_send_outside_from_org_ID_list.PS1
POWERSHELL .\Emails_within_org_list.PS1
POWERSHELL .\Unauthorized_emails_list.PS1
POWERSHELL .\report_writer.ps1


REM RENAMING THE REPORTS
POWERSHELL .\report_rename.ps1
