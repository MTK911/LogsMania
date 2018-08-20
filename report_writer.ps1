(Get-Date).AddDays(-1).ToString("MM/dd/yyyy") | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 11 -startcolumn 3


$Totalemails = Import-CSV '\Logs Mania\LOGS_DATA\1 Total No. of Email.CSV'
$Totalemails[0].Total_emails | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 18 -startcolumn 3


$EmailsWithinorg = Import-CSV '\Logs Mania\LOGS_DATA\2 Emails Within org.CSV'
$EmailsWithinorg[0].Emails_Within_org | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 19 -startcolumn 3


$orgIDRecivedoutsideemail = Import-CSV '\Logs Mania\LOGS_DATA\3 Emails received from outside.CSV'
$orgIDRecivedoutsideemail[0].orgID_Recived_outside_email | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 20 -startcolumn 3


$outsideemailbyorgID = Import-CSV '\Logs Mania\LOGS_DATA\4 Total emails sent outside.CSV'
$outsideemailbyorgID[0].outside_email_by_orgID | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 21 -startcolumn 3


$otherIDemailfromorg = Import-CSV '\Logs Mania\LOGS_DATA\5 Email Sent Outside By other Id.CSV'
$otherIDemailfromorg[0].otherID_email_from_org | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 22 -startcolumn 3


$orgunauthorizedemail = Import-CSV '\Logs Mania\LOGS_DATA\6 Un-authorized Email sent Outside.CSV'
$orgunauthorizedemail[0].org_unauthorized_email | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Summary' -startrow 23 -startcolumn 3


Import-CSV '\Logs Mania\LOGS_DATA\7 Emails within org.CSV' | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Emails within org'

Import-CSV '\Logs Mania\LOGS_DATA\8 Emails recived from outside.CSV' | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Emails received from outside'


Import-CSV '\Logs Mania\LOGS_DATA\9 Total emails sent outside list.CSV' | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Total emails sent outside'


Import-CSV '\Logs Mania\LOGS_DATA\10 Unauthorized emails.CSV' | export-Excel -Path '\Logs Mania\Reports\Email Log Analysis Report.xlsx' -WorksheetName 'Un-authorized Email'


Import-CSV '\Logs Mania\LOGS_DATA\10 Unauthorized emails.CSV' | export-Excel -Path '\Logs Mania\Reports\Unauth.xlsx' -WorksheetName 'ALL' -Append
