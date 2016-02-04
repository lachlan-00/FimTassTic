#checkservice
#
# check the status of a running service
#

# date
$Now = Get-date

# Get Date and Format Field to Match Log File Date
$YEAR = [string](Get-Date).Year
$MONTH = [string](Get-Date).Month
$DAY = [string](Get-Date).Day
If ($MONTH.length -eq 1) {
    $MONTH = "0${MONTH}"
}
If ($DAY.length -eq 1) {
    $DAY = "0${DAY}"
}
$LogDate = "${YEAR}-${MONTH}-${DAY}"


# Email attachments to add
$filestaff = “.\log\staff-${LogDate}.log”
$filestudent = “.\log\student-${LogDate}.log”

#EMAIL SETTINGS
# specify who gets notified 
$tonotification = "it@vnc.qld.edu.au"
#$bccnotification = "staff-library@vnc.qld.edu.au"
# specify where the notifications come from 
$fromnotification = "notifications@vnc.qld.edu.au"
# specify the SMTP server 
$smtpserver = "mail.vnc.qld.edu.au"
# message subject
$emailsubject = "AD User Creation Logs: ${Now}"
$emailbody = "This is an automated email containing the logs for user creation."

# Attachs files if they exist
$filearray = @()
if (Test-Path $filestaff) {
    $filearray += $filestaff
}
if (Test-Path $filestudent) {
    $filearray += $filestudent
}

# Only send mail is the files exist
if($filearray.Count -gt 0) {
    Send-MailMessage -From $fromnotification -Subject $emailsubject -To $tonotification -Attachments $filearray -Body $emailbody -SmtpServer $smtpserver
}