###############################################################################
###                                                                         ###
###  Email script logs to nominated account                                 ###
###                                                                         ###
###  ----------------Authors----------------                                ###
###  Lachlan de Waard <lachlan.00@gmail.com>                                ###
###  ----------------Licence----------------                                ###
###  GNU General Public License version 3                                   ###
###                                                                         ###
###  This program is free software: you can redistribute it and/or modify   ###
###  it under the terms of the GNU General Public License as published by   ###
###  the Free Software Foundation, either version 3 of the License, or      ###
###  (at your option) any later version.                                    ###
###                                                                         ###
###  This program is distributed in the hope that it will be useful,        ###
###  but WITHOUT ANY WARRANTY; without even the implied warranty of         ###
###  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          ###
###  GNU General Public License for more details.                           ###
###                                                                         ###
###  You should have received a copy of the GNU General Public License      ###
###  along with this program.  If not, see <http://www.gnu.org/licenses/>.  ###
###                                                                         ###
###############################################################################

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
$filestaff = "C:\DATA\log\staff-${LogDate}.log"
$filestudent = "C:\DATA\log\student-${LogDate}.log"
$filemailbox = "C:\DATA\log\Email-Creation-${LogDate}.log"

#EMAIL SETTINGS
# specify who gets notified
$tonotification = "it@vnc.qld.edu.au"
#$bccnotification = ""
# specify where the notifications come from
$fromnotification = "it@vnc.qld.edu.au"
# specify the SMTP server
$smtpserver = "mail.vnc.qld.edu.au"
# message subject
$emailsubject = "User Creation Logs: ${Now}"
$emailbody = "This is an automated email containing the logs for user creation."
$outsubject = $null

# Attachs files if they exist
$filearray = @()
if (Test-Path $filestaff) {
    $filearray += $filestaff
    $outsubject = "AD ${emailsubject}"
}
if (Test-Path $filestudent) {
    $filearray += $filestudent
    $outsubject = "AD ${emailsubject}"
}
if (Test-Path $filemailbox) {
    $filearray += $filestudent
    $outsubject = "EMAIL ${emailsubject}"
}

#set specific title based on files
if ($outsubject) {
    $emailsubject = $outsubject
}

# Only send mail is the files exist
if($filearray.Count -gt 0) {
    Send-MailMessage -From $fromnotification -Subject $emailsubject -To $tonotification -Attachments $filearray -Body $emailbody -SmtpServer $smtpserver
}
