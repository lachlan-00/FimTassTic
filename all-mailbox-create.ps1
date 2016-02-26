###############################################################################
###                                                                         ###
###  Create Staff and Student Email Accounts From TASS.web Data             ###
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

import-module activedirectory

If (Get-Command Get-Mailbox) {
    #write-host "Exchange is imported"
}
Else {
    #write-host "importing exchange module"
    . 'D:\Program Files\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto
}

# Input CSV's
$input = Import-CSV  "C:\DATA\csv\fim_staffALL.csv" -Encoding UTF8
$StudentInput = Import-CSV "C:\DATA\csv\fim_student_filtered.csv" -Encoding UTF8
$enrolledinput = Import-CSV "C:\DATA\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8

$userdomain = "VILLANOVA"

write-host "### Starting Mailbox Creation Script"
write-host

# Set the date to match the termination date for each type of user
$YEAR = [string](Get-Date).Year
$MONTH = [string](Get-Date).Month
If ($MONTH.length -eq 1) {
    $MONTH = "0${MONTH}"
}
$DAY = [string](Get-Date).Day
If ($DAY.length -eq 1) {
    $DAY = "0${DAY}"
}
# Student date format
$DATE = "${YEAR}/${MONTH}/${DAY}"
# Staff date format
$FULLDATE = "${DATE} 00:00:00"
$LogDate = "${YEAR}-${MONTH}-${DAY}"

write-host "### Parsing Staff File"
write-host

#############################################
### Create / Disable Staff Email Accounts ###
#############################################

# check log path
If (!(Test-Path "C:\DATA\log")) {
    mkdir "C:\DATA\log"
}

# set log file
$LogFile = "C:\DATA\log\Email-Creation-${LogDate}.log"
$LogContents = @()

foreach($line in $input) {

    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())

    # Check for staff who have left
    $Termination = $line.term_date.Trim()

    # Set user to confirm details
    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
    $TestEnabled = $TestUser.Enabled
    $TestDescription = $TestUser.Description

    ### Process Current Users ###
    If ($Termination.length -eq 0) {
        If ($TestUser) {
            # Enable mailbox for user If mail address is missing
            If (!($TestUser.mail)) {
                Enable-Mailbox -Identity "${userdomain}\${LoginName}" -Alias "${LoginName}"  -Database All-Staff -AddressBookPolicy "Staff Address Policy"
                Set-Mailbox -Identity "${userdomain}\${LoginName}" -RecipientLimits 50
                $LogContents += "Created mailbox for ${LoginName}" #| Out-File $LogFile -Append
            }
            #Enable Archive mailbox for user if missing
            If (!($TestUser.msExchArchiveDatabaseLink)) {
                Enable-Mailbox -Identity "${userdomain}\${LoginName}" -Archive  -ArchiveDatabase Archive-Staff
                $LogContents += "Created archive mailbox for ${LoginName}" #| Out-File $LogFile -Append

            }

        }
    }

    ### Process Terminated Users ###
    
    Else {
        # keep some users
        If ($TestDescription -eq "keep") {
            If ((!($LoginName -eq 'morrt')) -or (!($LoginName -eq 'latei')) -or (!($LoginName = 'obrij'))) {
                $LogContents += "${LoginName} Keeping terminated user" #| Out-File $LogFile -Append
            }
        }
        # Terminate Staff AFTER their Termination date
        ElseIf ($DATE -gt $Termination) {
            If ($TestUser) {
                # Disable mailbox for users that have a mail address
                If ($TestUser.mail) {
                    #Disable-Mailbox -Identity $TestUser.name -Confirm:$false
                    Disable-Mailbox -Identity "${userdomain}\${LoginName}" -Confirm:$false
                    $LogContents += "Disabled mailbox for ${LoginName}" #| Out-File $LogFile -Append
                }
            }
        }
        # Wait for the Termination date
        #ElseIf ((!($FULLDATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
        #    $LogContents += "Not Final Leaving Date for ${LoginName}" #| Out-File $LogFile -Append
        #    $LogContents += "Now: ${DATE}" #| Out-File $LogFile -Append
        #    $LogContents += "DOL: ${Termination}" #| Out-File $LogFile -Append
        #}
    }
}

write-host "### Parsing Student File"
write-host

###############################################
### Create / Disable Student Email Accounts ###
###############################################

#foreach($line in $StudentInput) {

#    # Get login name for processing
#    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())

#    # Check for students who have left
#    $Termination = $line.dol.Trim()

#    # Set user to confirm details
#    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
#    $TestEnabled = $TestUser.Enabled
#    $TestDescription = $TestUser.Description

#    ### Process Current Users ###
#    If ($Termination.length -eq 0) {
#        If ($TestUser) {
#            # Enable mailbox for user If mail address is missing
#            If (!($TestUser.mail)) {
#                Enable-Mailbox -Identity "${userdomain}\${LoginName}" -Alias "${LoginName}" -Database All-Student -AddressBookPolicy "Student Address Policy"
#                Set-Mailbox -Identity "${userdomain}\${LoginName}" -RecipientLimits 5
#                $LogContents += "Created mailbox for ${LoginName}" #| Out-File $LogFile -Append
#            }
#        }
#    }
#    ### Process Terminated Users ###
#    Else {
#        # keep some users
#        If ($TestDescription -eq "keep") {
#            If (!($LoginName -eq '10961')) {
#                $LogContents += "${LoginName} Keeping terminated user" #| Out-File $LogFile -Append
#            }
#        }
#        # Terminate Students AFTER their Termination date
#        ElseIf ($DATE -gt $Termination) {
#            If ($TestUser) {
#                # Disable mailbox for users that have a mail address
#                If ($TestUser.mail) {
#                    Disable-Mailbox -Identity "${userdomain}\${LoginName}" -Confirm:$false
#                    $LogContents += "Disabled mailbox for ${LoginName}" #| Out-File $LogFile -Append
#                }
#            }
#        }
#        # Wait for the Termination date
#        #ElseIf ((!($DATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
#        #    $LogContents += "Not Final Leaving Date for ${LoginName}" #| Out-File $LogFile -Append
#        #    $LogContents += "Now: ${DATE}" #| Out-File $LogFile -Append
#        #    $LogContents += "DOL: ${Termination}" #| Out-File $LogFile -Append
#        #}
#    }
#}

#write-host "### Parsing Future Student File"
#write-host

############################################
### Create FUTURE Student Email Accounts ###
############################################

#foreach($line in $enrolledinput) {

#    # Get login name for processing
#    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())

#    # Set user to confirm details
#    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
#    $TestEnabled = $TestUser.Enabled
#    $TestMail = $TestUser.mail

#    ### Process Current Users ###
#    If (($TestUser) -and ($TestEnabled)) {
#        # Enable mailbox for user If mail address is missing
#        If (!($TestMail)) {
#            Enable-Mailbox -Identity "${userdomain}\${LoginName}" -Alias "${LoginName}" -Database All-Student -AddressBookPolicy "Student Address Policy"
#            Set-Mailbox -Identity "${userdomain}\${LoginName}" -RecipientLimits 5
#            $LogContents += "Created mailbox for ${LoginName}" #| Out-File $LogFile -Append
#        }
#    }
#}

# Write log if changes have occurred
If ($LogContents.Count -gt 0) {
    Write-Host "Writing changes to log file"
    Write-output "" | Out-File $LogFile -Append
    Get-Date | Out-File $LogFile -Append
    foreach($line in $LogContents) {
        Write-Output $line | Out-File $LogFile -Append
    }
}
Else {
    Write-Host "No Changes occurred"
}

write-host
write-host "### Mailbox Creation Script Finished"
write-host
