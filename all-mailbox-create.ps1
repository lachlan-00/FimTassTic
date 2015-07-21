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
. "D:\Program Files\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1"; Connect-ExchangeServer -auto
$input = Import-CSV  .\csv\fim_staffALL.csv -Encoding UTF8
$StudentInput = Import-CSV .\csv\fim_student_filtered.csv -Encoding UTF8
$userdomain = "userdomain"

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
$FULLDATE = $DATE, "00:00:00"

write-host "### Parsing Staff File"
write-host

#############################################
### Create / Disable Staff Email Accounts ###
#############################################

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
        # Create mailbox for active users
        If ($TestUser) {
            # Enable mailbox for user If mail address is missing
            #if ((!($TestUser.mail)) -and (!($TestUser.Description -eq "Relief Teacher"))) {
            if (!($TestUser.mail)) {
                #Enable-Mailbox -Identity $TestUser.name -Database Staff -AddressBookPolicy "Staff Address Policy"
                Enable-Mailbox -Identity "${userdomain}\${LoginName}" -Alias "${LoginName}"  -Database All-Staff -AddressBookPolicy "Staff Address Policy"
                Set-Mailbox -Identity "${userdomain}\${LoginName}" -RecipientLimits 50
                write-host $LoginName "created mailbox"
                write-host
            }
        }
    }

    ### Process Terminated Users ###
    
    Else {
        # keep some users
        If ($TestDescription -eq "keep") {
            write-host "${LoginName} Keeping terminated user"
            write-host
        }
        # Terminate Staff AFTER their Termination date
        ElseIf ($FULLDATE -gt $Termination) {
            If ($TestUser) {
                # Disable mailbox for users that have a mail address
                if ($TestUser.mail) {
                    #Disable-Mailbox -Identity $TestUser.name -Confirm:$false
                    Disable-Mailbox -Identity "${userdomain}\${LoginName}" -Confirm:$false
                    write-host $LoginName "Disabled mailbox"
                }
            }
        }
        # Wait for the Termination date
        ElseIf ((!($FULLDATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
            write-host "Not Final Leaving Date for ${LoginName}"
            write-host $DATE
            write-host $Termination
            write-host
        }
    }
}

write-host "### Parsing Student File"
write-host

###############################################
### Create / Disable Student Email Accounts ###
###############################################

foreach($line in $StudentInput) {

    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())

    # Check for students who have left
    $Termination = $line.dol.Trim()

    # Set user to confirm details
    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
    $TestEnabled = $TestUser.Enabled
    $TestDescription = $TestUser.Description

    ### Process Current Users ###
    If ($Termination.length -eq 0) {
        If ($TestUser) {
            # Enable mailbox for user If mail address is missing
            if (!($TestUser.mail)) {
                #Enable-Mailbox -Identity $TestUser.name -Database Student -AddressBookPolicy "Student Address Policy"
                Enable-Mailbox -Identity "${userdomain}\${LoginName}" -Alias "${LoginName}" -Database All-Student -AddressBookPolicy "Student Address Policy"
                Set-Mailbox -Identity "${userdomain}\${LoginName}" -RecipientLimits 5
                write-host $LoginName "created mailbox"
                write-host
            }
        }
    }
    ### Process Terminated Users ###
    Else {
        # Keep some users Open
        If ($TestDescription -eq "keep") {
            write-host "${LoginName} Keeping terminated user"
            write-host
        }
        ElseIf ($DATE -gt $Termination) {
            If ($TestUser) {
                # Disable mailbox for users that have a mail address
                if ($TestUser.mail) {
                    Disable-Mailbox -Identity "${userdomain}\${LoginName}" -Confirm:$false
                    write-host $LoginName "Disabled mailbox"
                }
            }
        }
        # Wait for the Termination date
        ElseIf ((!($DATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
            write-host "Not Final Leaving Date for ${LoginName}"
            write-host $DATE
            write-host $Termination
            write-host
        }
    }
}

write-host "### Mailbox Creation Script Finished"
write-host

