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
. 'E:\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto
$input = Import-CSV  .\csv\telemf.csv
$StudentInput = Import-CSV .\csv\student.csv

write-host "### Starting Mailbox Creation Script"
write-host
write-host "### Parsing Staff File"
write-host

#############################################
### Create / Disable Staff Email Accounts ###
#############################################

foreach($line in $input)
{
    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())

    # Check for staff who have left
    $Termination = $line.term_date.Trim()
    
    ### Process Current Users ###
    
    If ($Termination.length -eq 0) {
        If (Get-ADUser -Filter { (SamAccountName -eq $LoginName) }) {
            # Set user to confirm details
            #$TestUser = Get-ADUser -Identity $LoginName -property "mail"
            $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

            # Enable mailbox for user If mail address is missing
            if ((!($TestUser.mail)) -and (!($TestUser.Description -eq "Relief Teacher"))) {
                Enable-Mailbox -Identity $TestUser.name -Database Staff
                write-host $LoginName "created mailbox"
            }
            # Disable accounts for relief teachers
            #if (($TestUser.mail) -and ($TestUser.Description -eq "Relief Teacher")) {
            #    Disable-Mailbox -Identity $TestUser.name -Confirm:$false
            #    write-host $LoginName "Disabled mailbox"
            #}
        }
    }

    ### Process Terminated Users ###
    
    Else {
        $YEAR = [string](Get-Date).Year
        $MONTH = [string](Get-Date).Month
        If ($MONTH.length -eq 1) {
            $MONTH = "0${MONTH}"
        }
        $DAY = [string](Get-Date).Day
        If ($DAY.length -eq 1) {
            $DAY = "0${DAY}"
        }

        $DATE = "${YEAR}-${MONTH}-${DAY}"
        $DATE = $DATE, "00:00:00"
        If (($DATE -gt $Termination) -and (!($Termination.length -eq 0))) {
            If (Get-ADUser -Filter { (SamAccountName -eq $LoginName) }) {
                # Set user to confirm details
                #$TestUser = Get-ADUser -Identity $LoginName -property "mail"
                $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

                # Disable mailbox for users that have a mail address
                if ($TestUser.mail) {
                    Disable-Mailbox -Identity $TestUser.name -Confirm:$false
                    write-host $LoginName "Disabled mailbox"
                }
            }
        }
        ElseIf ((!($DATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
            write-host "Not Final Leaving Date", $FullName
            write-host $DATE
            write-host $Termination
        }
    }
}

write-host "### Parsing Student File"
write-host

###############################################
### Create / Disable Student Email Accounts ###
###############################################

foreach($line in $StudentInput)
{
    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())

    # Check for students who have left
    $Termination = $line.dol.Trim()
    
    ### Process Current Users ###
    
    If ($Termination.length -eq 0) {
        If (Get-ADUser -Filter { (SamAccountName -eq $LoginName) }) {
            # Set user to confirm details
            $TestUser = Get-ADUser -Identity $LoginName -property "mail"

            # Enable mailbox for user If mail address is missing
            if (!($TestUser.mail)) {
                Enable-Mailbox -Identity $TestUser.name -Database Student
                write-host $LoginName "created mailbox"
                write-host
            }
        }
    }

    ### Process Terminated Users ###
    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

    If ($TestUser.Description -eq "keep") {
                    write-host "${LoginName} Keeping terminated user"
    }
    Else {
        $YEAR = [string](Get-Date).Year
        $MONTH = [string](Get-Date).Month
        If ($MONTH.length -eq 1) {
            $MONTH = "0${MONTH}"
        }
        $DAY = [string](Get-Date).Day
        If ($DAY.length -eq 1) {
            $DAY = "0${DAY}"
        }
        $DATE = "${YEAR}-${MONTH}-${DAY}"
        $DATE = $DATE, "00:00:00"
        If (($DATE -gt $Termination) -and (!($Termination.length -eq 0))) {
            If (Get-ADUser -Filter { (SamAccountName -eq $LoginName) }) {
                # Set user to confirm details
                $TestUser = Get-ADUser -Identity $LoginName -property "mail"

                # Disable mailbox for users that have a mail address
                if ($TestUser.mail) {
                    write-host $DATE
                    write-host $Termination
                    write-host
                    Disable-Mailbox -Identity $TestUser.name -Confirm:$false
                    write-host $LoginName "Disabled mailbox"
                }
            }
        }
        ElseIf ((!($DATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
            write-host "Not Final Leaving Date", $FullName
            write-host $DATE
            write-host $Termination
        }
    }
}

write-host "### Mailbox Creation Script Finished"
write-host