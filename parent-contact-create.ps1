###############################################################################
###                                                                         ###
###  Create Parent Mail Contact Objects From TASS.web Data                  ###
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
. 'D:\Program Files\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto
$input = Import-CSV  .\csv\fim_parent.csv -Encoding UTF8
$StudentInput = Import-CSV .\csv\fim_student.csv -Encoding UTF8
$ContactOU = "OU=example,DC=qld,DC=edu,DC=au"

write-host "### Starting Contact Creation Script"
write-host
write-host "### Parsing Student File"
write-host


###############################################
### Create / Disable Student Email Accounts ###
###############################################

foreach($line in $StudentInput)
{
    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())
    
    # Get Parent code for each student to search in the parent list
    $TestParCode = (Get-Culture).TextInfo.ToLower($line.par_code.Trim())
    
    # Check for students who have left
    $Termination = $line.dol.Trim()
 
         # Set lower case because powershell ignores uppercase word changes to title case
        If ((Get-Culture).TextInfo.ToUpper($line.preferred_name.Trim()) -eq $line.preferred_name.Trim()) {
            $PreferredName = (Get-Culture).TextInfo.ToLower($line.preferred_name.Trim())
            $PreferredName = (Get-Culture).TextInfo.ToTitleCase($line.preferred_name.Trim())
        }
        Else {
            $PreferredName = ($line.preferred_name.Trim())
        }
        if ((Get-Culture).TextInfo.ToUpper($line.given_name.Trim()) -eq $line.given_name.Trim()) {
            $GivenName = (Get-Culture).TextInfo.ToLower($line.given_name.Trim())
            $GivenName = (Get-Culture).TextInfo.ToTitleCase($line.given_name.Trim())
        }
        Else {
            $GivenName = ($line.given_name.Trim())
        }
        If ((Get-Culture).TextInfo.ToUpper($line.surname.Trim()) -eq $line.surname.Trim()) {
            $Surname = (Get-Culture).TextInfo.ToLower($line.surname.Trim())
            $Surname = (Get-Culture).TextInfo.ToTitleCase($line.surname.Trim())
        }
        Else {
            $Surname = ($line.surname.Trim())
        }

        # Replace Common Acronyms and name spellings
        $Surname = $Surname -replace "D'a", "D'A"
        $Surname = $Surname -replace "De L", "de L"
        $Surname = $Surname -replace "De S", "de S"
        $Surname = $Surname -replace "De W", "de W"
        $Surname = $Surname -replace "Macl", "MacL"
        $Surname = $Surname -replace "Mcb", "McB"
        $Surname = $Surname -replace "Mcc", "McC"
        $Surname = $Surname -replace "Mcd", "McD"
        $Surname = $Surname -replace "Mcg", "McG"
        $Surname = $Surname -replace "Mci", "McI"
        $Surname = $Surname -replace "Mck", "McK"
        $Surname = $Surname -replace "Mcl", "McL"
        $Surname = $Surname -replace "Mcm", "McM"
        $Surname = $Surname -replace "Mcn", "McN"
        $Surname = $Surname -replace "Mcp", "McP"
        $Surname = $Surname -replace "Mcw", "McW"
        $Surname = $Surname -replace "O'c", "O'C"
        $Surname = $Surname -replace "O'd", "O'D"
        $Surname = $Surname -replace "O'k", "O'K"
        $Surname = $Surname -replace "O'n", "O'N"
        $Surname = $Surname -replace "O'r", "O'R"

        # Set remaining details
        $FullName =  "${PreferredName} ${Surname} (${LoginName}) - Parent"
        $FullAlias = ("${PreferredName} ${Surname}") -replace " ", "_"
        $DefaultMail = "${FullAlias}@vnc.qld.edu.au"
        $YEAR = [string](Get-Date).Year
    
    ### Process Current Users ###
    
    If ($Termination.length -eq 0) {        
        If (Get-ADUser -Filter { (SamAccountName -eq $LoginName) }) {
            # Set user to confirm details
            $TestUser = Get-ADUser -Identity $LoginName -property *
            $ParCompany = $TestUser.Company
            $ParDept = "Parent"
            $ParTitle = "Parent - ${YEAR}"
            # Check the paraddress file for matching parent codes.
            foreach($parline in $input)
            {
                $ParEmail = $null
                $ParName = $null
                $ParAlias = $null
                $ParFullName = $null
                $ParNum = (Get-Culture).TextInfo.ToLower($parline.add_num.Trim())
                $ParCode = (Get-Culture).TextInfo.ToLower($parline.par_code.Trim())
                $contact = $null

                if (($TestParCode -eq $ParCode) -and ($ParNum -eq '1')) {
                    $ParName = $parline.salutation.Trim()
                    $ParEmail = (Get-Culture).TextInfo.ToLower($parline.e_mail.Trim())
                    $ParEmail = ($ParEmail.Split(';'))[0]
                    $ParFullName = "${FullName}${ParNum}"
                    $ParAlias = "Parent${ParNum}-${FullAlias}"
                    $emailsearch = $null
                    # Process email contacts
                    if ((!($ParEmail -eq $null)) -and ($ParEmail.length -gt 3) -and (!($ParFullName -eq $null))) {
                        # Search by name
                        try {
                            write-host "Looking for contact ${ParFullName}"
                            $contact = Get-MailContact -Identity $ParFullName
                        }
                        catch {
                            write-host "Failed"
                            $contact = $null
                        }
                        finally{
                            write-host "found contact: " $contact
                            write-host ""
                        }
                        #write-host " -0- "
                        if ($contact -eq $null) {
                            # Search by Email if name is not found
                            try {
                                write-host "Looking for email ${ParEmail}"
                                $emailsearch = Get-Mailcontact | where {$_.PrimarySmtpAddress -like $ParEmail}
                            }
                            catch {
                                write-host "Failed"
                                $emailsearch = $null
                            }
                            finally{
                                write-host "found mail: " $emailsearch
                                write-host ""
                            }
                        }
                        #write-host " -1- "
                        if ($contact) {
                            write-host $contact.PrimarySmtpAddress
                            if (!($contact.PrimarySmtpAddress -eq $ParEmail)) {
                                write-host "updating email address"
                                Set-MailContact -Identity $ParFullName -ExternalEmailAddress $ParEmail
                            }
                        }
                        # Update Contact for user If contact is present
                        if ($emailsearch) {
                            write-host "Updating Parent Name: " . ${$emailsearch}.Identity . " for ${ParFullName}"
                            Set-MailContact -Identity $emailsearch.Identity -Name $ParFullName -Alias $ParAlias
                            write-host "Updating Parent details"
                            Set-Contact -Identity $emailsearch.Identity -FirstName $ParName -Department $ParDept -Title $ParTitle
                        }
                        #write-host " -2- "
                        # Create Contact for user If contact is missing
                        if ((!($contact)) -and (!($emailsearch))) {
                            write-host "New Parent Email: ${ParEmail} for ${ParFullName}"
                            New-MailContact -ExternalEmailAddress $ParEmail -Name $ParFullName -OrganizationalUnit $ContactOU -Alias $ParAlias
                            Set-MailContact -Identity $ParFullName -EmailAddressPolicyEnabled $false -UseMapiRichTextFormat 'Never'
                        }
                    #write-host " -3- "
                    # get contact details.
                    $tempcontact = $contact
                    $contactCompany = $tempcontact.Company
                    $contactDepartment = $tempcontact.Department
                    $contactTitle = $tempcontact.Title
                    $contactName = $tempcontact.FirstName
                    $tempcontact2 = $emailsearch
                    $contactAddresses = $tempcontact2.EmailAddresses
                    
                    #write-host " -4- "
                    # Update existing contacts if different
                    if ((!($tempcontact -eq $null)) -and (!($tempcontact2 -eq $null)) -and (!($ParFullName -eq $null))) {
                        
                        # Remove contacts with no current email address
                        if (($tempcontact2) -and ($ParEmail -eq $null)) {
                            Remove-MailContact -Identity $ParFullName -Confirm:$false
                            write-host $LoginName "Removed Parent Contact"
                        }
                        
                        # Remove the default VNC email policy address (if present)
                        If (!($contactAddresses -eq $null)) {
                            if ($contactAddresses.Contains($DefaultMail)) {
                                Set-MailContact -Identity $ParFullName -EmailAddresses @{Remove=$DefaultMail}
                                write-host "${FullName} removed vnc Email Address" 
                            }
                        }
                    }
                    }
                }
            }
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
                $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

                If (!($TestUser.Enabled)) {
                    # Check the paraddress file for matching parent codes.
                    foreach($parline in $input)
                    {
                        $ParEmail = $null
                        $ParName = $null
                        $ParAlias = $null
                        $ParFullName = $null

                        if ($TestParCode -eq (Get-Culture).TextInfo.ToLower($parline.par_code.Trim())) {
                            $ParName = $parline.salutation.Trim()
                            $ParEmail = (Get-Culture).TextInfo.ToLower($parline.e_mail.Trim())
                            $ParEmail = ($ParEmail.Split(';'))[0]
                            $ParFullName = "${FullName}${parline.add_num}"
                            $ParAlias = "Parent-${parline.add_num}${FullAlias}"
                        }
                    }
                    
                    if (!($ParFullName -eq $null)) {
                        write-host "Checking disabled students for ${ParFullName}"
                        $TestContact = Get-MailContact -Identity "$ParFullName"
                    }
                    else {
                        $TestContact = $null
                    }
                        
                    if ($TestContact) {
                        Remove-MailContact -Identity $ParFullName -Confirm:$false
                        write-host $LoginName "Removed Parent Contact"
                    }
                }
            }
        }
        ElseIf ((!($DATE -gt $Termination)) -and (!($Termination.length -eq 0))) {
            write-host "Not Final Leaving Date", $LoginName
            write-host $DATE
            write-host $Termination
        }
    }
}

write-host "### Contact Creation Script Finished"
write-host
