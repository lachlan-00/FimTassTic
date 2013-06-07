###############################################################################
###                                                                         ###
###  Create Staff Accounts From TASS.web Data                               ###
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
$input = Import-CSV  .\csv\telemf.csv
$teacherinput = Import-CSV  .\csv\teacher.csv

write-host "### Starting Staff Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$UserPath = "OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TeacherPath = "OU=Teachers,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$NonTeacherPath = "OU=Teachers,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$ReliefTeacherPath = "OU=Relief and Preservice Teachers,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TutorPath = "OU=Tutors,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$DisablePath = "OU=staff,OU=Disabled,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
# Get membership for group Membership Tests
$TestStaff = Get-ADGroupMember -Identity "Staff"
$TestAllStaff = Get-ADGroupMember -Identity "All Staff"
$TestTeachers = Get-ADGroupMember -Identity "Teachers"
$TestAllTeachers = Get-ADGroupMember -Identity "Teachers - All"

write-host "### Completed importing groups"
write-host

##############################################
### Create / Edit / Disable Staff accounts ###
##############################################

foreach($line in $input) {

    ################################
    ### Configure User Variables ###
    ################################

    # Set lower case because powershell ignores uppercase word changes to title case
    $PreferredName = (Get-Culture).TextInfo.ToUpper($line.prefer_name_text.Trim())
    $Surname = (Get-Culture).TextInfo.ToUpper($line.surname_text.Trim())
    $Position = (Get-Culture).TextInfo.ToUpper($line.position_title.Trim())
    $Position2 = (Get-Culture).TextInfo.ToUpper($line.position_text.Trim())

    if ($PreferredName -eq $line.prefer_name_text.Trim()) {
        $PreferredName = (Get-Culture).TextInfo.ToLower($PreferredName)
        $PreferredName = (Get-Culture).TextInfo.ToTitleCase($PreferredName)
    }
    Else {
        $PreferredName = ($line.prefer_name_text.Trim())
    }
    if (($Surname) -eq $line.surname_text.Trim()) {
        $Surname = (Get-Culture).TextInfo.ToLower($Surname)
        $Surname = (Get-Culture).TextInfo.ToTitleCase($Surname)
    }
    Else {
        $Surname = ($line.surname_text.Trim())
    }
    If (($Position -ne $null) -and ($Position -ne '')) {
        if (($Position) -eq $line.position_title.Trim()) {
            $Position = (Get-Culture).TextInfo.ToLower($Position)
            $Position = (Get-Culture).TextInfo.ToTitleCase($Position)
            }
        Else {
            $Position = ($line.position_title.Trim())
        }
    }
    Elseif (($Position2 -ne $null) -and ($Position2 -ne '')) {
        if (($Position2) -eq $line.position_text.Trim()) {
            $Position2 = (Get-Culture).TextInfo.ToLower($Position2)
            $Position2 = (Get-Culture).TextInfo.ToTitleCase($Position2)
            $Position = $Position2
            }
        Else {
            $Position = ($line.position_text.Trim())
        }
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

    $FullName =  "${PreferredName} ${Surname}"
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())
    $UserPrincipalName = $LoginName + "@villanova.vnc.qld.edu.au"

    # Pull remaining details
    $Position = (Get-Culture).TextInfo.ToTitleCase($Position.Trim())
    $HomeDrive = "\\vncfs01\staffdata$\" + $LoginName
    $Telephone = $line.phone_w_text.Trim()
    If ($Telephone.length -le 1) {
        $Telephone = $null
    }

    # Check Termination Dates
    $Termination = $line.term_date.Trim()
    $YEAR = [string](Get-Date).Year
    $MONTH = [string](Get-Date).Month
    $DAY = [string](Get-Date).Day

    # Format $Date Field to Match Termination Date
    If ($MONTH.length -eq 1) {
        $MONTH = "0${MONTH}"
    }
    If ($DAY.length -eq 1) {
        $DAY = "0${DAY}"
    }
    $DATE = "${YEAR}-${MONTH}-${DAY}"
    $DATE = $DATE, "00:00:00"

    ######################################
    ### Create / Modify Staff Accounts ###
    ######################################

    If ($Termination.length -eq 0) {
        # create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            $LoginName = (Get-Culture).TextInfo.ToLower($LoginName)
            New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "Abc123" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
            Set-ADUser -Identity $LoginName -Description $Position -Office $Position -Title $Position
            write-host $LoginName, " created for ", $FullName
        }

        # Check Name Information
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
        If ($TestUser) {
            If ($TestUser.GivenName -cne $PreferredName) {
                write-host $TestUser.GivenName, "Changed to" $PreferredName
                Set-ADUser -Identity $LoginName -GivenName $PreferredName
            }
            If ($TestUser.Surname -cne $Surname) {
                write-host $TestUser.SurName, "Changed to" $SurName
                Set-ADUser -Identity $LoginName -Surname $Surname
            }
            If ($TestUser.Name -cne $FullName) {
                write-host $TestUser.Name, "Changed to" $FullName
                Set-ADUser -Identity $LoginName -DisplayName $FullName
            }
            If ($TestUser.CN -cne $FullName) {
                 write-host $LoginName "Changed common name", $FullName
                 write-host
            }
        }

        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Name -eq $FullName)) } -Properties *)
        $TestPhone = $TestUser.OfficePhone
        $TestDescription = $TestUser.Description

        # set additional user details if the user exists
        If ($TestUser) {
            # Enable use if disabled
            If (!($TestUser.Enabled)) {
                #Set-ADUser -Identity $LoginName -Enabled $true
                write-host "Enabling", $LoginName
            }

            # Add Position if there is one
            if (!($Position -ceq $TestDescription)) {
                write-host $LoginName, "setting position"
                write-host $Position
                write-host $TestDescription
                write-host
                Set-ADUser -Identity $LoginName -Description $Position ######## -Office $Position -Title $Position
            }

            # Add Telephone number if there is one
            if ($Telephone -ne $TestPhone) {
                If ($Telephone -eq $null) {
                    if ($TestPhone -ne '690') {
                        Set-ADUser -Identity $LoginName -OfficePhone '690'
                        write-host $LoginName, "setting Telephone to Default (690)"
                        write-host
                    }
                }
                Else {
                    Set-ADUser -Identity $LoginName -OfficePhone $Telephone
                    write-host $LoginName, "setting Telephone to:", $Telephone
                    write-host
                }
            }

            # Move user to their default OU if not already there
            if ($TestUser.distinguishedname.Contains($DisablePath)) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $UserPath
                write-host $LoginName "moved out of Disabled OU"
            }
            ElseIf ($Position.Contains("Relief") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $ReliefTeacherPath
                write-host $LoginName "moved to Relief Teacher OU"
            }
            ElseIf (($Position.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $TutorPath
                write-host $LoginName "moved to Music Tutor OU"
            }
            
            # Check Group Membership
            if (!($TestStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "Staff" -Member $LoginName
                        write-host $LoginName "added Staff"
            }
            if (!($TestAllStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "All Staff" -Member $LoginName
                        write-host $LoginName "added allstaff"
            }
        }
    }

    ###################################
    ### Disable Staff who have left ###
    ###################################
    
    Else {

        # Disable users with a termination date if they are still enabled
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }) {

            # Terminate Staff AFTER their Termination date
            If ($DATE -gt $Termination) {
                write-host "DISABLING ACCOUNT, '$($LoginName)'"

                # Set user to confirm details
                $TestUser = Get-ADUser -Identity $LoginName

                if (!($TestUser.distinguishedname.Contains($DisablePath))) {
                    # Move to disabled user OU if not already there
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $DisablePath
                }

                # Check Group Membership
                if ($TestStaff.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "Staff" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED Staff"
                }
                if ($TestAllStaff.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "All Staff" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED allstaff"
                }
                if ($TestTeachers.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "Teachers" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED Teachers"
                }
                if ($TestAllTeachers.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "Teachers - All" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED Teachers - All"
                }
            
                # Disable The account
                Set-ADUser -Identity $LoginName -Enabled $false
            }
            Else {
                write-host "Not Final Leaving Date", $FullName
                write-host $DATE
                write-host $Termination
                write-host
            }
        }
    }
}

write-host
write-host "### Staff Creation Finished"
write-host
write-Host "### Starting Teacher Modification"
write-host

###################################################################
### Create / Edit Teacher Info for Staff with existing accounts ###
###################################################################

foreach($line in $teacherinput)
{
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())

    If (($LoginName -ne $null) -and ($LoginName -ne '')) {

        # Set user to confirm details
        $TestUser = (Get-ADUser  -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }  -Properties *)
        $Description = $TestUser.Description

        If ($TestUser.Enabled) {

            # Move to Teacher OU if not already there
            if ($TestUser.distinguishedname.Contains($UserPath) -and (!($TestUser.distinguishedname.Contains($TeacherPath))) -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                If ($Description.Contains("Relief") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $ReliefTeacherPath
                    write-host $LoginName "moved to Relief Teacher OU"
                }
                ElseIf (($Description.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $TutorPath
                    write-host $LoginName "moved to Music Tutor OU"
                }
                ElseIf (!($TestUser.distinguishedname.Contains($TeacherPath)) -and (!($Description.Contains("Tutor")))) {
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $TeacherPath
                    write-host $LoginName "moved to Teacher OU"
                }
            }
            # Check Group Membership
            if (!($TestTeachers.name.contains($TestUser.name))) {
                Add-ADGroupMember -Identity "Teachers" -Member $LoginName
                write-host $LoginName "ADDED Teachers"
            }
            if (!($TestAllTeachers.name.contains($TestUser.name))) {
                Add-ADGroupMember -Identity "Teachers - All" -Member $LoginName
                write-host $LoginName "ADDED Teachers - All"
            }
        }
    }
}

write-Host "### Teacher Modification Finished"
write-host
write-host "### Staff Creation Script Finished"
write-host
