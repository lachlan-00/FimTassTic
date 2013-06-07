###############################################################################
###                                                                         ###
###  Create Student Accounts From TASS.web Data                             ###
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
$input = Import-CSV ".\csv\student.csv"

write-host "### Starting Student Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$DisablePath = "OU=students,OU=Disabled,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$5Path = "OU=Group H,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$6Path = "OU=Group G,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$7Path = "OU=Group F,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$8Path = "OU=Group E,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$9Path = "OU=Group D,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$10Path = "OU=Group C,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$11Path = "OU=Group B,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$12Path = "OU=Group A,OU=Students,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
# Security Group names for students
$5Name = "S-G_Group H"
$6Name = "S-G_Group G"
$7Name = "S-G_Group F"
$8Name = "S-G_Group E"
$9Name = "S-G_Group D"
$10Name = "S-G_Group C"
$11Name = "S-G_Group B"
$12Name = "S-G_Group A"
# Get membership for group Membership Tests
$StudentGroup = Get-ADGroupMember -Identity "Students"
$5Group = Get-ADGroupMember -Identity "S-G_Group H"
$6Group = Get-ADGroupMember -Identity "S-G_Group G"
$7Group = Get-ADGroupMember -Identity "S-G_Group F"
$8Group = Get-ADGroupMember -Identity "S-G_Group E"
$9Group = Get-ADGroupMember -Identity "S-G_Group D"
$10Group = Get-ADGroupMember -Identity "S-G_Group C"
$11Group = Get-ADGroupMember -Identity "S-G_Group B"
$12Group = Get-ADGroupMember -Identity "S-G_Group A"

write-host "### Completed importing groups"
write-host

################################################
### Create / Edit / Disable student accounts ###
################################################

foreach($line in $input) {

    ################################
    ### Configure User Variables ###
    ################################

    # Get Year level information for groups and home drive
    $YearGroup = $line.year_grp
    IF ($YearGroup -eq "5") {
        $YearIdent = "h"
        $UserPath = $5Path
        $ClassGroup = $5Name
    }
    IF ($YearGroup -eq "6") {
        $YearIdent = "g"
        $UserPath = $6Path
        $ClassGroup = $6Name
    }
    IF ($YearGroup -eq "7") {
        $YearIdent = "f"
        $UserPath = $7Path
        $ClassGroup = $7Name
    }
    IF ($YearGroup -eq "8") {
        $YearIdent = "e"
        $UserPath = $8Path
        $ClassGroup = $8Name
    }
    IF ($YearGroup -eq "9") {
        $YearIdent = "d"
        $UserPath = $9Path
        $ClassGroup = $9Name
    }
    IF ($YearGroup -eq "10") {
        $YearIdent = "c"
        $UserPath = $10Path
        $ClassGroup = $10Name
    }
    IF ($YearGroup -eq "11") {
        $YearIdent = "b"
        $UserPath = $11Path
        $ClassGroup = $11Name
    }
    IF ($YearGroup -eq "12") {
        $YearIdent = "a"
        $UserPath = $12Path
        $ClassGroup = $12Name
    }

    # Set lower case because powershell ignores uppercase word changes
    if ((Get-Culture).TextInfo.ToUpper($line.preferred_name.Trim()) -eq $line.preferred_name.Trim()) {
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
    if ((Get-Culture).TextInfo.ToUpper($line.surname.Trim()) -eq $line.surname.Trim()) {
        $Surname = (Get-Culture).TextInfo.ToLower($line.surname.Trim())
        $Surname = (Get-Culture).TextInfo.ToTitleCase($line.surname.Trim())
    }
    Else {
        $Surname = ($line.surname.Trim())
    }

    # Set Login and display names/text
    if (((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "").length -gt 3)
        {
        $LoginName = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "").substring(0,4) + $PreferredName.substring(0,2)
        $Test0 = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "").substring(0,4) + $GivenName.substring(0,2)
        }
    else
        {
        $LoginName = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "") + $PreferredName.substring(0,2)
        $Test0 = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "") + $GivenName.substring(0,2)
        }

    # Check for given name
    If ($LoginName -notcontains $Test0) {
        If (Get-ADUser -Filter { SamAccountName -eq $Test0 }) {
            $LoginName = (Get-Culture).TextInfo.ToLower($Test0)
        }
        Else {
            $LoginName = (Get-Culture).TextInfo.ToLower($LoginName)
        }
    }
    Else {
        $LoginName = (Get-Culture).TextInfo.ToLower($LoginName)
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

    $FullName =  ($PreferredName + " " + $Surname)
    $UserPrincipalName = $LoginName + "@villanova.vnc.qld.edu.au"

    # Correct usercode if opened in Excel
    $UserCode = (Get-Culture).TextInfo.ToUpper($line.stud_code.Trim())
    If ($UserCode.Length -ne 5) {
        $UserCode = "0${UserCode}"
    }

    # pull remaining details
    $Position = 'Year', $YearGroup.Trim()
    $HomeDrive = "\\vncfs01\studentdata$\" + $YearIdent + "\" + $LoginName

    # Check Termination Dates
    $Termination = $line.dol.Trim()
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

    ########################################
    ### Create / Modify Student Accounts ###
    ########################################

    If ($Termination.length -eq 0) {
        # Check for existing users before creating new ones.
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            write-host "missing login", $LoginName
            $LoginName = $UserCode
            $UserPrincipalName = $LoginName + "@villanova.vnc.qld.edu.au"
            $HomeDrive = "\\vncfs01\studentdata$\" + $YearIdent + "\" + $LoginName
            $Position = 'Year', $YearGroup
        }
        
        # Create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "Abc123" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
            write-host $LoginName, " created for ", $FullName
        }

        # Check Name Information
        $TestUser = (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode)) } -Properties *)
        If ($TestUser) {
            If ($TestUser.GivenName -cne $PreferredName) {
                write-host $TestUser.GivenName, "Changed to" $PreferredName
                Set-ADUser -Identity $LoginName -GivenName $PreferredName
            }
            If ($TestUser.Surname -cne $Surname) {
                write-host $TestUser.SurName, "Changed to" $SurName
                Set-ADUser -Identity $LoginName -Surname $Surname
            }
            If (($TestUser.Name -cne $FullName)) {
                write-host $TestUser.Name, "Changed to" $FullName
                Set-ADUser -Identity $LoginName -DisplayName $FullName
            }
            If ($TestUser.CN -cne $FullName) {
                 write-host $LoginName "Changed common name", $FullName
                 write-host
            }
        }
        
        # Set user to confirm details
        $TestUser = Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode) -and (Name -eq $FullName)) }

        # set additional user details if the user exists
        If ($TestUser) {
            # Enable use if disabled
            If (!($TestUser.Enabled)) {
                Set-ADUser -Identity $LoginName -Enabled $true
                write-host "RE-ENABLING", $($LoginName)
            }

            # Move user to the default OU if not already there
            if (!($TestUser.distinguishedname.Contains($UserPath))) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $UserPath
                write-host $LoginName
                write-host "Taking From:" $TestUser.distinguishedname
                write-host "Moving To:" $UserPath
            }

            # Set Year Level and Title
            Set-ADUser -Identity $LoginName -Office $YearGroup -Title "Student - ${YEAR}"

            # Check Group Membership
            if (!($StudentGroup.name.contains($TestUser.name))) {
                Add-ADGroupMember -Identity Students -Member $LoginName
                write-host $LoginName "added Students Group"
            }
            # Remove groups for other grades and add the correct grade
            IF ($YearGroup -eq "5") {
                # Add Correct Year Level
                if (!($5Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $5Name -Member $LoginName
                    write-host $LoginName "added 5"
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "6") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($6Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $6Name -Member $LoginName
                    write-host $LoginName "added 6"
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "7") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($7Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $7Name -Member $LoginName
                    write-host $LoginName "added 7"
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "8") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($8Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $8Name -Member $LoginName
                    write-host $LoginName "added 8"
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "9") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($9Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $9Name -Member $LoginName
                    write-host $LoginName "added 9"
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "10") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($10Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $10Name -Member $LoginName
                    write-host $LoginName "added 10"
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "11") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($11Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $11Name -Member $LoginName
                    write-host $LoginName "added 11"
                }
                if ($12Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "12") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($12Group.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $12Name -Member $LoginName
                    write-host $LoginName "added 12"
                }
            }
        }
        Else {
            write-host "missing", $FullName
            write-host $LoginName
            write-host $UserCode
            write-host
        }
    }

    ######################################
    ### Disable Students who have left ###
    ######################################

    Else {
        # Disable users with a termination date if they are still enabled
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True") -and (Name -eq $FullName)) }) {

            # Terminate Students AFTER their Termination date
            If ($DATE -gt $Termination) {
                write-host "DISABLING ACCOUNT, '$($LoginName)'"\

                # Set user to confirm details
                $TestUser = Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode)) }

                If (!($TestUser -eq $null)) {
                    if (!($TestUser.distinguishedname.Contains($DisablePath))) {

                        # Move to disabled user OU if not already there
                        Get-ADUser $LoginName | Move-ADObject -TargetPath $DisablePath
                        write-host $LoginName "MOVED to Disabled OU"
                    }

                    # Check Group Membership
                    if ($StudentGroup.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "Students" -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED Students"
                    }
                    if ($5Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 5"
                    }
                    if ($6Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 6"
                    }
                    if ($7Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 7"
                    }
                    if ($8Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 8"
                    }
                    if ($9Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 9"
                    }
                    if ($10Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 10"
                    }
                    if ($11Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 11"
                    }
                    if ($12Group.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                        write-host $LoginName "REMOVED 12"
                    }

                    # Disable The account
                    Set-ADUser -Identity $LoginName -Enabled $false
            }
        }
    }
}

write-host "### Student Creation Script Finished"
write-host
