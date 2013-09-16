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

write-host
write-host "### Starting Student Echo Script"
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
$DenyAdmin = Get-ADGroupMember -Identity "S-G_Student-Deny-Admin"
$AllowAdmin = Get-ADGroupMember -Identity "S-G_Student-Admin"
# Get Date and Format Field to Match Termination Date
$YEAR = [string](Get-Date).Year
$MONTH = [string](Get-Date).Month
$DAY = [string](Get-Date).Day
If ($MONTH.length -eq 1) {
    $MONTH = "0${MONTH}"
}
If ($DAY.length -eq 1) {
    $DAY = "0${DAY}"
}
$DATE = "${YEAR}-${MONTH}-${DAY}"
$DATE = "${DATE} 00:00:00"

write-host "### Completed importing groups"
write-host

###############################################
### Create / Edit /Disable student accounts ###
###############################################

foreach($line in $input) {

    # UserCode/Loginname is the Unique Identifier for Students
    $UserCode = (Get-Culture).TextInfo.ToUpper($line.stud_code.Trim())

    # Correct usercode if opened in Excel
    If ($UserCode.Length -ne 5) {
            $UserCode = "0${UserCode}"
    }

    # Set Login Name
    $LoginName = $UserCode

    # Check Termination Dates
    $Termination = $line.dol.Trim()

    ################################
    ### Process Current Students ###
    ################################

    If ($Termination.length -eq 0) {

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
        $FullName =  "${PreferredName} ${Surname}"
        $UserPrincipalName = "${LoginName}@villanova.vnc.qld.edu.au"
        $Position = "Year ${YearGroup}"
        $HomeDrive = "\\vncfs01\studentdata$\${YearIdent}\${LoginName}"
        $JobTitle = "Student - ${YEAR}"

        ########################################
        ### Create / Modify Student Accounts ###
        ########################################
        
        # Create basic user if you can't find one
        If (!(Get-ADUser -Filter { (SamAccountName -eq $LoginName)})) {
            #New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "Abc123" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
            write-host "${LoginName} created for ${FullName}"
        }

        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

        # Check Name Information
        If ($TestUser) {
            If ($TestUser.GivenName -cne $PreferredName) {
                #write-host $TestUser.GivenName, "Changed Given Name to ${PreferredName}"
                #Set-ADUser -Identity $TestUser.SamAccountName -GivenName $PreferredName
            }
            If ($TestUser.Surname -cne $Surname) {
                write-host $TestUser.Surname, "Changed Surname to ${Surname}"
                #Set-ADUser -Identity $TestUser.SamAccountName -Surname $Surname
            }
            If (($TestUser.Name -cne $FullName)) {
                write-host $TestUser.Name, "Changed Full Name Name to ${FullName}"
                #Set-ADUser -Identity $TestUser.SamAccountName -DisplayName $FullName
            }
            If ($TestUser.CN -cne $FullName) {
                 write-host $TestUser.Name, "Changed Common Name to ${FullName}"
                 write-host
            }
        }

        # set additional user details if the user exists
        If ($TestUser) {
            # Enable use if disabled
            If (!($TestUser.Enabled)) {
                #Set-ADUser -Identity $TestUser.SamAccountName -Enabled $true
                write-host "Enabling", $TestUser.SamAccountName
            }

            # Move user to the default OU if not already there
            if (!($TestUser.distinguishedname.Contains($UserPath))) {
                #Get-ADUser $TestUser.SamAccountName | Move-ADObject -TargetPath $UserPath
                write-host $TestUser.SamAccountName
                write-host "Taking From:" $TestUser.distinguishedname
                write-host "Moving To:" $UserPath
            }
            # Set Year Level and Title
            If (!($TestUser.Title).contains($JobTitle)) {
                #Set-ADUser -Identity $TestUser.SamAccountName -Title $JobTitle
                write-host $LoginName, "Title change to Student - ${YEAR}"
            }
            If (!($TestUser.Office) -eq ("${YearGroup}")) {
                #Set-ADUser -Identity $TestUser.SamAccountName -Office $YearGroup
                write-host $TestUser.SamAccountName, "year level change to ${YearGroup}" 
            }

            # Check Group Membership
            if (!($StudentGroup.name.contains($TestUser.name))) {
                #Add-ADGroupMember -Identity Students -Member $TestUser.SamAccountName
                write-host $TestUser.SamAccountName "added Students Group"
            }
            # Check user for admin rights.
            if ((!($DenyAdmin.name.contains($TestUser.name))) -and (($AllowAdmin.name.contains($TestUser.name)))) {
                #Add-ADGroupMember -Identity "S-G_Student-Admin" -Member $LoginName
                write-host "${LoginName} added to the local deny admin group and removed form admins"
            }
            # Remove groups for other grades and add the correct grade
            IF ($YearGroup -eq "5") {
                # Add Correct Year Level
                if (!($5Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 5"
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "6") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($6Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 6"
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "7") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($7Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 7"
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "8") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($8Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 8"
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "9") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($9Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 9"
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "10") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($10Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 10"
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "11") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($11Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 11"
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "12") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($12Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName
                    write-host $TestUser.SamAccountName "added 12"
                }
            }
        }
        Else {
            write-host "missing ${FullName}"
            write-host $TestUser.SamAccountName
            write-host $UserCode
            write-host
        }
    }

    ######################################
    ### Disable Students who have left ###
    ######################################

    Else {

        # Disable users with a termination date if they are still enabled
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }) {

            # Terminate Students AFTER their Termination date
            If ($DATE -gt $Termination) {
                write-host "DISABLING ACCOUNT ${TestUser.SamAccountName}"
                write-host $UserCode
                write-host $DATE
                write-host $Termination
                write-host

                # Set user to confirm details
                $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

                If (!($TestUser -eq $null)) {
                    if (!($TestUser.distinguishedname.Contains($DisablePath))) {

                        # Move to disabled user OU if not already there
                        #Get-ADUser $TestUser.SamAccountName | Move-ADObject -TargetPath $DisablePath
                        write-host $TestUser.SamAccountName "MOVED to Disabled OU"
                    }

                    # Check Group Membership
                    if ($StudentGroup.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity "Students" -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED Students"
                    }
                    if ($5Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $5Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 5"
                    }
                    if ($6Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $6Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 6"
                    }
                    if ($7Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $7Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 7"
                    }
                    if ($8Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $8Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 8"
                    }
                    if ($9Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $9Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 9"
                    }
                    if ($10Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $10Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 10"
                    }
                    if ($11Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $11Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 11"
                    }
                    if ($12Group.name.contains($TestUser.name)) {
                        #Remove-ADGroupMember -Identity $12Name -Member $TestUser.SamAccountName -Confirm:$false
                        write-host $TestUser.SamAccountName "REMOVED 12"
                    }

                    # Disable The account
                    #Set-ADUser -Identity $TestUser.SamAccountName -Enabled $false
                }
            }
        }
    }
}

write-host
write-host "### Student Echo Script Finished"
write-host
