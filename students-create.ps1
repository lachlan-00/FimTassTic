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
$input = Import-CSV ".\csv\fim_student_filtered.csv" -Encoding UTF8
$idinput = Import-CSV  ".\csv\_CUSTOM_STUDENT_ID.csv" -Encoding UTF8

write-host
write-host "### Starting Student Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for different user types
$DisablePath = "OU=student,OU=users,OU=disabled,DC=example,DC=com,DC=au"
$5Path = "OU=year5,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$6Path = "OU=year6,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$7Path = "OU=year7,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$8Path = "OU=year8,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$9Path = "OU=year9,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$10Path = "OU=year10,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$11Path = "OU=year11,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
$12Path = "OU=year12,OU=student,OU=UserAccounts,DC=example,DC=com,DC=au"
# Security Group names for students
$StudentName = "CN=Students,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$5Name = "CN=S-G_year5,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$6Name = "CN=S-G_year6,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$7Name = "CN=S-G_year7,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$8Name = "CN=S-G_year8,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$9Name = "CN=S-G_year9,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$10Name = "CN=S-G_year10,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$11Name = "CN=S-G_year11,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$12Name = "CN=S-G_year12,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$UserAdmin = "CN=Local-Users-Administrators,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$UserPower = "CN=Local-Users-Power_Users,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$UserRegular = "CN=Local-Users-Users,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$MoodleStudent = "CN=MoodleStudent,OU=RoleAssignment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au"
# Get membership for group Membership Tests
$VillanovaGroups = Get-ADGroup -Filter *  -SearchBase "OU=UserGroups,DC=example,DC=com,DC=au"
$StudentGroup = Get-ADGroupMember -Identity $StudentName
$5Group = Get-ADGroupMember -Identity $5Name
$6Group = Get-ADGroupMember -Identity $6Name
$7Group = Get-ADGroupMember -Identity $7Name
$8Group = Get-ADGroupMember -Identity $8Name
$9Group = Get-ADGroupMember -Identity $9Name
$10Group = Get-ADGroupMember -Identity $10Name
$11Group = Get-ADGroupMember -Identity $11Name
$12Group = Get-ADGroupMember -Identity $12Name
$LocalAdmin = Get-ADGroupMember -Identity $UserAdmin
$LocalPower = Get-ADGroupMember -Identity $UserPower
$LocalUser = Get-ADGroupMember -Identity $UserRegular
$MoodleStudentMembers = Get-ADGroupMember -Identity $MoodleStudent
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
$DATE = "${YEAR}/${MONTH}/${DAY}"
$DATE = "${DATE} 00:00:00"

write-host "### Completed importing groups"
write-host

################################################
### Create / Edit / Disable student accounts ###
################################################

foreach($line in $input) {

    # UserCode is the Unique Identifier for Students
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
            $UserPath = $5Path
            $ClassGroup = $5Name
        }
        IF ($YearGroup -eq "6") {
            $UserPath = $6Path
            $ClassGroup = $6Name
        }
        IF ($YearGroup -eq "7") {
            $UserPath = $7Path
            $ClassGroup = $7Name
        }
        IF ($YearGroup -eq "8") {
            $UserPath = $8Path
            $ClassGroup = $8Name
        }
        IF ($YearGroup -eq "9") {
            $UserPath = $9Path
            $ClassGroup = $9Name
        }
        IF ($YearGroup -eq "10") {
            $UserPath = $10Path
            $ClassGroup = $10Name
        }
        IF ($YearGroup -eq "11") {
            $UserPath = $11Path
            $ClassGroup = $11Name
        }
        IF ($YearGroup -eq "12") {
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
        If ($LoginName -eq "11334") {
            $PreferredName = "Seán"
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
        $Surname = $Surname -replace "O'b", "O'B"
        $Surname = $Surname -replace "O'c", "O'C"
        $Surname = $Surname -replace "O'd", "O'D"
        $Surname = $Surname -replace "O'k", "O'K"
        $Surname = $Surname -replace "O'n", "O'N"
        $Surname = $Surname -replace "O'r", "O'R"

        # Set remaining details
        $FullName =  "${PreferredName} ${Surname}"
        $UserPrincipalName = "${LoginName}@example.com.au"
        $Position = "Year ${YearGroup}"
        # Home Folders are only for younger grades
        IF (($YearGroup -eq "5")-or ($YearGroup -eq "6")) {
            $HomeDrive = "\\example.com.au\home\Student\${LoginName}"
        }
        Else {
            $HomeDrive = $null
        }
        $JobTitle = "Student - ${YEAR}"

        ########################################
        ### Create / Modify Student Accounts ###
        ########################################

        # Create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            if (!($HomeDrive -eq $null)) {
                Try  {
                    New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "mypassword" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
                    write-host "${LoginName} created for ${FullName}"
                    write-host
                }
                Catch {
                    write-host "${LoginName} already exists for ${FullName}"
                    write-host
                }
            }
            Else {
                Try  {
                    New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "mypassword" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $True
                    write-host "${LoginName} created for ${FullName}"
                    write-host
                }
                Catch {
                    write-host "${LoginName} already exists for ${FullName}"
                    write-host
                }
            }
        }

        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

        # Get User Details
        $TestName = $TestUser.Name
        $TestGiven = $TestUser.GivenName
        $TestSurname = $TestUser.SurName
        $TestDisplayName = $TestUser.DisplayName
        $TestDN = $TestUser.distinguishedname
        $TestAccountName = $TestUser.SamAccountName
        $TestEnabled = $TestUser.Enabled
        $TestTitle = $TestUser.Title
        $TestCompany = $TestUser.Company
        $TestOffice = $TestUser.Office
        $TestDescription = $TestUser.Description
        $TestDepartment = $TestUser.Department
        $TestNumber = $TestUser.employeeNumber
        $TestID = $TestUser.employeeID

        # set additional user details if the user exists
        If ($TestUser) {

            # Enable user if disabled
            If ((!($TestEnabled)) -and (!($TestDescription -eq "disable"))) {
                Set-ADUser -Identity $LoginName -Enabled $true
                write-host "Enabling", $TestAccountName
            }

            # Move user to the default OU for their year level if not there
            if (($TestEnabled) -and (!($TestDN.Contains($UserPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                write-host $TestAccountName
                write-host "Taking From: ${TestDN}"
                write-host "Moving To:" $UserPath
            }

            If ((!($TestDescription -eq $UserCode)) -and (!($TestDescription -eq "disable"))) {
                Set-ADUser -Identity $LoginName -Description $UserCode
                write-host $FullName, "changing description from ${TestDescription} to ${UserCode}"
                write-host
            }

            # Check Name Information
            If ($TestGiven -cne $PreferredName) {
                Set-ADUser -Identity $LoginName -GivenName $PreferredName
                write-host "${TestAccountName} Changed Given Name to ${PreferredName}"
            }
            If ($TestSurname -cne $Surname) {
                Set-ADUser -Identity $LoginName -Surname $Surname
                write-host "${TestAccountName} Changed Surname to ${SurName}"
            }
            If (($TestName -cne $FullName)) {
                Rename-ADObject -Identity $TestAccountName -NewName $FullName
                write-host "${TestAccountName} Changed Object Name to: ${FullName}"
            }
            If (($TestDisplayName -cne $FullName)) {
                Set-ADUser -Identity $LoginName -DisplayName $FullName
                write-host "${TestAccountName} Changed Display Name to: ${FullName}"
            }

            # Set company for automatic mail group filtering
            if ($UserPath -eq $5Path) {
                if (!($TestCompany -ceq "year5")) {
                    Set-ADUser -Identity $LoginName -Company "year5"
                    write-host $TestName "set company to year5"
                }
            }
            if ($UserPath -eq $6Path) {
                if (!($TestCompany -ceq "year6")) {
                    Set-ADUser -Identity $LoginName -Company "year6"
                    write-host $TestName "set company to year6"
                }
            }
            if ($UserPath -eq $7Path) {
                if (!($TestCompany -ceq "year7")) {
                    Set-ADUser -Identity $LoginName -Company "year7"
                    write-host $TestName "set company to year7"
                }
            }
            if ($UserPath -eq $8Path) {
                if (!($TestCompany -ceq "year8")) {
                    Set-ADUser -Identity $LoginName -Company "year8"
                    write-host $TestName "set company to year8"
                }
            }
            if ($UserPath -eq $9Path) {
                if (!($TestCompany -ceq "year9")) {
                    Set-ADUser -Identity $LoginName -Company "year9"
                    write-host $TestName "set company to year9"
                }
            }
            if ($UserPath -eq $10Path) {
                if (!($TestCompany -ceq "year10")) {
                    Set-ADUser -Identity $LoginName -Company "year10"
                    write-host $TestName "set company to year10"
                }
            }
            if ($UserPath -eq $11Path) {
                if (!($TestCompany -ceq "year11")) {
                    Set-ADUser -Identity $LoginName -Company "year11"
                    write-host $TestName "set company to year11"
                }
            }
            if ($UserPath -eq $12Path) {
                if (!($TestCompany -ceq "year12")) {
                    Set-ADUser -Identity $LoginName -Company "year12"
                    write-host $TestName "set company to year12"
                }
            }

            # Set Year Level and Title
            If (($TestTitle) -eq $null) {
                Set-ADUser -Identity $LoginName -Title $JobTitle
                write-host $LoginName, "Title change to: ${JobTitle}"
            }
            ElseIf (!($TestTitle).contains($JobTitle)) {
                Set-ADUser -Identity $LoginName -Title $JobTitle
                write-host $LoginName, "Title change to: ${JobTitle}"
            }

            # Get the year level of the current office string
            If ($TestOffice) {
                $test1 = $TestOffice.Substring($TestOffice.length-1,1)
                $test2 = $TestOffice.Substring($TestOffice.length-2,2)
            }

            # set Office to current year level
            If ($YearGroup.length -eq 1) {
                If (($YearGroup) -ne ($test1)) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host $LoginName, "year level change from ${TestOffice} to ${Position}"
                }
            }
            ElseIf ($YearGroup.length -eq 2) {
                If (($YearGroup) -ne ($test2)) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host $LoginName, "year level change from ${TestOffice} to ${Position}"
                }
            }

            # Set Department to identify current students
            If (!(($TestDepartment) -ceq ("Student"))) {
                Set-ADUser -Identity $LoginName -Department "Student"
                write-host "${TestName} Setting Position from ${TestDepartment} to Student"
            }

            # Add Employee Number if there is one
            if (!($LoginName -ceq $TestNumber)) {
                Set-ADUser -Identity $LoginName -EmployeeNumber $LoginName
                write-host "Setting employee Number (${employeeNumber}) for ${TestAccountName}"
                write-host
            }

            # Check Group Membership
            if (!($StudentGroup.name.contains($TestName))) {
                Add-ADGroupMember -Identity "Students" -Member $LoginName
                write-host $LoginName "added Students Group"
            }
            if (!($LocalUser.name.contains($TestName))) {
                write-host $TestName, "Add to local user group for domain workstations"
                Add-ADGroupMember -Identity $UserRegular -Member $LoginName
            }
            # $MoodleStudentMembers
            if (!($MoodleStudentMembers.name.contains($TestName))) {
                Add-ADGroupMember -Identity $MoodleStudent -Member $LoginName
                write-host $LoginName "added MoodleStudent Group"
            }

            # Remove groups for other grades and add the correct grade
            IF ($YearGroup -eq "5") {
                # Add Correct Year Level
                if (!($5Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $5Name -Member $TestAccountName
                    write-host $LoginName "added 5"
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "6") {
                if ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($6Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $6Name -Member $TestAccountName
                    write-host $LoginName "added 6"
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "7") {
                if ($5Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($7Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $7Name -Member $TestAccountName
                    write-host $LoginName "added 7"
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "8") {
                if ($5Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($8Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $8Name -Member $TestAccountName
                    write-host $LoginName "added 8"
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "9") {
                if ($5Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($9Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $9Name -Member $TestAccountName
                    write-host $LoginName "added 9"
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "10") {
                if ($5Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($10Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $10Name -Member $TestAccountName
                    write-host $LoginName "added 10"
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "11") {
                if ($5Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($11Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $11Name -Member $TestAccountName
                    write-host $LoginName "added 11"
                }
                if ($12Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            IF ($YearGroup -eq "12") {
                if ($5Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                if ($6Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                if ($7Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                if ($8Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                if ($9Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                if ($10Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                if ($11Group.name.contains($TestName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                if (!($12Group.name.contains($TestName))) {
                    Add-ADGroupMember -Identity $12Name -Member $TestAccountName
                    write-host $LoginName "added 12"
                }
            }
        }
        Else {
            write-host "missing or ignoring ${FullName}: ${LoginName}"
            write-host
        }
        foreach($line in $idinput) {
            $tmpName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())
            $tmpID = (Get-Culture).TextInfo.ToUpper($line.idcard_nfc.Trim())
            If ($TestAccountName -eq $tmpName) {
                If ($TestUser) {
                    If (!($TestID -ceq $tmpID)) {
                        if (($tmpID -eq "") -or ($tmpID -eq "#null!")) {
                            #write-host "No ID Found for ${LoginName}"
                            Set-ADUser -Identity $LoginName -EmployeeID $null
                        }
                        else {
                            Set-ADUser -Identity $LoginName -EmployeeID $tmpID
                            write-host "Setting ID (${tmpID}) for ${LoginName}"
                            write-host
                        }
                    }
                }
            }
        }
    }

    ######################################
    ### Disable Students who have left ###
    ######################################

    Else {
        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
        $TestDN = $TestUser.distinguishedname
        $TestDescription = $TestUser.Description
        $TestEnabled = $TestUser.Enabled
        $TestAccountName = $TestUser.SamAccountName
        $TestMembership = $TestUser.MemberOf

        # Disable users with a termination date if they are still enabled
        If ($TestEnabled) {

            # Don't disable users we want to keep
            If ($TestDescription -eq "keep") {
                write-host "${LoginName} Keeping terminated user"
                write-host
            }
            # Terminate Students AFTER their Termination date
            ElseIf ($DATE -gt $Termination) {
                # Disable The account when we don't want to keep it
                If ($TestUser) {
                    Set-ADUser -Identity $LoginName -Enabled $false
                    write-host "DISABLING ACCOUNT ${$LoginName}"
                    write-host
                }
            }
        }
        Else {
            # Enforce Group and OU changes for disabled students
            If ($TestUser) {
                # Move to disabled user OU if not already there
                if (!($TestDN.Contains($DisablePath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                    write-host "Moving: ${TestAccountName} to Disabled Student OU"
                    write-host
                }

                # Remove groups if they are a member of any additional groups
                If ($TestMembership) {
                    write-host "Removing groups for ${TestAccountName}"
                    write-host
                    #remove All Villanova  Groups
                    Foreach($GroupName In $TestMembership) {
                        Try {
                            Remove-ADGroupMember -Identity $GroupName -Member $TestAccountName -Confirm:$false
                        }
                        Catch {
                            Write-Host "Error Removing ${GroupName}"
                        }
                    }
                }
            }
        }
    }
}

write-host
write-host "### Student Creation Script Finished"
write-host
