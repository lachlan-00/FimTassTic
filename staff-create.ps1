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
$input = Import-CSV  ".\csv\fim_staffALL.csv" -Encoding UTF8
$classinput = Import-CSV  ".\csv\fim_classes.csv" -Encoding UTF8
$idinput = Import-CSV  ".\csv\_CUSTOM_STAFF_ID.csv" -Encoding UTF8

write-host
write-host "### Starting Staff Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$UserPath = "OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$ITPath = "OU=it,OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$TeacherPath = "OU=teaching,OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$NonTeacherPath = "OU=nonteaching,OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$ReliefTeacherPath = "OU=relief,OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$OtherPath = "OU=other,OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$TutorPath = "OU=tutors,OU=staff,OU=UserAccounts,DC=example,DC=com,DC=au"
$DisablePath = "OU=staff,OU=users,OU=disabled,DC=example,DC=com,DC=au"
# Get membership for group Membership Tests
$VillanovaGroups = Get-ADGroup -Filter * -SearchBase "OU=UserGroups,DC=example,DC=com,DC=au"
$TestStaff = Get-ADGroupMember -Identity "CN=Staff,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$TestCanonStaff = Get-ADGroupMember -Identity "CN=Printer-Canon-Staff,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
#$TestAllStaff = Get-ADGroupMember -Identity "CN=allstaff,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$TestTeachers = Get-ADGroupMember -Identity "CN=S-G_Teachers,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$TestMoodleTeachers = Get-ADGroupMember -Identity "CN=MoodleTeacher,OU=RoleAssignment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au"
$TestMapTeachers = Get-ADGroupMember -Identity "CN=Map-Teachers,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
#$OwncloudGroup = Get-ADGroupMember -Identity "CN=owncloud_staff,OU=owncloud,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
#Year Levels From teaching class lists
$teaches5 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year5,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches6 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year6,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches7 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year7,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches8 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year8,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches9 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year9,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches10 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year10,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches11 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year11,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
$teaches12 = Get-ADGroupMember -Identity "CN=S-G_Teacher-Year12,OU=security,OU=UserGroups,DC=example,DC=com,DC=au"
#Year level teaching mail groups
$mail5 = Get-ADGroupMember -Identity "CN=Teachers - Year 5,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail6 = Get-ADGroupMember -Identity "CN=Teachers - Year 6,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail7 = Get-ADGroupMember -Identity "CN=Teachers - Year 7,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail8 = Get-ADGroupMember -Identity "CN=Teachers - Year 8,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail9 = Get-ADGroupMember -Identity "CN=Teachers - Year 9,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail10 = Get-ADGroupMember -Identity "CN=Teachers - Year 10,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail11 = Get-ADGroupMember -Identity "CN=Teachers - Year 11,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$mail12 = Get-ADGroupMember -Identity "CN=Teachers - Year 12,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
# Teacher Pastoral groups
$JuniorPastoral = Get-ADGroupMember -Identity "CN=Teachers - Junior Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$MiddlePastoral = Get-ADGroupMember -Identity "CN=Teachers - Middle Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"
$SeniorPastoral = Get-ADGroupMember -Identity "CN=Teachers - Senior Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au"

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

##############################################
### Create / Edit / Disable Staff accounts ###
##############################################

foreach($line in $input) {

    # LoginName is the Unique Identifier for Staff
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())

    # teacher code is only given to teachers
    $TeacherCode = (Get-Culture).TextInfo.ToLower($line.tch_code.Trim())

    # Check Termination Dates
    $Termination = $line.term_date.Trim()

    #############################
    ### Process Current Staff ###
    #############################

    If ($Termination.length -eq 0) {

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

        IF (($Position2 -contains "Music Tutor") -or ($Position2 -contains "Relief Teacher") -or ($Position2 -contains "Teacher Relief")) {
            $Position2 = (Get-Culture).TextInfo.ToLower($Position2)
            $Position2 = (Get-Culture).TextInfo.ToTitleCase($Position2)
            If ($Position2 -eq "Teacher Relief") {
                $Position = "Relief Teacher"
            }
            Else {
                $Position = $Position2
            }
        }

        # Replace Common Acronyms and name spellings
        $Position = $Position -replace "Ict", "ICT"
        $Position = $Position -replace "DistrICT", "District"
        $Position = $Position -replace "Aic", "AIC"
        $Position = $Position -replace "Rto", "RTO"
        $Position = $Position -replace "Sor", "SOR"
        $Position = $Position -replace "Cic", "CIC"
        $Position = $Position -replace "Qcmf", "QCMF"
        $Position = $Position -replace " Of ", " of "
        $Position = $Position -replace " To ", " to "
        $Position = $Position -replace " The ", " the "
        $Position = $Position -replace " And ", " and "
        $Position = $Position -replace " Le Program Leader", " LE Program Leader"
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

        $FullName =  "${PreferredName} ${Surname}"
        $DisplayName = $FullName
        $DisplayName = $DisplayName -replace "Peter Wieneke", "Fr. Peter Wieneke OSA"
        $DisplayName = $DisplayName -replace "Peter Morris", "Dr. Peter Morris"
        $DisplayName = $DisplayName -replace "Irene Lategan", "Dr. Irene Lategan"
        # Set remaining details
        $UserPrincipalName = "${LoginName}@example.com.au"
        $HomeDrive = "\\example.com.au\home\Staff\${LoginName}"
        $Telephone = $line.phone_w_text.Trim()
        If ($Telephone.length -le 1) {
            $Telephone = $null
        }
        $employeeNumber = (Get-Culture).TextInfo.ToLower($line.record_id.Trim())

        ######################################
        ### Create / Modify Staff Accounts ###
        ######################################

        # create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "mypassword" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
            Set-ADUser -Identity $LoginName -Description $Position -Office $Position -Title $Position
            write-host "${LoginName} created for ${FullName}"
        }

        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

        # set additional user details if the user exists
        If (($TestUser) -and (!($TestUser.Department -eq "IGNORE"))) {

            # Get user info
            $TestAccountName = $TestUser.SamAccountName
            $TestHome = $TestUser.homedirectory
            $TestPhone = $TestUser.OfficePhone
            $TestDescription = $TestUser.Description
            $TestOffice = $TestUser.Office
            $TestTitle = $TestUser.Title
            $TestCompany = $TestUser.Company
            $TestPath = ($TestUser.distinguishedname.Contains($UserPath))
            $TestTeacherPath = ($TestUser.distinguishedname.Contains($TeacherPath))
            $TestNonTeacherPath = ($TestUser.distinguishedname.Contains($NonTeacherPath))
            $TestTutorPath = ($TestUser.distinguishedname.Contains($TutorPath))
            $TestITPath = ($TestUser.distinguishedname.Contains($ITPath))
            $TestReliefTeacherPath = ($TestUser.distinguishedname.Contains($ReliefTeacherPath))
            $TestNumber = $TestUser.EmployeeNumber
            $TestID = $TestUser.EmployeeID

            # Check Name Information
            If ($TestUser.GivenName -cne $PreferredName) {
                write-host $TestUser.GivenName, "Changed Given Name to ${PreferredName}"
                Set-ADUser -Identity $TestAccountName -GivenName $PreferredName
            }
            If ($TestUser.Surname -cne $Surname) {
                write-host $TestUser.SurName, "Changed Surname to ${SurName}"
                Set-ADUser -Identity $TestAccountName -Surname $Surname
            }
            If (($TestUser.Name -cne $FullName)) {
                write-host $TestUser.Name, "Changed Object Name to: ${FullName}"
                Rename-ADObject -Identity $TestUser -NewName $FullName
            }
            If (($TestUser.DisplayName -cne $DisplayName)) {
                write-host $TestUser.DisplayName, "Changed Display Name to: ${DisplayName}"
                Set-ADUser -Identity $TestUser.SamAccountName -DisplayName $DisplayName
            }
            #If ($TestUser.CN -cne $FullName) {
            #    write-host $TestUser.CN, "Changed Common Name to: ${FullName}"
            #    Set-ADUser -Identity $TestUser.SamAccountName -DisplayName $FullName
            #}

            # Enable user if disabled
            If (!($TestUser.Enabled)) {
                Set-ADUser -Identity $TestAccountName -Enabled $true
                write-host "Enabling", $TestAccountName
            }

            # Set userprofile path if is doesn't match
            If (!($TestHome -eq $HomeDrive)) {
                Set-ADUser -Identity $TestAccountName -homedrive "H:" -homedirectory $HomeDrive
                write-host "updated ${TestAccountName} home profile directory to: ${HomeDrive}"
            }

            # create home folder if it doesn't exist
            if (!(Test-Path $HomeDrive)) {
                New-Item -ItemType Directory -Force -Path $HomeDrive
            }

            # Add Position if there is one
            if (!($Position -ceq $TestDescription) -and (!($Position.length -eq 0))) {
                write-host $TestAccountName, "setting position"
                write-host "-${Position}-"
                write-host "-${TestDescription}-"
                write-host
                Set-ADUser -Identity $TestAccountName -Description $Position
            }

            # Add Office title
            if (!("Villanova College" -ceq $TestOffice)) {
                write-host $TestAccountName, "setting Office"
                write-host "-${Position}-"
                write-host "-${TestDescription}-"
                write-host
                Set-ADUser -Identity $TestAccountName -Office "Villanova College"
            }

            # Add title
            If (!($Position -ceq $TestTitle) -and (!($Position.length -eq 0))) {
                Set-ADUser -Identity $TestAccountName -Title $Position
                write-host $TestUser.Name, "Missing Title"
                write-host $Position
                write-host
            }

            # Add Employee Number if there is one
            if (!($employeeNumber -ceq $TestNumber) -and (!($employeeNumber.length -eq 0))) {
                write-host "Setting employee Number (${employeeNumber}) for ${TestAccountName}"
                write-host
                Set-ADUser -Identity $TestAccountName -EmployeeNumber $employeeNumber
            }

            # Set Department to identify current staff
            If (!(($TestUser.Department) -ceq ("Staff"))) {
                write-host $TestUser.Name, "Setting Position:", $TestUser.Department
                Set-ADUser -Identity $TestAccountName -Department "Staff"
                write-host "Staff"
            }

            # Add Telephone number if there is one
            if ($Telephone -ne $TestPhone) {
                If ($Telephone -eq $null) {
                    if ($TestPhone -ne '690') {
                        Set-ADUser -Identity $TestAccountName -OfficePhone '690'
                        write-host $TestAccountName, "setting Telephone to Default (690)"
                        write-host
                    }
                }
                Else {
                    Set-ADUser -Identity $TestAccountName -OfficePhone $Telephone
                    write-host $TestAccountName, "setting Telephone to:", $Telephone
                    write-host
                }
            }

            # Move user to their default OU if not already there
            if ($TestUser.description -eq $null) {
                if (!($TestUser.distinguishedname.Contains($OtherPath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $OtherPath
                    write-host "Moving to staff\other OU: no description for ${LoginName}"
                write-host
                }
            }
            Elseif ($TestUser.distinguishedname.Contains($DisablePath)) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                write-host $TestAccountName "moved out of Disabled OU"
                write-host
            }
            ElseIf (($TestCompany  -ceq "Relief Teacher") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                write-host $TestAccountName "moved to Relief Teacher OU"
                write-host
            }
            ElseIf ($TestCompany -ceq "Teacher") {
                If ($TestUser.distinguishedname.Contains($ReliefTeacherPath)) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                    write-host $TestAccountName "moved to  Teacher OU from Relief Teachers"
                    write-host
                }
                ElseIf (!($TestUser.distinguishedname.Contains($TeacherPath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                    write-host $TestAccountName "moved to  Teacher OU"
                    write-host
                }
            }
            ElseIf (($TestCompany -ceq "Tutors") -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                write-host $TestAccountName "moved to Music Tutor OU"
                write-host
            }
            ElseIf (($TestPath -and (!($TestTeacherPath))) -and (!($TestNonTeacherPath)) -and (!($TestITPath)) -and (!($TestTutorPath)) -and (!($TestReliefTeacherPath))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $NonTeacherPath
                write-host $TestUser.Name "moved to non-teaching"
                write-host $TestUser.DistinguishedName
                write-host
            }

            # Set company for automatic mail group filtering
            if (($TestUser.distinguishedname.Contains($NonTeacherPath)) -and (!($TeacherCode))) {
                if ((!($TestCompany -ceq "Admin")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $TestAccountName -Company "Admin"
                    write-host $TestUser.Name "set company to Admin"
                }
            }
            if ($TestUser.distinguishedname.Contains($ITPath)) {
                if ((!($TestCompany -ceq "ICT")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $TestAccountName -Company "ICT"
                    write-host $TestUser.Name "set company to ICT"
                }
            }
            if ($TestUser.distinguishedname.Contains($TeacherPath)) {
                if ((!($TestCompany -ceq "Teacher")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $TestAccountName -Company "Teacher"
                    write-host $TestUser.Name "set company to Teacher"
                }
            }
            if ($TestUser.distinguishedname.Contains($ReliefTeacherPath)) {
                if ((!($TestCompany -ceq "Relief")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $TestAccountName -Company "Relief"
                    write-host $TestUser.Name "set company to Teacher"
                }
            }
            if ($TestUser.distinguishedname.Contains($TutorPath)) {
                if ((!($TestCompany -ceq "Tutors")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $TestAccountName -Company "Tutors"
                    write-host $TestUser.Name "set company to Tutors"
                }
            }

            # Check Group Membership
            if (!($TestStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "Staff" -Member $TestAccountName
                        write-host $TestAccountName "added Staff"
            }
            if (!($TestCanonStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "CN=Printer-Canon-Staff,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                        write-host $TestAccountName "added Printer-Canon-Staff"
            }

            foreach($line in $idinput) {
                $tmpName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())
                $tmpID = (Get-Culture).TextInfo.ToUpper($line.idcard_nfc.Trim())
                If ($TestAccountName -eq $tmpName) {
                    If ($TestUser) {
                        If (!($TestID -eq $tmpID)) {
                            if (($tmpID -eq '') -or ($tmpID -eq '#null!')) {
                                #write-host "No ID Found for ${LoginName}"
                                Set-ADUser -Identity $TestAccountName -EmployeeID $null
                            }
                            else {
                                write-host "Setting ID (${tmpID}) for ${LoginName}"
                                write-host
                                Set-ADUser -Identity $TestAccountName -EmployeeID $tmpID
                            }
                        }
                    }
                }
            }
        }

        ###################################################################
        ### Create / Edit Teacher Info for Staff with existing accounts ###
        ###################################################################

        If (($TeacherCode -ne $null) -and ($TeacherCode -ne '')) {

            # Set user to confirm details
            $TestUser = (Get-ADUser  -Filter { (SamAccountName -eq $LoginName) }  -Properties *)
            $TestAccountName = $TestUser.SamAccountName
            $Description = $TestUser.Description

            #Make sure Teachers have the correct Company
            if ((!($TestCompany -ceq "Teacher")) -and (!($Description.Contains("Relief Teacher")))) {
                Set-ADUser -Identity $TestAccountName -Company "Teacher"
                write-host "Changing Company for ${TestAccountName} to Teacher"
                write-host
                #refresh details again
                $TestUser = (Get-ADUser  -Filter { (SamAccountName -eq $LoginName) }  -Properties *)
                $TestAccountName = $TestUser.SamAccountName
                $Description = $TestUser.Description
            }

            If ($TestUser.Enabled) {

                # Move to Teacher OU if not already there
                if ($TestUser.distinguishedname.Contains($UserPath) -and (!($TestUser.distinguishedname.Contains($TeacherPath))) -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                    If ($Description.Contains("Relief Teacher") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                        Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                        write-host $TestAccountName "moved to Relief Teacher OU"
                    }
                    ElseIf (($Description.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                        Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                        write-host $TestAccountName "moved to Music Tutor OU"
                    }
                    ElseIf (!($TestUser.distinguishedname.Contains($TeacherPath)) -and (!($Description.Contains("Tutor")))) {
                        Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                        write-host $TestAccountName "moved to Teacher OU"
                    }
                }
                # Check Group Membership
                if (!($TestTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity "CN=S-G_Teachers,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    write-host $TestAccountName "ADDED to Teachers Group"
                }
                if (!($TestMoodleTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity "CN=MoodleTeacher,OU=RoleAssignment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    write-host $TestAccountName "ADDED to MoodleTeachers Group"
                }
                if (!($TestMapTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity "CN=Map-Teachers,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    write-host $TestAccountName "ADDED to Map-Teachers Group"
                }

                # Year year level teacher groups
                $classin5 = $false
                $classin6 = $false
                $classin7 = $false
                $classin8 = $false
                $classin9 = $false
                $classin10 = $false
                $classin11 = $false
                $classin12 = $false
                $classjuniorpastoral = $false
                $classmiddlepastoral = $false
                $classseniorpastoral = $false
                # Parse the class list to identify if the teacher is in a class
                #write-host
                #write-host "Checking ${LoginName} for classes"
                foreach($line in $classinput) {
                    $tmpteach = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())
                    $tmpyear = (Get-Culture).TextInfo.ToLower($line.year_grp.Trim())
                    $tmpsubtitle = (Get-Culture).TextInfo.ToLower($line.sub_long.Trim())
                    If ($LoginName -eq $tmpteach) {
                        If (($tmpyear -eq '5') -and (!($classin5))) {
                            $classin5 = $true
                            #write-host "Found Year 5 Class"
                            if ($tmpsubtitle -eq "Junior School Pastoral") {
                                $classjuniorpastoral = $true
                            }
                        }
                        ElseIf (($tmpyear -eq '6') -and (!($classin6))) {
                            $classin6 = $true
                            #write-host "Found Year 6 Class"
                            if ($tmpsubtitle -eq "Junior School Pastoral") {
                                $classjuniorpastoral = $true
                            }
                        }
                        ElseIf (($tmpyear -eq '7') -and (!($classin7))) {
                            $classin7 = $true
                            #write-host "Found Year 7 Class"
                            if ($tmpsubtitle -eq "Crane Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Goold Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Heavey Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Murray Pastoral") {
                                $classmiddlepastoral = $true
                            }
                        }
                        ElseIf (($tmpyear -eq '8') -and (!($classin8))) {
                            $classin8 = $true
                            #write-host "Found Year 8 Class"
                            if ($tmpsubtitle -eq "Crane Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Goold Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Heavey Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Murray Pastoral") {
                                $classmiddlepastoral = $true
                            }
                        }
                        ElseIf (($tmpyear -eq '9') -and (!($classin9))) {
                            $classin9 = $true
                            #write-host "Found Year 9 Class"
                            if ($tmpsubtitle -eq "Crane Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Goold Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Heavey Pastoral") {
                                $classmiddlepastoral = $true
                            }
                            if ($tmpsubtitle -eq "Murray Pastoral") {
                                $classmiddlepastoral = $true
                            }
                        }
                        ElseIf ($tmpyear -eq '10') {
                            $classin10 = $true
                            #write-host "Found Year 10 Class"
                            if ($tmpsubtitle -eq "Crane Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Goold Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Heavey Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Murray Pastoral") {
                                $classseniorpastoral = $true
                            }
                        }
                        ElseIf ($tmpyear -eq '11') {
                            $classin11 = $true
                            #write-host "Found Year 11 Class"
                            if ($tmpsubtitle -eq "Crane Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Goold Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Heavey Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Murray Pastoral") {
                                $classseniorpastoral = $true
                            }
                        }
                        ElseIf ($tmpyear -eq '12') {
                            $classin12 = $true
                            #write-host "Found Year 12 Class"
                            if ($tmpsubtitle -eq "Crane Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Goold Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Heavey Pastoral") {
                                $classseniorpastoral = $true
                            }
                            if ($tmpsubtitle -eq "Murray Pastoral") {
                                $classseniorpastoral = $true
                            }
                        }
                    }
                }

                #Part Time STAFF ??? HACK
                If ($TestAccountName -eq 'liddym') {
                    $classin5 = $true
                }

                # add teachers to year level teaching groups from classes
                # remove teachers from year level teaching groups if there are no classes found
                If ($classin5) {
                    #write-host "Found Year 5 Class"
                    Try{
                        #write-host "Adding Year 5 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year5,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 5 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 5,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year5,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 5,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin6) {
                    #write-host "Found Year 6 Class"
                    Try{
                        #write-host "Adding Year 6 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year6,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 6 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 6,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year6,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 6,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin7) {
                    #write-host "Found Year 7 Class"
                    Try{
                        #write-host "Adding Year 7 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year7,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 7 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 7,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year7,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 7,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin8) {
                    #write-host "Found Year 8 Class"
                    Try{
                        #write-host "Adding Year 8 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year8,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 8 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 8,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year8,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 8,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin9) {
                    #write-host "Found Year 9 Class"
                    Try{
                        #write-host "Adding Year 9 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year9,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 9 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 9,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year9,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 9,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin10) {
                    #write-host "Found Year 10 Class"
                    Try{
                        #write-host "Adding Year 10 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year10,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 10 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 10,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year10,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 10,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin11) {
                    #write-host "Found Year 11 Class"
                    Try{
                        #write-host "Adding Year 11 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year11,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 11 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 11,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year11,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 11,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classin12) {
                    #write-host "Found Year 12 Class"
                    Try{
                        #write-host "Adding Year 12 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=S-G_Teacher-Year12,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 12 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Year 12,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=S-G_Teacher-Year12,OU=security,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity "CN=Teachers - Year 12,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classjuniorpastoral) {
                    Try{
                        #write-host "Adding Junior Pastoral Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Junior Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=Teachers - Junior Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classmiddlepastoral) {
                    Try{
                        #write-host "Adding Middle Pastoral Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Middle Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=Teachers - Middle Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
                If ($classseniorpastoral) {
                    Try{
                        #write-host "Adding Senior Pastoral Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity "CN=Teachers - Senior Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity "CN=Teachers - Senior Pastoral,OU=distribution,OU=UserGroups,DC=example,DC=com,DC=au" -Member $TestAccountName -Confirm:$false
                }
            }
        }
    }

    ###################################
    ### Disable Staff who have left ###
    ###################################

    Else {
        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
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
            # Terminate Staff AFTER their Termination date
            ElseIf ($DATE -gt $Termination) {
                # Disable The account when we don't want to keep it
                If ($TestUser) {
                    Set-ADUser -Identity $TestAccountName -Enabled $false
                    write-host "DISABLING ACCOUNT ${$LoginName}"
                    write-host
                }
            }
        }
        Else {
            # Enforce Group and OU changes for disabled staff
            If ($TestUser) {
                # Move to disabled user OU if not already there
                if (!($TestUser.distinguishedname.Contains($DisablePath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                    write-host "Moving: ${TestAccountName} to Disabled Staff OU"
                    write-host
                }

                # Remove groups if they are a member of any additional groups
                If ($TestMembership) {
                    write-host "Removing groups for ${TestAccountName}"
                    write-host
                    #remove All Villanova  Groups
                    Foreach($GroupName In $VillanovaGroups) {
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
write-host "### Staff Creation Script Finished"
write-host
