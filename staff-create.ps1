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
$input = Import-CSV  ".\csv\fim_staff.csv" -Encoding UTF8
$classinput = Import-CSV  ".\csv\fim_classes.csv" -Encoding UTF8
$idinput = Import-CSV  ".\csv\_CUSTOM_STAFF_ID.csv" -Encoding UTF8

write-host
write-host "### Starting Staff Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$UserPath = "OU=example,DC=qld,DC=edu,DC=au"
$ITPath = "OU=example,DC=qld,DC=edu,DC=au"
$TeacherPath = "OU=example,DC=qld,DC=edu,DC=au"
$NonTeacherPath = "OU=example,DC=qld,DC=edu,DC=au"
$ReliefTeacherPath = "OU=example,DC=qld,DC=edu,DC=au"
$TutorPath = "OU=example,DC=qld,DC=edu,DC=au"
$DisablePath = "OU=example,DC=qld,DC=edu,DC=au"
# Get membership for group Membership Tests
$TestStaff = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$TestAllStaff = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$TestTeachers = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$TestMoodleTeachers = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
#$OwncloudGroup = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
#Year Levels From teaching class lists
$teaches5 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches6 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches7 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches8 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches9 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches10 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches11 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$teaches12 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
#Year level teaching mail groups
$mail5 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail6 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail7 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail8 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail9 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail10 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail11 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$mail12 = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
# Teacher Pastoral groups
$JuniorPastoral = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$MiddlePastoral = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
$SeniorPastoral = Get-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au"
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
        $UserPrincipalName = "${LoginName}@villanova.vnc.qld.edu.au"
        $HomeDrive = "\\fileserver\home\Staff\${LoginName}"
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
            New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "mypasswordhere" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
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
                write-host "no description for ${LoginName}"
            }
            Elseif ($TestUser.distinguishedname.Contains($DisablePath)) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                write-host $TestAccountName "moved out of Disabled OU"
            }
            ElseIf (($TestUser.Description.Contains("Relief Teacher")) -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                write-host $TestAccountName "moved to Relief Teacher OU"
            }
            ElseIf (($TestUser.Description.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                write-host $TestAccountName "moved to Music Tutor OU"
            }
            ElseIf (($TestPath -and (!($TestTeacherPath))) -and (!($TestNonTeacherPath)) -and (!($TestITPath)) -and (!($TestTutorPath)) -and (!($TestReliefTeacherPath))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $NonTeacherPath
                write-host $TestUser.Name "moved to non-teaching"
                write-host $TestUser.DistinguishedName
                write-host
            }

            # Set company for automatic mail group filtering
            if ($TestUser.distinguishedname.Contains($NonTeacherPath)) {
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
            if (!($TestAllStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "All Staff" -Member $TestAccountName
                        write-host $TestAccountName "added allstaff"
            }
            #if (!($OwncloudGroup.name.contains($TestUser.name))) {
            #            Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
            #            write-host $TestAccountName "added owncloud"
            #}
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
                    Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    write-host $TestAccountName "ADDED to Teachers Group"
                }
                if (!($TestMoodleTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    write-host $TestAccountName "ADDED to MoodleTeachers Group"
                }

                # Chear year level teacher groups
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
                # add teachers to year level teaching groups from classes
                If ($classin5) {
                    if (!($teaches5.name.contains($TestUser.name))) {
                        write-host "Found Year 5 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail5.name.contains($TestUser.name))) {
                        write-host "Found Year 5 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin6) {
                    if (!($teaches6.name.contains($TestUser.name))) {
                        write-host "Found Year 6 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail6.name.contains($TestUser.name))) {
                        write-host "Found Year 6 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin7) {
                    if (!($teaches7.name.contains($TestUser.name))) {
                        write-host "Found Year 7 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail7.name.contains($TestUser.name))) {
                        write-host "Found Year 7 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin8) {
                    if (!($teaches8.name.contains($TestUser.name))) {
                        write-host "Found Year 8 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail8.name.contains($TestUser.name))) {
                        write-host "Found Year 8 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin9) {
                    if (!($teaches9.name.contains($TestUser.name))) {
                        write-host "Found Year 9 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail9.name.contains($TestUser.name))) {
                        write-host "Found Year 9 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin10) {
                    if (!($teaches10.name.contains($TestUser.name))) {
                        write-host "Found Year 10 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail10.name.contains($TestUser.name))) {
                        write-host "Found Year 10 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin11) {
                    if (!($teaches11.name.contains($TestUser.name))) {
                        write-host "Found Year 11 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail11.name.contains($TestUser.name))) {
                        write-host "Found Year 11 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classin12) {
                    if (!($teaches12.name.contains($TestUser.name))) {
                        write-host "Found Year 12 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                    if (!($mail12.name.contains($TestUser.name))) {
                        write-host "Found Year 12 Class"
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classjuniorpastoral) {
                    if (!($JuniorPastoral.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classmiddlepastoral) {
                    if (!($MiddlePastoral.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                If ($classseniorpastoral) {
                    if (!($SeniorPastoral.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName
                    }
                }
                # remove teachers from year level teaching groups if there are no classes found
                If (!($classin5)) {
                    #write-host "REMOVED Year 5 Class"
                    if ($teaches5.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail5.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin6)) {
                    #write-host "REMOVED Year 6 Class"
                    if ($teaches6.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail6.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin7)) {
                    #write-host "REMOVED Year 7 Class"
                    if ($teaches7.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail7.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin8)) {
                    #write-host "REMOVED Year 8 Class"
                    if ($teaches8.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail8.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin9)) {
                    #write-host "REMOVED Year 9 Class"
                    if ($teaches9.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail9.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin10)) {
                    #write-host "REMOVED Year 10 Class"
                    if ($teaches10.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail10.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin11)) {
                    #write-host "REMOVED Year 11 Class"
                    if ($teaches11.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail11.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classin12)) {
                    #write-host "REMOVED Year 12 Class"
                    if ($teaches12.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                    if ($mail12.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classjuniorpastoral)) {
                    if ($JuniorPastoral.name.contains($TestUser.name)) {
                       Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classmiddlepastoral)) {
                    if ($MiddlePastoral.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
                If (!($classseniorpastoral)) {
                    if ($SeniorPastoral.name.contains($TestUser.name)) {
                        Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    }
                }
            }
        }
    }

    ###################################
    ### Disable Staff who have left ###
    ###################################
    
    Else {

        # Disable users with a termination date if they are still enabled
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }) {

            # Set user to confirm details
            $TestUser =  (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
            $TestAccountName = $TestUser.SamAccountName
            
            If ($TestUser.Description -eq "keep") {
                write-host "${LoginName} Keeping terminated user"
            }
            # Terminate Staff AFTER their Termination date
            ElseIf ($DATE -gt $Termination) {
                write-host "DISABLING ACCOUNT", $TestAccountName
                write-host $DATE
                write-host $Termination
                write-host

                if (!($TestUser.distinguishedname.Contains($DisablePath))) {
                    # Move to disabled user OU if not already there
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                }
            
                # Disable The account
                Set-ADUser -Identity $TestAccountName -Enabled $false
            }
            Else {
                write-host "Not Final Leaving Date", $TestUser.name
                write-host $DATE
                write-host $Termination
                write-host
            }
        }

        # Remove Disabled Staff from Groups
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "False")) }) {

            # Set user to confirm details
            $TestUser =  (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
            $TestAccountName = $TestUser.SamAccountName

            If ($TestUser) {
                #remove Staff Groups
                if (($TestStaff.name.contains($TestUser.name))) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed Staff"
                }
                if (($TestAllStaff.name.contains($TestUser.name))) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed allstaff"
                }
                if (($TestTeachers.name.contains($TestUser.name))) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed sg Teachers"
                }
                if (($TestMoodleTeachers.name.contains($TestUser.name))) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed Moodle Teachers"
                }
                #Year Levels From teaching class lists
                if ($teaches5.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 5"
                }
                if ($teaches6.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 6"
                }
                if ($teaches7.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 7"
                }
                if ($teaches8.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 8"
                }
                if ($teaches9.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 9"
                }
                if ($teaches10.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 10"
                }
                if ($teaches11.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed  sg teacher 11"
                }
                if ($teaches12.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed sg teacher 12"
                }
                #Year level teaching mail groups
                if ($mail5.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 5"
                }
                if ($mail6.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 6"
                }
                if ($mail7.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 7"
                }
                if ($mail8.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 8"
                }
                if ($mail9.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 9"
                }
                if ($mail10.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 10"
                }
                if ($mail11.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 11"
                }
                if ($mail12.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed mail teacher 12"
                }
                # Teacher Pastoral groups
                if ($JuniorPastoral.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed junior pastoral"
                }
                if ($MiddlePastoral.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed middle pastoral"
                }
                if ($SeniorPastoral.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "OU=example,DC=qld,DC=edu,DC=au" -Member $TestAccountName -Confirm:$false
                    write-host $TestAccountName "removed senior pastoral"
                }
            }
        }
    }
}

write-host
write-host "### Staff Creation Script Finished"
write-host
