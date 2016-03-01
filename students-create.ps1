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
$input = Import-CSV "C:\DATA\csv\fim_student.csv" -Encoding UTF8
$inputcount = (Import-CSV "C:\DATA\csv\fim_student.csv" -Encoding UTF8 | Measure-Object).Count
$idinput = Import-CSV "C:\DATA\csv\_CUSTOM_STUDENT_ID.csv" -Encoding UTF8

### Get Default Password From Secure String File
### http://www.adminarsenal.com/admin-arsenal-blog/secure-password-with-powershell-encrypting-credentials-part-1/
###
$userpass = cat "C:\DATA\DefaultPassword.txt" | convertto-securestring

write-host
write-host "### Starting Current Student Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for different user types
$DisablePath = "OU=student,OU=users,OU=disabled,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$5Path = "OU=year5,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$6Path = "OU=year6,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$7Path = "OU=year7,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$8Path = "OU=year8,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$9Path = "OU=year9,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$10Path = "OU=year10,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$11Path = "OU=year11,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$12Path = "OU=year12,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

# Security Group names for students
$StudentName = "CN=Students,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$GenericPrintCode = "CN=9001,OU=print,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$5Name = "CN=S-G_year5,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$6Name = "CN=S-G_year6,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$7Name = "CN=S-G_year7,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$8Name = "CN=S-G_year8,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$9Name = "CN=S-G_year9,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$10Name = "CN=S-G_year10,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$11Name = "CN=S-G_year11,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$12Name = "CN=S-G_year12,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$UserAdmin = "CN=Local-Users-Administrators,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$UserPower = "CN=Local-Users-Power_Users,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$UserRegular = "CN=Local-Users-Users,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$MoodleStudent = "CN=MoodleStudent,OU=RoleAssignment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$MoodleTechHelp = "CN=tech-help-students,OU=student,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

# Deny Access group is used to remove problem students from important services
$DenyAccessGroup = "CN=S-G-Deny-Access,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

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
$LogDate = "${YEAR}-${MONTH}-${DAY}"

#EMAIL SETTINGS
# specify who gets notified 
$tonotification = "it@vnc.qld.edu.au"
# specify where the notifications come from 
$fromnotification = "notifications@vnc.qld.edu.au"
# specify the SMTP server 
$smtpserver = "mail.vnc.qld.edu.au"
# message for created users
$emailsubject = "New AD User Created:"
# message for disabled users
$disableemailsubject = "Current AD User Disabled:"
$disableemailbody = "Current AD user disabled
This is an automated email that is sent when an existing user is disabled."

# Get membership for group Membership Tests
$VillanovaGroups = Get-ADGroup -Filter * -SearchBase "OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$StudentGroup = Get-ADGroupMember -Identity $StudentName
$TestPrintGroup = Get-ADGroupMember -Identity $GenericPrintCode
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
$MoodleTechHelpMembers = Get-ADGroupMember -Identity $MoodleTechHelp

# Deny Access Members
$DenyAccessGroupMembers = Get-ADGroupMember -Identity $DenyAccessGroup

write-host "### Completed importing groups"
write-host


################################################
### Create / Edit / Disable student accounts ###
################################################

# check log path
If (!(Test-Path "C:\DATA\log")) {
    mkdir "C:\DATA\log"
}

# set log file
$LogFile = "C:\DATA\log\student-${LogDate}.log"
$LogContents = @()
$tmpcount = 0
$lastprogress = $NULL

write-host "### Processing Current Student File..."
Write-Host

foreach($line in $input) {
    $progress = ((($tmpcount / $inputcount) * 100) -as [int]) -as [string]
    If (((((($tmpcount / $inputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
        Write-Host "Progress: ${progress}%"
    }
    $tmpcount = $tmpcount + 1
    $lastprogress = $progress

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

    If (($Termination.length -eq 0) -or ($DATE -le $Termination)) {

        ################################
        ### Configure User Variables ###
        ################################

        # Get Year level information for groups and home drive
        $YearGroup = $line.year_grp
        If ($YearGroup -eq "5") {
            $UserPath = $5Path
            $ClassGroup = $5Name
        }
        If ($YearGroup -eq "6") {
            $UserPath = $6Path
            $ClassGroup = $6Name
        }
        If ($YearGroup -eq "7") {
            $UserPath = $7Path
            $ClassGroup = $7Name
        }
        If ($YearGroup -eq "8") {
            $UserPath = $8Path
            $ClassGroup = $8Name
        }
        If ($YearGroup -eq "9") {
            $UserPath = $9Path
            $ClassGroup = $9Name
        }
        If ($YearGroup -eq "10") {
            $UserPath = $10Path
            $ClassGroup = $10Name
        }
        If ($YearGroup -eq "11") {
            $UserPath = $11Path
            $ClassGroup = $11Name
        }
        If ($YearGroup -eq "12") {
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
        If ((Get-Culture).TextInfo.ToUpper($line.given_name.Trim()) -eq $line.given_name.Trim()) {
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
        $Surname = $Surname -replace "O'g", "O'G"
        $Surname = $Surname -replace "O'k", "O'K"
        $Surname = $Surname -replace "O'n", "O'N"
        $Surname = $Surname -replace "O'r", "O'R"

        # Set remaining details
        $FullName =  "${PreferredName} ${Surname}"
        $AltFullName =  "${GivenName} ${Surname}"
        ### Office 365 change ###$UserPrincipalName = "${LoginName}@villanova.vnc.qld.edu.au"
        $UserPrincipalName = "${LoginName}@vnc.qld.edu.au"
        $Position = "Year ${YearGroup}"
        $HomeDrive = $null
        $JobTitle = "Student - ${YEAR}"
        
        $emailbody = "There has been a new AD user created on the network.

Full Name: ${FullName}
User Name: ${LoginName}
Year Level: ${Position}
Email Address: ${UserPrincipalName}

An office 365 account will be created shortly using their new email address.

###########################
This is an automated email.
###########################"

        ########################################
        ### Create / Modify Student Accounts ###
        ########################################

        # Create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            Try  {
                New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $False
                $LogContents += "New User ${LoginName} created for ${FullName}"
                Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
            }
            Catch {
                Try {
                    # Error's can occur when the name of a student matches and they are in the same grade.
                    New-ADUser -SamAccountName $LoginName -Name $AltFullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $False
                    # Possible duplicate name
                    $LogContents += "New User ${LoginName} created for ${AltFullName}"
                    Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
                }
                Catch {
                    $LogContents += "The User ${LoginName} already exists for ${FullName} we tried: ${AltFullName}"
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
        $TestHomeDir = $TestUser.HomeDirectory
        $TestEnabled = $TestUser.Enabled
        $TestTitle = $TestUser.Title
        $TestCompany = $TestUser.Company
        $TestOffice = $TestUser.Office
        $TestDescription = $TestUser.Description
        $TestDepartment = $TestUser.Department
        $TestNumber = $TestUser.employeeNumber
        $TestID = $TestUser.employeeID
        $TestMembership = $TestUser.MemberOf

        # Get office365 details
        $TestEmail = $TestUser.mail
        If ($TestEmail) {
            $TestEmail = $TestEmail.ToLower()
        }
        $TestPrincipal = $TestUser.UserPrincipalName

        # Remove groups if they are a member of any additional groups
        If ($DenyAccessGroupMembers.SamAccountName.contains($LoginName)) {
            If ($TestMembership) {
                write-host "Removing groups for ${TestAccountName}"
                write-host
                #remove All Villanova  Groups
                Foreach($GroupName In $TestMembership) {
                    write-host $GroupName
                    #Try {
                    #    Remove-ADGroupMember -Identity $GroupName -Member $TestAccountName -Confirm:$false
                    #}
                    #Catch {
                    #$LogContents += "Error Removing ${TestAccountName} from ${GroupName}"
                    #}
                }
            }
        }

        # set additional user details if the user exists
        If ($TestUser) {

            # Check that UPN is set to email. but only if an email exists
            If (($TestEmail) -and (!($TestEmail -ceq $TestPrincipal))) {
                Set-ADUser -Identity $TestDN -UserPrincipalName $TestEmail
                $LogContents += "UPN CHANGE: ${TestPrincipal} to ${TestEmail}"
                Write-Host "UPN CHANGE: ${TestPrincipal} to ${TestEmail}"
            }

            # Enable user if disabled
            If ((!($TestEnabled)) -and (!($TestDescription -eq "disable"))) {
                Set-ADUser -Identity $LoginName -Enabled $true
                $LogContents += "Enabling ${TestAccountName}"
            }
            # Disable if description contains disable
            ElseIf (($TestEnabled) -and ($TestDescription -eq "disable")) {
                Set-ADUser -Identity $LoginName -Enabled $false
                $LogContents += "Disabling ${TestAccountName}"
            }

            # Move user to the default OU for their year level if not there
            If (($TestEnabled) -and (!($TestDN.Contains($UserPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                $LogContents += "Taking ${TestAccountName} From: ${TestDN}"
                $LogContents += "Moving ${TestAccountName} To: ${UserPath}"
            }

            If ((!($TestDescription -eq $UserCode)) -and (!($TestDescription -eq "disable"))) {
                Set-ADUser -Identity $LoginName -Description $UserCode
                write-host $FullName, "changing description from ${TestDescription} to ${UserCode}"
                write-host
            }

            # Remove HomeDirectory for students who are not in year 5 or 6
            If ((!($YearGroup -eq "5")) -and (!($YearGroup -eq "6")) -and $TestHomeDir) {
                Set-ADUser -Identity $LoginName -HomeDirectory $null
                Write-Host "${LoginName} doesn't need a home folder"
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
            If (($TestName -cne $FullName) -and ($TestName -cne $AltFullName)) {
                Try {
                    Rename-ADObject -Identity $TestDN -NewName $FullName
                    write-host "${TestAccountName} Changed Object Name to: ${FullName}"
                }
                Catch {
                    Rename-ADObject -Identity $TestDN -NewName $AltFullName
                    write-host "${TestAccountName} Changed Object Name to: ${AltFullName}"
                }
            }
            If (($TestDisplayName -cne $FullName)) {
                Set-ADUser -Identity $LoginName -DisplayName $FullName
                write-host "${TestAccountName} Changed Display Name to: ${FullName}"
            }

            # Set company for automatic mail group filtering
            If ($UserPath -eq $5Path) {
                If (!($TestCompany -ceq "year5")) {
                    Set-ADUser -Identity $LoginName -Company "year5"
                    write-host $TestName "set company to year5"
                }
            }
            If ($UserPath -eq $6Path) {
                If (!($TestCompany -ceq "year6")) {
                    Set-ADUser -Identity $LoginName -Company "year6"
                    write-host $TestName "set company to year6"
                }
            }
            If ($UserPath -eq $7Path) {
                If (!($TestCompany -ceq "year7")) {
                    Set-ADUser -Identity $LoginName -Company "year7"
                    write-host $TestName "set company to year7"
                }
            }
            If ($UserPath -eq $8Path) {
                If (!($TestCompany -ceq "year8")) {
                    Set-ADUser -Identity $LoginName -Company "year8"
                    write-host $TestName "set company to year8"
                }
            }
            If ($UserPath -eq $9Path) {
                If (!($TestCompany -ceq "year9")) {
                    Set-ADUser -Identity $LoginName -Company "year9"
                    write-host $TestName "set company to year9"
                }
            }
            If ($UserPath -eq $10Path) {
                If (!($TestCompany -ceq "year10")) {
                    Set-ADUser -Identity $LoginName -Company "year10"
                    write-host $TestName "set company to year10"
                }
            }
            If ($UserPath -eq $11Path) {
                If (!($TestCompany -ceq "year11")) {
                    Set-ADUser -Identity $LoginName -Company "year11"
                    write-host $TestName "set company to year11"
                }
            }
            If ($UserPath -eq $12Path) {
                If (!($TestCompany -ceq "year12")) {
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
            Else {
                Set-ADUser -Identity $LoginName -Office $Position
                write-host $LoginName, "Office missing; set to ${Position}"
            }

            # set Office to current year level
            If ($YearGroup.length -eq 1) {
                If ($YearGroup -ne $test1) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host $LoginName, "year level change from ${TestOffice} to ${Position}"
                }
                ElseIf ($TestOffice -eq "Future Year ${YearGroup}") {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host $LoginName, "year level change from ${TestOffice} to ${Position}"
                }
            }
            ElseIf ($YearGroup.length -eq 2) {
                If ($YearGroup -ne $test2) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host $LoginName, "year level change from ${TestOffice} to ${Position}"
                }
                ElseIf ($TestOffice -eq "Future Year ${YearGroup}") {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host $LoginName, "year level change from ${TestOffice} to ${Position}"
                }
            }

            # Set Department to identify current students
            If (!(($TestDepartment) -ceq ("Student"))) {
                Set-ADUser -Identity $LoginName -Department "Student"
                write-host "${TestName} Setting Position from ${TestDepartment} to Student"
            }

            # Check Group Membership
            If (!($StudentGroup.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity "Students" -Member $LoginName
                write-host $LoginName "added Students Group"
            }
            If (!($LocalUser.SamAccountName.contains($LoginName))) {
                write-host $TestName, "Add to local user group for domain workstations"
                Add-ADGroupMember -Identity $UserRegular -Member $LoginName
            }
            # $MoodleStudentMembers
            If (!($MoodleStudentMembers.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity $MoodleStudent -Member $LoginName
                write-host $LoginName "added MoodleStudent Group"
            }
            # $MoodleTechHelpMembers
            If (!($MoodleTechHelpMembers.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity $MoodleTechHelp -Member $LoginName
                write-host $LoginName "added MoodleTechHelp Group"
            }
            # $TestPrintGroup
            If (!($TestPrintGroup.name.contains($TestUser.name))) {
                Add-ADGroupMember -Identity $GenericPrintCode -Member $TestAccountName
                write-host $TestAccountName "added default printer group ${GenericPrintCode}"
            }

            # Remove groups for other grades and add the correct grade
            If ($YearGroup -eq "5") {
                # Add Correct Year Level
                If (!($5Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $5Name -Member $TestAccountName
                    write-host $LoginName "added 5"
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "6") {
                If ($5Group.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($6Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $6Name -Member $TestAccountName
                    write-host $LoginName "added 6"
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "7") {
                If ($5Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($7Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $7Name -Member $TestAccountName
                    write-host $LoginName "added 7"
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "8") {
                If ($5Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($8Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $8Name -Member $TestAccountName
                    write-host $LoginName "added 8"
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "9") {
                If ($5Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($9Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $9Name -Member $TestAccountName
                    write-host $LoginName "added 9"
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "10") {
                If ($5Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($10Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $10Name -Member $TestAccountName
                    write-host $LoginName "added 10"
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "11") {
                If ($5Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($11Group.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $11Name -Member $TestAccountName
                    write-host $LoginName "added 11"
                }
                If ($12Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $12Name -Member $TestAccountName -Confirm:$false
                }
            }
            If ($YearGroup -eq "12") {
                If ($5Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $5Name -Member $TestAccountName -Confirm:$false
                }
                If ($6Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $6Name -Member $TestAccountName -Confirm:$false
                }
                If ($7Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $7Name -Member $TestAccountName -Confirm:$false
                }
                If ($8Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $8Name -Member $TestAccountName -Confirm:$false
                }
                If ($9Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $9Name -Member $TestAccountName -Confirm:$false
                }
                If ($10Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $10Name -Member $TestAccountName -Confirm:$false
                }
                If ($11Group.SamAccountName.contains($LoginName)) {
                    Remove-ADGroupMember -Identity $11Name -Member $TestAccountName -Confirm:$false
                }
                # Add Correct Year Level
                If (!($12Group.SamAccountName.contains($LoginName))) {
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
            $tmpID = (Get-Culture).TextInfo.ToUpper($line.idcard_dec.Trim())
            $tmpNum = (Get-Culture).TextInfo.ToUpper($line.idcard_nfc.Trim())
            If ($TestAccountName -eq $tmpName) {
                If ($TestUser) {
                    # warn about mismatched fields
                    If ((!($tmpID.length -eq 0)) -and ($tmpNum.length -eq 0)) {
                        write-host "missing hex for ${TestAccountName}"
                    }
                    If (($tmpID.length -eq 0) -and (!($tmpNum.length -eq 0))) {
                        write-host "missing decimal for ${TestAccountName}"
                    }
                    # Add Employee ID if there is one
                    If ((!($TestID -ceq $tmpID)) -and (!($tmpID.length -eq 0))) {
                        Set-ADUser -Identity $LoginName -EmployeeID $tmpID
                        write-host "Setting decimal employeeID (${tmpID}) for ${TestAccountName}"
                        $LogContents += "Setting decimal employeeID (${tmpID}) for ${LoginName}"
                    }
                    # Add Employee Number if there is one
                    If (!($TestNumber -ceq $tmpNum) -and (!($tmpNum.length -eq 0))) {
                        Set-ADUser -Identity $LoginName -EmployeeNumber $tmpNum
                        write-host "Setting Hex employeeNumber (${tmpNum}) for ${TestAccountName}"
                        $LogContents += "Setting Hex employeeNumber (${tmpNum}) for ${LoginName}"
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
        $TestTitle = $TestUser.Title
        $TestCompany = $TestUser.Company
        $TestDepartment = $TestUser.Department
        $TestEnabled = $TestUser.Enabled
        $TestAccountName = $TestUser.SamAccountName
        $TestMembership = $TestUser.MemberOf

        # Disable users with a termination date if they are still enabled
        If ($TestEnabled) {

            # Don't disable users we want to keep
            If ($TestDescription -eq "keep") {
                If (!($LoginName -eq '10961')) {
                    $LogContents += "${LoginName} Keeping terminated user"
                }
            }
            # Terminate Students AFTER their Termination date
            ElseIf ($DATE -gt $Termination) {
                # Disable The account when we don't want to keep it
                If ($TestUser) {
                    Set-ADUser -Identity $LoginName -Enabled $false
                    $LogContents += "DISABLING ACCOUNT ${TestAccountName}"
                    $LogContents += "Now: ${DATE}"
                    $LogContents += "DOL: ${Termination}"
                    Send-MailMessage -From $fromnotification -Subject "${disableemailsubject} ${LoginName}" -To $tonotification -Body $disableemailbody -SmtpServer $smtpserver
                }
            }
        }
        ElseIf ($TestUser) {
            # Move to disabled user OU if not already there
            If (!($TestDN.Contains($DisablePath))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                $LogContents += "Moving ${LoginName}: ${TestAccountName} to Disabled Student OU"
            }
            else {
                # Set Department to "Disabled" to help identify current students
                If (!(($TestDepartment) -ceq ("Disabled"))) {
                    Set-ADUser -Identity $LoginName -Department "Disabled"
                    write-host "${LoginName} Setting Position from ${TestDepartment} to Disabled"
                }
                # Set Company to "Disabled" to help identify current students
                If (!($TestCompany -ceq "Disabled")) {
                    Set-ADUser -Identity $LoginName -Company "Disabled"
                    write-host "${LoginName} set company to Disabled"
                }
                # Set Title to "Disabled" to help identify current students
                If (!($TestTitle -ceq "Disabled")) {
                    Set-ADUser -Identity $LoginName -Title "Disabled"
                    write-host "${LoginName} Title change to: Disabled"
                }
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
                        $LogContents += "Error Removing ${TestAccountName} from ${GroupName}"
                    }
                }
            }
        }
    }
}

write-host
write-host "### Current Student Creation Script Finished"
write-host

#######################
### FUTURE STUDENTS ###
#######################

write-host "### Starting Future Student Creation Script"
write-host
write-host "### Processing Future Student File..."
Write-Host

# Future students csv
$enrolledinput = Import-CSV "C:\DATA\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8
$enrolledcount = (Import-CSV  "C:\DATA\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8 | Measure-Object).Count

# OU paths for different user types
$FuturePath = "OU=future,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

$tmpcount = 0
$lastprogress = $NULL

foreach($line in $enrolledinput) {
    $progress = ((($tmpcount / $enrolledcount) * 100) -as [int]) -as [string]
    If (((((($tmpcount / $enrolledcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
        Write-Host "Progress: ${progress}%"
    }
    $tmpcount = $tmpcount + 1
    $lastprogress = $progress

    # UserCode is the Unique Identifier for Students
    $UserCode = (Get-Culture).TextInfo.ToUpper($line.stud_code.Trim())
    $YearGroup = $line.entry_year_grp

    # Correct usercode if opened in Excel
    If ($UserCode.Length -ne 5) {
            $UserCode = "0${UserCode}"
    }

    # Set Login Name
    $LoginName = $UserCode

    # Use future student OU by default
    $UserPath = $FuturePath

    ################################
    ### Configure User Variables ###
    ################################

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
    If ((Get-Culture).TextInfo.ToUpper($line.given_name.Trim()) -eq $line.given_name.Trim()) {
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
    $Surname = $Surname -replace "O'g", "O'G"
    $Surname = $Surname -replace "O'k", "O'K"
    $Surname = $Surname -replace "O'n", "O'N"
    $Surname = $Surname -replace "O'r", "O'R"

    # Set remaining details
    $FullName =  "${PreferredName} ${Surname}"
    $UserPrincipalName = "${LoginName}@vnc.qld.edu.au"
    
    # Home Folders are only for younger grades
    $HomeDrive = $null

    # Position and title

    $JobTitle = "Future Student"
    $Position = "Future Year ${YearGroup}"
    
        $emailbody = "New AD user created
This is an automated email that is sent when a new user is created.
${FullName}
${LoginName}
${Position}"

    ########################################
    ### Create / Modify Student Accounts ###
    ########################################

    # Create basic user if you can't find one
    If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
        Try  {
            New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $False
            $LogContents += "${LoginName} created for ${FullName}"
            Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
        }
        Catch {
            # ignore future student error
            #$LogContents += "The username ${LoginName} for Future student ${FullName} already exists"
        }
    }
    Else {
        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

        # Get User Details
        $TestName = $TestUser.Name
        $TestGiven = $TestUser.GivenName
        $TestSurname = $TestUser.SurName
        $TestDisplayName = $TestUser.DisplayName
        $TestDN = $TestUser.distinguishedname
        $TestAccountName = $TestUser.SamAccountName
        $TestHomeDir = $TestUser.HomeDirectory
        $TestEnabled = $TestUser.Enabled
        $TestTitle = $TestUser.Title
        $TestCompany = $TestUser.Company
        $TestOffice = $TestUser.Office
        $TestDescription = $TestUser.Description
        $TestDepartment = $TestUser.Department
        $TestNumber = $TestUser.employeeNumber
        $TestID = $TestUser.employeeID

        # Get office365 details
        $TestEmail = $TestUser.mail
        If ($TestEmail) {
            $TestEmail = $TestEmail.ToLower()
        }
        $TestPrincipal = $TestUser.UserPrincipalName

        # set additional user details if the user exists
        # (If a student has previously left and been renrolled they will be disabled until their return.)
        If (($TestUser) -and ($TestDN -contains $UserPath)) {

            # Check that UPN is set to email. but only if an email exists
            If (($TestEmail) -and (!($TestEmail -ceq $TestPrincipal))) {
                Set-ADUser -Identity $TestDN -UserPrincipalName $TestEmail
                $LogContents += "UPN CHANGE: ${TestPrincipal} to ${TestEmail}"
            }

            # set description to stud_code
            If (!($TestDescription -eq $UserCode)) {
                Set-ADUser -Identity $LoginName -Description $UserCode
                write-host "${TestAccountName} changing description from ${TestDescription} to ${UserCode}"
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
                Rename-ADObject -Identity $TestDN -NewName $FullName
                write-host "${TestAccountName} Changed Object Name to: ${FullName}"
            }
            If (($TestDisplayName -cne $FullName)) {
                Set-ADUser -Identity $LoginName -DisplayName $FullName
                write-host "${TestAccountName} Changed Display Name to: ${FullName}"
            }

            # Set Year Level and Title
            If (($TestTitle) -eq $null) {
                Set-ADUser -Identity $LoginName -Title $JobTitle
                write-host "${TestAccountName} Title set: ${JobTitle}"
            }
            ElseIf (!($TestTitle).contains($JobTitle)) {
                Set-ADUser -Identity $LoginName -Title $JobTitle
                write-host "${TestAccountName} Title change to: ${JobTitle}"
            }

            # Get the year level of the current office string
            If ($TestOffice) {
                $test1 = $TestOffice.Substring($TestOffice.length-1,1)
                $test2 = $TestOffice.Substring($TestOffice.length-2,2)
            }
            Else {
                Set-ADUser -Identity $LoginName -Office $Position
                write-host "${TestAccountName} Office missing; set to ${Position}"
            }

            # set Office to current year level
            If ($YearGroup.length -eq 1) {
                If ($YearGroup -ne $test1) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host "${TestAccountName} year level change from ${TestOffice} to ${Position}"
                }
            }
            ElseIf ($YearGroup.length -eq 2) {
                If ($YearGroup -ne $test2) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host "${TestAccountName} year level change from ${TestOffice} to ${Position}"
                }
            }

            # Set Department to identify current students
            If (!(($TestDepartment) -ceq ("Future"))) {
                Set-ADUser -Identity $LoginName -Department "Future"
                write-host "${TestAccountName} Setting Position  to 'Future'"
            }

            # Add Employee Number if there is one
            If (!($LoginName -ceq $TestNumber)) {
                Set-ADUser -Identity $LoginName -EmployeeNumber $LoginName
                write-host "${TestAccountName} Setting employee Number"
                write-host
            }
        }
        #Else {
        #    write-host "missing or ignoring ${FullName}: ${LoginName}"
        #    write-host
        #}
    }
}

Write-Host
Write-Host "### Future Student file finished"
Write-Host

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
    Write-Host "No Important Changes were logged"
}

Write-Host
Write-Host "DONE"
