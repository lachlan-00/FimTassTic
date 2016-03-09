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
$Studmail = "CN=Students - All,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$5mail = "CN=Students - Year 5,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$6mail = "CN=Students - Year 6,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$7mail = "CN=Students - Year 7,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$8mail = "CN=Students - Year 8,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$9mail = "CN=Students - Year 9,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$10mail = "CN=Students - Year 10,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$11mail = "CN=Students - Year 11,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$12mail = "CN=Students - Year 12,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$UserAdmin = "CN=Local-Users-Administrators,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$UserPower = "CN=Local-Users-Power_Users,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$UserRegular = "CN=Local-Users-Users,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$MoodleStudent = "CN=MoodleStudent,OU=RoleAssignment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$MoodleTechHelp = "CN=tech-help-students,OU=student,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

# Deny Access group is used to remove problem students from important services
$DenyAccessGroup = "CN=S-G_Deny-Access,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$DomainUsersGroup = "CN=Domain Users,CN=Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

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
$ccnotification = "jlane@vnc.qld.edu.au"
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
$StudmailGroup = Get-ADGroupMember -Identity $Studmail
$5mailGroup = Get-ADGroupMember -Identity $5mail
$6mailGroup = Get-ADGroupMember -Identity $6mail
$7mailGroup = Get-ADGroupMember -Identity $7mail
$8mailGroup = Get-ADGroupMember -Identity $8mail
$9mailGroup = Get-ADGroupMember -Identity  $9mail
$10mailGroup = Get-ADGroupMember -Identity $10mail
$11mailGroup = Get-ADGroupMember -Identity $11mail
$12mailGroup = Get-ADGroupMember -Identity $12mail
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
            $StudCompany = "year5"
        }
        If ($YearGroup -eq "6") {
            $UserPath = $6Path
            $ClassGroup = $6Name
            $StudCompany = "year6"
        }
        If ($YearGroup -eq "7") {
            $UserPath = $7Path
            $ClassGroup = $7Name
            $StudCompany = "year7"
        }
        If ($YearGroup -eq "8") {
            $UserPath = $8Path
            $ClassGroup = $8Name
            $StudCompany = "year8"
        }
        If ($YearGroup -eq "9") {
            $UserPath = $9Path
            $ClassGroup = $9Name
            $StudCompany = "year9"
        }
        If ($YearGroup -eq "10") {
            $UserPath = $10Path
            $ClassGroup = $10Name
            $StudCompany = "year10"
        }
        If ($YearGroup -eq "11") {
            $UserPath = $11Path
            $ClassGroup = $11Name
            $StudCompany = "year11"
        }
        If ($YearGroup -eq "12") {
            $UserPath = $12Path
            $ClassGroup = $12Name
            $StudCompany = "year12"
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

        # Deny Access email body
        $DenyBody = "The User ${LoginName}/${FullName} has been removed from Villanova Groups due to acceptable use policy violations.

This affects access to Villanova services such as:
 * Network File Access
 * Villanova Website Access (Excluding Moodle's Tech Help page)
 * Villanova WiFi Access (BYOD & Student)
 * Villanova Google Account

###########################
This is an automated email.
###########################"
        # New User email body
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


        # Check for Deny-Access members
        If (!($DenyAccessGroupMembers)) {
            #Empty Group
        }
        # Remove User from groups if they are a member of Deny Access
        ElseIf ($DenyAccessGroupMembers.SamAccountName.contains($LoginName)) {
            $RemovalCheck = $False
            If ($TestMembership) {
                # Remove All Villanova  Groups
                Foreach($GroupName In $TestMembership) {
                    If (!(($GroupName -eq $MoodleTechHelp) -or ($GroupName -eq $DenyAccessGroup) -or ($GroupName -eq $DomainUsersGroup))) {
                        Try {
                            Remove-ADGroupMember -Identity $GroupName -Member $TestAccountName -Confirm:$false
                            If ($GroupName -eq $StudentName) {
                                $RemovalCheck = $true
                            }
                        }
                        Catch {
                        $LogContents += "Error Removing ${TestAccountName} from ${GroupName}"
                        }
                    }
                }
                If ($RemovalCheck) {
                    $LogContents += $DenyBody
                    Send-MailMessage -From $fromnotification -Subject "DENY ACCESS ${LoginName}" -To $tonotification -Cc $ccnotification -Body $DenyBody -SmtpServer $smtpserver
                }
                # Change Details to block WiFi and email groups
                If (!(($TestDepartment) -ceq ("DENYACCESS"))) {
                    Set-ADUser -Identity $LoginName -Department "DENYACCESS"
                    write-host "${LoginName} Setting Position from ${TestDepartment} to DENYACCESS"
                }
                If (!($TestCompany -ceq "DENYACCESS")) {
                    Set-ADUser -Identity $LoginName -Company "DENYACCESS"
                    write-host "${LoginName} set company to DENYACCESS"
                }
                If (!($TestTitle -ceq "DENYACCESS")) {
                    Set-ADUser -Identity $LoginName -Title "DENYACCESS"
                    write-host "${LoginName} Title change to: DENYACCESS"
                }
            }
        }

        # set additional user details if the user exists
        ElseIf ($TestUser) {

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
                write-host "${LoginName}: ${TestAccountName} Changed Given Name to ${PreferredName}"
            }
            If ($TestSurname -cne $Surname) {
                Set-ADUser -Identity $LoginName -Surname $Surname
                write-host "${LoginName}: ${TestAccountName} Changed Surname to ${SurName}"
            }
            If (($TestName -cne $FullName) -and ($TestName -cne $AltFullName)) {
                Try {
                    Rename-ADObject -Identity $TestDN -NewName $FullName
                    write-host "${LoginName}: ${TestAccountName} Changed Object Name to: ${FullName}"
                }
                Catch {
                    Rename-ADObject -Identity $TestDN -NewName $AltFullName
                    write-host "${LoginName}: ${TestAccountName} Changed Object Name to: ${AltFullName}"
                }
            }
            If (($TestDisplayName -cne $FullName)) {
                Set-ADUser -Identity $LoginName -DisplayName $FullName
                write-host "${LoginName}: ${TestAccountName} Changed Display Name to: ${FullName}"
            }

            # Set $StudCompany for automatic mail group filtering
            If (!($TestCompany -ceq $StudCompany)) {
                Set-ADUser -Identity $LoginName -Company $StudCompany
                write-host "${LoginName} set company to ${StudCompany}"
            }

            # Set Year Level and Title
            If (($TestTitle) -eq $null) {
                Set-ADUser -Identity $LoginName -Title $JobTitle
                write-host "${LoginName} Title change to: ${JobTitle}"
            }
            ElseIf (!($TestTitle).contains($JobTitle)) {
                Set-ADUser -Identity $LoginName -Title $JobTitle
                write-host "${LoginName} Title change to: ${JobTitle}"
            }

            # Get the year level of the current office string
            If ($TestOffice) {
                $test1 = $TestOffice.Substring($TestOffice.length-1,1)
                $test2 = $TestOffice.Substring($TestOffice.length-2,2)
            }
            Else {
                Set-ADUser -Identity $LoginName -Office $Position
                write-host "${LoginName} Office missing; set to ${Position}"
            }

            # set Office to current year level
            If ($YearGroup.length -eq 1) {
                If ($YearGroup -ne $test1) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host "${LoginName} year level change from ${TestOffice} to ${Position}"
                }
                ElseIf ($TestOffice -eq "Future Year ${YearGroup}") {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host "${LoginName} year level change from ${TestOffice} to ${Position}"
                }
            }
            ElseIf ($YearGroup.length -eq 2) {
                If ($YearGroup -ne $test2) {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host "${LoginName} year level change from ${TestOffice} to ${Position}"
                }
                ElseIf ($TestOffice -eq "Future Year ${YearGroup}") {
                    Set-ADUser -Identity $LoginName -Office $Position
                    write-host "${LoginName} year level change from ${TestOffice} to ${Position}"
                }
            }

            # Set Department to identify current students
            If (!(($TestDepartment) -ceq ("Student"))) {
                Set-ADUser -Identity $LoginName -Department "Student"
                write-host "${LoginName} Setting Position from ${TestDepartment} to Student"
            }

            # Check Group Membership
            If (!($StudentGroup.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity "Students" -Member $LoginName
                write-host "${LoginName} added Students Group"
            }
            # Check Group Membership
            If (!($StudmailGroup.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity $Studmail -Member $LoginName
                write-host "${LoginName} added Students Mail Group"
            }
            If (!($LocalUser.SamAccountName.contains($LoginName))) {
                write-host "${LoginName} Add to local user group for domain workstations"
                Add-ADGroupMember -Identity $UserRegular -Member $LoginName
            }
            # $MoodleStudentMembers
            If (!($MoodleStudentMembers.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity $MoodleStudent -Member $LoginName
                write-host "${LoginName} added MoodleStudent Group"
            }
            # $MoodleTechHelpMembers
            If (!($MoodleTechHelpMembers.SamAccountName.contains($LoginName))) {
                Add-ADGroupMember -Identity $MoodleTechHelp -Member $LoginName
                write-host "${LoginName} added MoodleTechHelp Group"
            }
            # $TestPrintGroup
            If (!($TestPrintGroup.name.contains($TestUser.name))) {
                Add-ADGroupMember -Identity $GenericPrintCode -Member $TestAccountName
                write-host "${LoginName} added default printer group ${GenericPrintCode}"
            }

            ### Remove groups for other grades and add the correct grade ###

            # Confirm membership to Year 5
            If ($YearGroup -eq "5") { # -and (!($5Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $5Name -Member $LoginName
                    Add-ADGroupMember -Identity $5mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "5")) {
                Remove-ADGroupMember -Identity $5Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $5mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 6
            If ($YearGroup -eq "6") { #-and (!($6Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $6Name -Member $LoginName
                    Add-ADGroupMember -Identity $6mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "6")) {
                Remove-ADGroupMember -Identity $6Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $6mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 7
            If ($YearGroup -eq "7") { # -and (!($7Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $7Name -Member $LoginName
                    Add-ADGroupMember -Identity $7mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "7")) {
                Remove-ADGroupMember -Identity $7Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $7mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 8
            If ($YearGroup -eq "8") { # -and (!($8Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $8Name -Member $LoginName
                    Add-ADGroupMember -Identity $8mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "8")) {
                Remove-ADGroupMember -Identity $8Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $8mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 9
            If ($YearGroup -eq "9") { # -and (!($9Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $9Name -Member $LoginName
                    Add-ADGroupMember -Identity $9mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "9")) {
                Remove-ADGroupMember -Identity $9Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $9mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 10
            If ($YearGroup -eq "10") { # -and (!($10Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $10Name -Member $LoginName
                    Add-ADGroupMember -Identity $10mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "10")) {
                Remove-ADGroupMember -Identity $10Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $10mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 11
            If ($YearGroup -eq "11") { # -and (!($11Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $11Name -Member $LoginName
                    Add-ADGroupMember -Identity $11mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "11")) {
                Remove-ADGroupMember -Identity $11Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $11mail -Member $LoginName -Confirm:$false
            }
            # Confirm membership to Year 12
            If ($YearGroup -eq "12") { # -and (!($12Group.SamAccountName.contains($LoginName)))) {
                    Add-ADGroupMember -Identity $12Name -Member $LoginName
                    Add-ADGroupMember -Identity $12mail -Member $LoginName
            }
            Elseif (!($YearGroup -eq "12")) {
                Remove-ADGroupMember -Identity $12Name -Member $LoginName -Confirm:$false
                Remove-ADGroupMember -Identity $12mail -Member $LoginName -Confirm:$false
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
                    # Remove all groups leaving domain users only
                    If (!($GroupName -eq $DomainUsersGroup)) {
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
