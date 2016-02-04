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
$input = Import-CSV ".\csv\fim_student.csv" -Encoding UTF8
$inputcount = (Import-CSV  ".\csv\fim_student.csv" -Encoding UTF8 | Measure-Object).Count
$idinput = Import-CSV  ".\csv\_CUSTOM_STUDENT_ID.csv" -Encoding UTF8

### Get Default Password From Secure String File
### http://www.adminarsenal.com/admin-arsenal-blog/secure-password-with-powershell-encrypting-credentials-part-1/
###
$userpass = cat C:\DATA\DefaultPassword.txt | convertto-securestring

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
$emailbody = "New AD user created
This is an automated email that is sent when a new user is created."
# message for disabled users
$disableemailsubject = "Current AD User Disabled:"
$disableemailbody = "Current AD user disabled
This is an automated email that is sent when an existing user is disabled."

# Get membership for group Membership Tests
$VillanovaGroups = Get-ADGroup -Filter *  -SearchBase "OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
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

write-host "### Completed importing groups"
write-host


################################################
### Create / Edit / Disable student accounts ###
################################################

# check log path
if (!(Test-Path ".\log")) {
    mkdir ".\log"
}

# set log file
$LogFile = “.\log\student-${LogDate}.log”
$LogContents = @()
$tmpcount = 0
$lastprogress = $NULL

write-host "### Processing Current Student File..."
Write-Host

foreach($line in $input) {
    $progress = ((($tmpcount / $inputcount) * 100) -as [int]) -as [string]
    if (((((($tmpcount / $inputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
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

    # Treat students who are not at their termination date as current
    #If ($DATE -le $Termination) {
    #    Write-Output "${LoginName} has termination date in the future. Process as current Student" | Out-File $LogFile -Append
    #    Write-Output "${Termination}" | Out-File $LogFile -Append
    #}

    If (($Termination.length -eq 0) -or ($DATE -le $Termination)) {

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
        # Home Folders are only for younger grades
        #IF (($YearGroup -eq "5")-or ($YearGroup -eq "6")) {
        #    $HomeDrive = "\\villanova.vnc.qld.edu.au\home\Student\${LoginName}"
        #}
        #Else {
        $HomeDrive = $null
        #}
        $JobTitle = "Student - ${YEAR}"
        
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
            #if (!($HomeDrive -eq $null)) {
            #    Try  {
            #        New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $False -homedrive "H" -homedirectory $HomeDrive
            #        $LogContents += "New User ${LoginName} created for ${FullName}" #| Out-File $LogFile -Append
            #        Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
            #    }
            #    Catch {
            #        $LogContents += "The User ${LoginName} already exists for ${FullName}" #| Out-File $LogFile -Append
            #    }
            #}
            #Else {
            Try  {
                New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $False
                $LogContents += "New User ${LoginName} created for ${FullName}" #| Out-File $LogFile -Append
                Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
            }
            Catch {
                Try {
                    # Error's can occur when the name of a student matches and they are in the same grade.
                    New-ADUser -SamAccountName $LoginName -Name $AltFullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -Description $UserCode -ChangePasswordAtLogon $False
                    #write-host "Possible duplicate name"
                    $LogContents += "New User ${LoginName} created for ${AltFullName}" #| Out-File $LogFile -Append
                    Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
                }
                Catch {
                    $LogContents += "The User ${LoginName} already exists for ${FullName} we tried: ${AltFullName}" #| Out-File $LogFile -Append
                }
            }
            #}
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

        # Get office365 details
        $TestEmail = $TestUser.mail
        If ($TestEmail) {
            $TestEmail = $TestEmail.ToLower()
        }
        $TestPrincipal = $TestUser.UserPrincipalName

        # set additional user details if the user exists
        If ($TestUser) {

            # Check that UPN is set to email. but only if an email exists
            If (($TestEmail) -and (!($TestEmail -ceq $TestPrincipal))) {
                Set-ADUser -Identity $TestDN -UserPrincipalName $TestEmail
                $LogContents += "UPN CHANGE: ${TestPrincipal} to ${TestEmail}" #| Out-File $LogFile -Append
                Write-Host "UPN CHANGE: ${TestPrincipal} to ${TestEmail}"
            }

            # Enable user if disabled
            If ((!($TestEnabled)) -and (!($TestDescription -eq "disable"))) {
                Set-ADUser -Identity $LoginName -Enabled $true
                $LogContents += "Enabling ${TestAccountName}" #| Out-File $LogFile -Append
            }
            # Disable if description contains disable
            ElseIf (($TestEnabled) -and ($TestDescription -eq "disable")) {
                Set-ADUser -Identity $LoginName -Enabled $false
                $LogContents += "Disabling ${TestAccountName}" #| Out-File $LogFile -Append
            }

            # Move user to the default OU for their year level if not there
            if (($TestEnabled) -and (!($TestDN.Contains($UserPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                $LogContents += "Taking ${TestAccountName} From: ${TestDN}" #| Out-File $LogFile -Append
                $LogContents += "Moving ${TestAccountName} To: ${UserPath}" #| Out-File $LogFile -Append
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
            # $MoodleTechHelpMembers
            if (!($MoodleTechHelpMembers.name.contains($TestName))) {
                Add-ADGroupMember -Identity $MoodleTechHelp -Member $LoginName
                write-host $LoginName "added MoodleTechHelp Group"
            }
            # $TestPrintGroup
            if (!($TestPrintGroup.name.contains($TestUser.name))) {
                Add-ADGroupMember -Identity $GenericPrintCode -Member $TestAccountName
                write-host $TestAccountName "added default printer group ${GenericPrintCode}"
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
            $tmpID = (Get-Culture).TextInfo.ToUpper($line.idcard_dec.Trim())
            $tmpNum = (Get-Culture).TextInfo.ToUpper($line.idcard_nfc.Trim())
            If ($TestAccountName -eq $tmpName) {
                If ($TestUser) {
                    # warn about mismatched fields
                    if ((!($tmpID.length -eq 0)) -and ($tmpNum.length -eq 0)) {
                        write-host "missing hex for ${TestAccountName}"
                    }
                    if (($tmpID.length -eq 0) -and (!($tmpNum.length -eq 0))) {
                        write-host "missing decimal for ${TestAccountName}"
                    }
                    # Add Employee ID if there is one
                    if ((!($TestID -ceq $tmpID)) -and (!($tmpID.length -eq 0))) {
                        Set-ADUser -Identity $LoginName -EmployeeID $tmpID
                        write-host "Setting decimal employeeID (${tmpID}) for ${TestAccountName}"
                        $LogContents += "Setting decimal employeeID (${tmpID}) for ${LoginName}"
                    }
                    # Add Employee Number if there is one
                    if (!($TestNumber -ceq $tmpNum) -and (!($tmpNum.length -eq 0))) {
                        Set-ADUser -Identity $LoginName -EmployeeNumber $tmpNum
                        write-host "Setting Hex employeeNumber (${tmpNum}) for ${TestAccountName}"
                        $LogContents += "Setting Hex employeeNumber (${tmpID}) for ${LoginName}"
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
                If (!($LoginName -eq '10961')) {
                    $LogContents += "${LoginName} Keeping terminated user" #| Out-File $LogFile -Append
                }
            }
            # Terminate Students AFTER their Termination date
            ElseIf ($DATE -gt $Termination) {
                # Disable The account when we don't want to keep it
                If ($TestUser) {
                    Set-ADUser -Identity $LoginName -Enabled $false
                    $LogContents += "DISABLING ACCOUNT ${TestAccountName}" #| Out-File $LogFile -Append
                    $LogContents += "Now: ${DATE}" #| Out-File $LogFile -Append
                    $LogContents += "DOL: ${Termination}" #| Out-File $LogFile -Append
                    Send-MailMessage -From $fromnotification -Subject "${disableemailsubject} ${LoginName}" -To $tonotification -Body $disableemailbody -SmtpServer $smtpserver
                }
            }
        }
        Else {
            # Enforce Group and OU changes for disabled students
            If ($TestUser) {
                # Move to disabled user OU if not already there
                if (!($TestDN.Contains($DisablePath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                    $LogContents += "Moving: ${TestAccountName} to Disabled Student OU" #| Out-File $LogFile -Append
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
                            $LogContents += "Error Removing ${TestAccountName} from ${GroupName}" #| Out-File $LogFile -Append
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
$enrolledinput = Import-CSV ".\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8
$enrolledcount = (Import-CSV  ".\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8 | Measure-Object).Count

# OU paths for different user types
$FuturePath = "OU=future,OU=student,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

$tmpcount = 0
$lastprogress = $NULL

foreach($line in $enrolledinput) {
    $progress = ((($tmpcount / $enrolledcount) * 100) -as [int]) -as [string]
    if (((((($tmpcount / $enrolledcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
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
            $LogContents += "${LoginName} created for ${FullName}" #| Out-File $LogFile -Append
            Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
        }
        Catch {
            # ignore future student error
            #$LogContents += "The username ${LoginName} for Future student ${FullName} already exists" #| Out-File $LogFile -Append
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
        If ($TestUser) {

            # Check that UPN is set to email. but only if an email exists
            If (($TestEmail) -and (!($TestEmail -ceq $TestPrincipal))) {
                Set-ADUser -Identity $TestDN -UserPrincipalName $TestEmail
                $LogContents += "UPN CHANGE: ${TestPrincipal} to ${TestEmail}" #| Out-File $LogFile -Append
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
            if (!($LoginName -ceq $TestNumber)) {
                Set-ADUser -Identity $LoginName -EmployeeNumber $LoginName
                write-host "${TestAccountName} Setting employee Number"
                write-host
            }
        }
        Else {
            write-host "missing or ignoring ${FullName}: ${LoginName}"
            write-host
        }
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