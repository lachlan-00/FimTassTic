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
$inputcount = (Import-CSV  ".\csv\fim_staffALL.csv" -Encoding UTF8 | Measure-Object).Count
$classinput = Import-CSV  ".\csv\fim_classes.csv" -Encoding UTF8
$idinput = Import-CSV  ".\csv\_CUSTOM_STAFF_ID.csv" -Encoding UTF8


# Check for the length of the import so you don't overwrite the content
$classCount = (Import-CSV  ".\csv\fim_classes.csv").count

### Get Default Password From Secure String File
### http://www.adminarsenal.com/admin-arsenal-blog/secure-password-with-powershell-encrypting-credentials-part-1/
###
$userpass = cat C:\DATA\DefaultPassword.txt | convertto-securestring

write-host
write-host "### Starting Staff Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$UserPath = "OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$ITPath = "OU=it,OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TeacherPath = "OU=teaching,OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$NonTeacherPath = "OU=nonteaching,OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$ReliefTeacherPath = "OU=relief,OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$OtherPath = "OU=other,OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TutorPath = "OU=tutors,OU=staff,OU=UserAccounts,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$DisablePath = "OU=staff,OU=users,OU=disabled,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

# Security Group names for staff
$StaffName = "CN=Staff,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TeacherName = "CN=S-G_Teachers,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TeacherMapName = "CN=Map-Teachers,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$MoodleName = "CN=MoodleTeacher,OU=RoleAssignment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
#$CanonName = "CN=Printer-Canon-Staff,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$GenericPrintCode = "CN=9000,OU=print,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach5Name = "CN=S-G_Teacher-Year5,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach6Name = "CN=S-G_Teacher-Year6,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach7Name = "CN=S-G_Teacher-Year7,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach8Name = "CN=S-G_Teacher-Year8,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach9Name = "CN=S-G_Teacher-Year9,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach10Name = "CN=S-G_Teacher-Year10,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach11Name = "CN=S-G_Teacher-Year11,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Teach12Name = "CN=S-G_Teacher-Year12,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail5Name = "CN=Teachers - Year 5,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail6Name = "CN=Teachers - Year 6,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail7Name = "CN=Teachers - Year 7,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail8Name = "CN=Teachers - Year 8,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail9Name = "CN=Teachers - Year 9,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail10Name = "CN=Teachers - Year 10,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail11Name = "CN=Teachers - Year 11,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$Mail12Name = "CN=Teachers - Year 12,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$JunPastName = "CN=Teachers - Junior Pastoral,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$MidPastName = "CN=Teachers - Middle Pastoral,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$SenPastName = "CN=Teachers - Senior Pastoral,OU=distribution,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

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
#student $DATE = "${DATE} 00:00:00"
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
$VillanovaGroups = Get-ADGroup -Filter * -SearchBase "OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TestStaff = Get-ADGroupMember -Identity $StaffName
#$TestCanonStaff = Get-ADGroupMember -Identity $CanonName
$TestPrintGroup = Get-ADGroupMember -Identity $GenericPrintCode
$TestTeachers = Get-ADGroupMember -Identity $TeacherName
$TestMoodleTeachers = Get-ADGroupMember -Identity $MoodleName
$TestMapTeachers = Get-ADGroupMember -Identity $TeacherMapName
#Year Levels From teaching class lists
$teaches5 = Get-ADGroupMember -Identity $Teach5Name
$teaches6 = Get-ADGroupMember -Identity $Teach6Name
$teaches7 = Get-ADGroupMember -Identity $Teach7Name
$teaches8 = Get-ADGroupMember -Identity $Teach8Name
$teaches9 = Get-ADGroupMember -Identity $Teach9Name
$teaches10 = Get-ADGroupMember -Identity $Teach10Name
$teaches11 = Get-ADGroupMember -Identity $Teach11Name
$teaches12 = Get-ADGroupMember -Identity $Teach12Name
#Year level teaching mail groups
$mail5 = Get-ADGroupMember -Identity $Mail5Name
$mail6 = Get-ADGroupMember -Identity $Mail6Name
$mail7 = Get-ADGroupMember -Identity $Mail7Name
$mail8 = Get-ADGroupMember -Identity $Mail8Name
$mail9 = Get-ADGroupMember -Identity $Mail9Name
$mail10 = Get-ADGroupMember -Identity $Mail10Name
$mail11 = Get-ADGroupMember -Identity $Mail11Name
$mail12 = Get-ADGroupMember -Identity $Mail12Name
# Teacher Pastoral groups
$JuniorPastoral = Get-ADGroupMember -Identity $JunPastName
$MiddlePastoral = Get-ADGroupMember -Identity $MidPastName
$SeniorPastoral = Get-ADGroupMember -Identity $SenPastName

write-host "### Completed importing groups"
write-host

##############################################
### Create / Edit / Disable Staff accounts ###
##############################################

# check log path
if (!(Test-Path ".\log")) {
    mkdir ".\log"
}

# set log file
$LogFile = “.\log\staff-${LogDate}.log”
#Write-output "" | Out-File $LogFile -Append
#Get-Date | Out-File $LogFile -Append
$LogContents = @()
$tmpcount = 0
$lastprogress = $NULL

write-host "### Processing Staff File..."
Write-Host

foreach($line in $input) {
    $progress = ((($tmpcount / $inputcount) * 100) -as [int]) -as [string]
    if (((((($tmpcount / $inputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
        Write-Host "Progress: ${progress}%"
    }
    $tmpcount = $tmpcount + 1
    $lastprogress = $progress

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
        If (($Position -ne $null) -and ($Position -ne "")) {
            if (($Position) -eq $line.position_title.Trim()) {
                $Position = (Get-Culture).TextInfo.ToLower($Position)
                $Position = (Get-Culture).TextInfo.ToTitleCase($Position)
                }
            Else {
                $Position = ($line.position_title.Trim())
            }
        }
        Elseif (($Position2 -ne $null) -and ($Position2 -ne "")) {
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
        $Position = $Position -replace "Av Technical Coordinator", "AV Technical Coordinator"
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

        $FullName =  "${PreferredName} ${Surname}"
        $DisplayName = $FullName
        $DisplayName = $DisplayName -replace "Peter Wieneke", "Fr. Peter Wieneke OSA"
        #$DisplayName = $DisplayName -replace "Peter Morris", "Dr. Peter Morris"
        #$DisplayName = $DisplayName -replace "Irene Lategan", "Dr. Irene Lategan"
        # Set remaining details
        ### Office 365 change ###$UserPrincipalName = "${LoginName}@villanova.vnc.qld.edu.au"
        $UserPrincipalName = "${LoginName}@vnc.qld.edu.au"
        $HomeDrive = "\\villanova.vnc.qld.edu.au\home\Staff\${LoginName}"
        $Telephone = $line.phone_w_text.Trim()
        If ($Telephone.length -le 1) {
            $Telephone = $null
        }
        $employeeNumber = (Get-Culture).TextInfo.ToLower($line.record_id.Trim())

        $emailbody = "New AD user created
This is an automated email that is sent when a new user is created.
${FullName}
${LoginName}
${Position}"

        ######################################
        ### Create / Modify Staff Accounts ###
        ######################################

        # create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            
            Try {
                New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $True
                # removing # -homedrive "H" -homedirectory $HomeDrive
                Set-ADUser -Identity $LoginName -Description $Position -Office $Position -Title $Position
                $LogContents += "New User ${LoginName} created for ${FullName}"
                Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
            }
            Catch {
                $LogContents += "${LoginName} already exists for ${FullName}"
            }
        }

        # Set user to confirm details
        $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)

        # Get Name Details
        $TestName = $TestUser.Name
        $TestGiven = $TestUser.GivenName
        $TestSurname = $TestUser.SurName
        $TestDisplayName = $TestUser.DisplayName

        # Get user details
        $TestDN = $TestUser.distinguishedname
        $TestAccountName = $TestUser.SamAccountName
        $TestEnabled = $TestUser.Enabled
        $TestHome = $TestUser.homedirectory

        # Get office365 details
        $TestEmail = $TestUser.mail
        If ($TestEmail) {
            $TestEmail = $TestEmail.ToLower()
        }
        $TestPrincipal = $TestUser.UserPrincipalName

        # Get general contact info and descriptions
        $TestTitle = $TestUser.Title
        $TestCompany = $TestUser.Company
        $TestOffice = $TestUser.Office
        $TestDescription = $TestUser.Description
        $TestDepartment = $TestUser.Department
        $TestNumber = $TestUser.employeeNumber
        $TestID = $TestUser.employeeID
        $TestPhone = $TestUser.OfficePhone


        # Check DN paths
        $TestPath = ($TestUser.distinguishedname.Contains($UserPath))
        $TestTeacherPath = ($TestUser.distinguishedname.Contains($TeacherPath))
        $TestNonTeacherPath = ($TestUser.distinguishedname.Contains($NonTeacherPath))
        $TestTutorPath = ($TestUser.distinguishedname.Contains($TutorPath))
        $TestITPath = ($TestUser.distinguishedname.Contains($ITPath))
        $TestReliefTeacherPath = ($TestUser.distinguishedname.Contains($ReliefTeacherPath))

        # set additional user details if the user exists
        If (($TestUser) -and (!($TestDepartment -eq "IGNORE"))) {

            # Check that UPN is set to email. but only if an email exists
            If (($TestEmail) -and (!($TestEmail -ceq $TestPrincipal))) {
                Set-ADUser -Identity $TestDN -UserPrincipalName $TestEmail
                #Write-Output "UPN CHANGE: ${TestPrincipal} to ${TestEmail}" | Out-File $LogFile -Append
                $LogContents += "UPN CHANGE: ${TestPrincipal} to ${TestEmail}"
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
            If (($TestDisplayName -cne $DisplayName)) {
                Set-ADUser -Identity $LoginName -DisplayName $DisplayName
                write-host "${TestAccountName} Changed Display Name to: ${DisplayName}"
            }
            #If ($TestUser.CN -cne $FullName) {
            #    Set-ADUser -Identity $LoginName -DisplayName $FullName
            #    write-host $TestUser.CN, "Changed Common Name to: ${FullName}"
            #}

            # Enable user if disabled
            If (!($TestEnabled)) {
                Set-ADUser -Identity $LoginName -Enabled $true
                write-host "Enabling", $TestAccountName
            }

            # Set userprofile path if is doesn't match
            If (!($TestHome -eq $HomeDrive)) {
                Set-ADUser -Identity $LoginName -homedrive "H:" -homedirectory $HomeDrive
                write-host "updated ${TestAccountName} home profile directory to: ${HomeDrive}"
            }

            # create home folder if it doesn't exist
            if (!(Test-Path $HomeDrive)) {
                New-Item -ItemType Directory -Force -Path $HomeDrive
            }

            # Add Position if there is one
            if (!($Position -ceq $TestDescription) -and (!($Position.length -eq 0))) {
                Set-ADUser -Identity $LoginName -Description $Position
                write-host $TestAccountName, "setting position"
                write-host "-${Position}-"
                write-host "-${TestDescription}-"
                write-host
            }

            # Add Office title
            if (!("Villanova College" -ceq $TestOffice)) {
                Set-ADUser -Identity $LoginName -Office "Villanova College"
                write-host $TestAccountName, "setting Office"
                write-host "-${Position}-"
                write-host "-${TestDescription}-"
                write-host
            }

            # Add title
            If (!($Position -ceq $TestTitle) -and (!($Position.length -eq 0))) {
                Set-ADUser -Identity $LoginName -Title $Position
                write-host $TestUser.Name, "Missing Title"
                write-host $Position
                write-host
            }
            
            # Set Department to identify current staff
            If (!(($TestUser.Department) -ceq ("Staff"))) {
                Set-ADUser -Identity $LoginName -Department "Staff"
                write-host $TestUser.Name, "Setting Position:", $TestUser.Department
                write-host
            }

            # Add Telephone number if there is one
            if ($Telephone -ne $TestPhone) {
                If ($Telephone -eq $null) {
                    if ($TestPhone -ne "690") {
                        Set-ADUser -Identity $LoginName -OfficePhone "690"
                        write-host $TestAccountName, "setting Telephone to Default (690)"
                        write-host
                    }
                }
                Else {
                    Set-ADUser -Identity $LoginName -OfficePhone $Telephone
                    write-host $TestAccountName, "setting Telephone to:", $Telephone
                    write-host
                }
            }

            # Move user to their default OU if not already there
            if ($TestUser.description -eq $null) {
                if (!($TestUser.distinguishedname.Contains($OtherPath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $OtherPath
                    #write-output "Moving to staff\other OU: no description for ${LoginName}" | Out-File $LogFile -Append
                    $LogContents += "Moving to staff\other OU: no description for ${LoginName}"
                }
            }
            Elseif ($TestUser.distinguishedname.Contains($DisablePath)) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                #write-output $TestAccountName "moved out of Disabled OU" | Out-File $LogFile -Append
                $LogContents += "${TestAccountName} moved out of Disabled OU"
            }
            ElseIf (($TestCompany  -ceq "Relief Teacher") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                #write-output $TestAccountName "moved to Relief Teacher OU" | Out-File $LogFile -Append
                $LogContents += "${TestAccountName} moved to Relief Teacher OU"
            }
            ElseIf ($TestCompany -ceq "Teacher") {
                If ($TestUser.distinguishedname.Contains($ReliefTeacherPath)) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                    #write-output $TestAccountName "moved to  Teacher OU from Relief Teachers" | Out-File $LogFile -Append
                    $LogContents += "${TestAccountName} moved to  Teacher OU from Relief Teachers"
                }
                ElseIf (!($TestUser.distinguishedname.Contains($TeacherPath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                    #write-output $TestAccountName "moved to  Teacher OU" | Out-File $LogFile -Append
                    $LogContents += "${TestAccountName} moved to  Teacher OU"
                }
            }
            ElseIf (($TestCompany -ceq "Tutors") -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                #write-output $TestAccountName "moved to Music Tutor OU" | Out-File $LogFile -Append
                $LogContents += "${TestAccountName} moved to Music Tutor OU"
            }
            ElseIf (($TestPath -and (!($TestTeacherPath))) -and (!($TestNonTeacherPath)) -and (!($TestITPath)) -and (!($TestTutorPath)) -and (!($TestReliefTeacherPath))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $NonTeacherPath
                #write-output $TestAccountName "moved to non-teaching" | Out-File $LogFile -Append
                $LogContents += "${TestAccountName} moved to non-teaching"
            }

            # Set company for automatic mail group filtering
            if (($TestUser.distinguishedname.Contains($NonTeacherPath)) -and (!($TeacherCode))) {
                if ((!($TestCompany -ceq "Admin")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Admin"
                    write-host $TestUser.Name "set company to Admin"
                }
            }
            if ($TestUser.distinguishedname.Contains($ITPath)) {
                if ((!($TestCompany -ceq "ICT")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "ICT"
                    write-host $TestUser.Name "set company to ICT"
                }
            }
            if ($TestUser.distinguishedname.Contains($TeacherPath)) {
                if ((!($TestCompany -ceq "Teacher")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Teacher"
                    write-host $TestUser.Name "set company to Teacher"
                }
            }
            if ($TestUser.distinguishedname.Contains($ReliefTeacherPath)) {
                if ((!($TestCompany -ceq "Relief")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Relief"
                    write-host $TestUser.Name "set company to Teacher"
                }
            }
            if ($TestUser.distinguishedname.Contains($TutorPath)) {
                if ((!($TestCompany -ceq "Tutors")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Tutors"
                    write-host $TestUser.Name "set company to Tutors"
                }
            }

            # Check Group Membership
            if (!($TestStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "Staff" -Member $TestAccountName
                        write-host $TestAccountName "added Staff"
            }
            #if (!($TestCanonStaff.name.contains($TestUser.name))) {
            #            Add-ADGroupMember -Identity $CanonName -Member $TestAccountName
            #            write-host $TestAccountName "added Printer-Canon-Staff"
            #}
            # $TestPrintGroup
            if (!($TestPrintGroup.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity $GenericPrintCode -Member $TestAccountName
                        write-host $TestAccountName "added default printer group ${GenericPrintCode}"
            }

            foreach($line in $idinput) {
                $tmpName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())
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

        ###################################################################
        ### Create / Edit Teacher Info for Staff with existing accounts ###
        ###################################################################

        If (($TeacherCode -ne $null) -and ($TeacherCode -ne "")) {

            # Set user to confirm details
            $TestUser = (Get-ADUser  -Filter { (SamAccountName -eq $LoginName) }  -Properties *)
            $TestAccountName = $TestUser.SamAccountName
            $TestDN = $TestUser.distinguishedname
            $TestDescription = $TestUser.Description

            #Make sure Teachers have the correct Company
            if ((!($TestCompany -ceq "Teacher")) -and (!($TestDescription.Contains("Relief Teacher")))) {
                Set-ADUser -Identity $LoginName -Company "Teacher"
                write-host "Changing Company for ${TestAccountName} to Teacher"
                write-host
                #refresh details again
                $TestUser = (Get-ADUser  -Filter { (SamAccountName -eq $LoginName) }  -Properties *)
                $TestAccountName = $TestUser.SamAccountName
                $TestDescription = $TestUser.Description
            }

            If ($TestUser.Enabled) {

                # Move to Teacher OU if not already there
                if ($TestDN.Contains($UserPath) -and (!($TestDN.Contains($TeacherPath))) -and (!($TestDN.Contains($ReliefTeacherPath)))) {
                    If ($TestDescription.Contains("Relief Teacher") -and (!($TestDN.Contains($ReliefTeacherPath)))) {
                        Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                        write-host $TestAccountName "moved to Relief Teacher OU"
                    }
                    ElseIf (($TestDescription.Contains("Tutor")) -and (!($TestDN.Contains($TutorPath)))) {
                        Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                        write-host $TestAccountName "moved to Music Tutor OU"
                    }
                    ElseIf ((!($Description.Contains("Tutor"))) -and (!($TestDN.Contains($TeacherPath)))) {
                        Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                        write-host $TestAccountName "moved to Teacher OU"
                    }
                }
                # Check Group Membership
                if (!($TestTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $TeacherName -Member $TestAccountName
                    write-host $TestAccountName "ADDED to Teachers Group"
                }
                if (!($TestMoodleTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $MoodleName -Member $TestAccountName
                    write-host $TestAccountName "ADDED to MoodleTeachers Group"
                }
                if (!($TestMapTeachers.name.contains($TestUser.name))) {
                    Add-ADGroupMember -Identity $TeacherMapName -Member $TestAccountName
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
                If ($classCount -lt 500) {
                    write-host " No Classes Available"
                    $LogContents += "Not enough classes to do group management: Counted ${classCount}."
                }
                Else {
                    foreach($line in $classinput) {
                        $tmpteach = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())
                        $tmpyear = (Get-Culture).TextInfo.ToLower($line.year_grp.Trim())
                        $tmpsubtitle = (Get-Culture).TextInfo.ToLower($line.sub_long.Trim())
                        If ($LoginName -eq $tmpteach) {
                            If (($tmpyear -eq "5") -and (!($classin5))) {
                                $classin5 = $true
                                #write-host "Found Year 5 Class"
                                if ($tmpsubtitle -eq "Junior School Pastoral") {
                                    $classjuniorpastoral = $true
                                }
                            }
                            ElseIf (($tmpyear -eq "6") -and (!($classin6))) {
                                $classin6 = $true
                                #write-host "Found Year 6 Class"
                                if ($tmpsubtitle -eq "Junior School Pastoral") {
                                    $classjuniorpastoral = $true
                                }
                            }
                            ElseIf (($tmpyear -eq "7") -and (!($classin7))) {
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
                            ElseIf (($tmpyear -eq "8") -and (!($classin8))) {
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
                            ElseIf (($tmpyear -eq "9") -and (!($classin9))) {
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
                            ElseIf ($tmpyear -eq "10") {
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
                            ElseIf ($tmpyear -eq "11") {
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
                            ElseIf ($tmpyear -eq "12") {
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
                }

                #Part Time STAFF ??? HACK
                If ($TestAccountName -eq "wilsom") {
                    $classin5 = $true
                }

                # add teachers to year level teaching groups from classes
                # remove teachers from year level teaching groups if there are no classes found
                If ($classin5) {
                    #write-host "Found Year 5 Class"
                    Try{
                        #write-host "Adding Year 5 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach5Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 5 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail5Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach5Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail5Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin6) {
                    #write-host "Found Year 6 Class"
                    Try{
                        #write-host "Adding Year 6 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach6Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 6 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail6Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach6Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail6Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin7) {
                    #write-host "Found Year 7 Class"
                    Try{
                        #write-host "Adding Year 7 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach7Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 7 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail7Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach7Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail7Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin8) {
                    #write-host "Found Year 8 Class"
                    Try{
                        #write-host "Adding Year 8 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach8Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 8 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail8Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach8Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail8Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin9) {
                    #write-host "Found Year 9 Class"
                    Try{
                        #write-host "Adding Year 9 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach9Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 9 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail9Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach9Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail9Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin10) {
                    #write-host "Found Year 10 Class"
                    Try{
                        #write-host "Adding Year 10 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach10Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 10 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail10Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach10Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail10Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin11) {
                    #write-host "Found Year 11 Class"
                    Try{
                        #write-host "Adding Year 11 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach11Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 11 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail11Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach11Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail11Name -Member $TestAccountName -Confirm:$false
                }
                If ($classin12) {
                    #write-host "Found Year 12 Class"
                    Try{
                        #write-host "Adding Year 12 Security Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Teach12Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try{
                        #write-host "Adding Year 12 Mail Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $Mail12Name -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $Teach12Name -Member $TestAccountName -Confirm:$false
                    Remove-ADGroupMember -Identity $Mail12Name -Member $TestAccountName -Confirm:$false
                }
                If ($classjuniorpastoral) {
                    Try{
                        #write-host "Adding Junior Pastoral Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $JunPastName -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $JunPastName -Member $TestAccountName -Confirm:$false
                }
                If ($classmiddlepastoral) {
                    Try{
                        #write-host "Adding Middle Pastoral Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $MidPastName -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $MidPastName -Member $TestAccountName -Confirm:$false
                }
                If ($classseniorpastoral) {
                    Try{
                        #write-host "Adding Senior Pastoral Group for: ${TestAccountName}"
                        Add-ADGroupMember -Identity $SenPastName -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $SenPastName -Member $TestAccountName -Confirm:$false
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
                #write-output "${LoginName} Keeping terminated user" | Out-File $LogFile -Append
                if ((!($LoginName = 'obrij')) -or (!($LoginName = 'latei'))) {
                    $LogContents += "${LoginName} Keeping terminated user"
                }
            }
            # Terminate Staff AFTER their Termination date
            ElseIf ($DATE -gt $Termination) {
                # Disable The account when we don't want to keep it
                If ($TestUser) {
                    Set-ADUser -Identity $LoginName -Enabled $false
                    #write-output "DISABLING ACCOUNT ${TestAccountName}" | Out-File $LogFile -Append
                    #Write-Output "Now: ${DATE}" | Out-File $LogFile -Append
                    #Write-Output "DOL: ${Termination}" | Out-File $LogFile -Append
                    $LogContents += "DISABLING ACCOUNT ${TestAccountName}"
                    $LogContents += "Now: ${DATE}"
                    $LogContents += "DOL: ${Termination}"
                    Send-MailMessage -From $fromnotification -Subject "${disableemailsubject} ${LoginName}" -To $tonotification -Body $disableemailbody -SmtpServer $smtpserver
                }
            }
        }
        Else {
            # Enforce Group and OU changes for disabled staff
            If ($TestUser) {
                # Move to disabled user OU if not already there
                if (!($TestUser.distinguishedname.Contains($DisablePath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                    #write-output "Moving: ${TestAccountName} to Disabled Staff OU" | Out-File $LogFile -Append
                    $LogContents += "Moving: ${TestAccountName} to Disabled Staff OU"
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
                            #Write-output "Error Removing ${TestAccountName} from ${GroupName}" | Out-File $LogFile -Append
                            $LogContents +=  "Error Removing ${TestAccountName} from ${GroupName}"
                        }
                    }
                }
            }
        }
    }
}

Write-Host
Write-Host "### Staff file finished"
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