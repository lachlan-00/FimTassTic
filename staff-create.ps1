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
$input = Import-CSV "C:\DATA\csv\fim_staffALL.csv" -Encoding UTF8
$inputcount = (Import-CSV "C:\DATA\csv\fim_staffALL.csv" -Encoding UTF8 | Measure-Object).Count
$classinput = Import-CSV "C:\DATA\csv\fim_classes.csv" -Encoding UTF8
$idinput = Import-CSV "C:\DATA\csv\_CUSTOM_STAFF_ID.csv" -Encoding UTF8


# Check for the length of the import so you don't overwrite the content
$classCount = (Import-CSV "C:\DATA\csv\fim_classes.csv").count

### Get Default Password From Secure String File
### http://www.adminarsenal.com/admin-arsenal-blog/secure-password-with-powershell-encrypting-credentials-part-1/
###
### REPLACED WITH RANDOM GENERATED PASSWORD ###
###$userpass = cat "C:\DATA\DefaultPassword.txt" | convertto-securestring


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
$MoodlePlaypen = "CN=PP-teachers,OU=teacher,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
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
$tonotification = "notifications@vnc.qld.edu.au"
# specify where the notifications come from
$fromnotification = "notifications@vnc.qld.edu.au"
# specify the SMTP server
$smtpserver = "mail.vnc.qld.edu.au"
# message for created users
$emailsubject = "New AD User Created:"
# message for disabled users
$disableemailsubject = "Current AD User Disabled:"

# Get membership for group Membership Tests
$VillanovaGroups = Get-ADGroup -Filter * -SearchBase "OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TestStaff = Get-ADGroupMember -Identity $StaffName
#$TestCanonStaff = Get-ADGroupMember -Identity $CanonName
$TestPrintGroup = Get-ADGroupMember -Identity $GenericPrintCode
$TestTeachers = Get-ADGroupMember -Identity $TeacherName
$TestMoodleTeachers = Get-ADGroupMember -Identity $MoodleName
$TestMoodlePlaypen = Get-ADGroupMember -Identity $MoodlePlaypen
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
If (!(Test-Path "C:\DATA\log")) {
    mkdir "C:\DATA\log"
}

# set log file
$LogFile = "C:\DATA\log\staff-${LogDate}.log"
$LogContents = @()
$tmpcount = 0
$lastprogress = $NULL

write-host "### Processing Staff File..."
Write-Host

foreach($line in $input) {
    $progress = ((($tmpcount / $inputcount) * 100) -as [int]) -as [string]
    If (((((($tmpcount / $inputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
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

        If ($PreferredName -eq $line.prefer_name_text.Trim()) {
            $PreferredName = (Get-Culture).TextInfo.ToLower($PreferredName)
            $PreferredName = (Get-Culture).TextInfo.ToTitleCase($PreferredName)
        }
        Else {
            $PreferredName = ($line.prefer_name_text.Trim())
        }
        If (($Surname) -eq $line.surname_text.Trim()) {
            $Surname = (Get-Culture).TextInfo.ToLower($Surname)
            $Surname = (Get-Culture).TextInfo.ToTitleCase($Surname)
        }
        Else {
            $Surname = ($line.surname_text.Trim())
        }
        If (($Position -ne $null) -and ($Position -ne "")) {
            If (($Position) -eq $line.position_title.Trim()) {
                $Position = (Get-Culture).TextInfo.ToLower($Position)
                $Position = (Get-Culture).TextInfo.ToTitleCase($Position)
                }
            Else {
                $Position = ($line.position_title.Trim())
            }
        }
        ElseIf (($Position2 -ne $null) -and ($Position2 -ne "")) {
            If (($Position2) -eq $line.position_text.Trim()) {
                $Position2 = (Get-Culture).TextInfo.ToLower($Position2)
                $Position2 = (Get-Culture).TextInfo.ToTitleCase($Position2)
                $Position = $Position2
                }
            Else {
                $Position = ($line.position_text.Trim())
            }
        }

        If (($Position2 -contains "Music Tutor") -or ($Position2 -contains "Relief Teacher") -or ($Position2 -contains "Teacher Relief")) {
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
        $Position = $Position -replace "Director of Music", "Director of Music / Artistic Director QCMF"
        $Position = $Position -replace "Music Department Assistant", "Villanova Music Department / QCMF Facilitator"
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
        $Surname = $Surname -replace "O'h", "O'H"
        $Surname = $Surname -replace "O'k", "O'K"
        $Surname = $Surname -replace "O'n", "O'N"
        $Surname = $Surname -replace "O'r", "O'R"

        $FullName = "${PreferredName} ${Surname}"
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

        ### Generate a Random password
        ### http://kunaludapi.blogspot.com.au/2013/11/generate-random-password-powershell.html
        ###
        $alphabets= "abcdefghjkmnopqstuvwxyz1234567890"
        $char = for ($i = 0; $i -lt $alphabets.length; $i++) { $alphabets[$i] }
        $randompass = ""
        for ($i = 1; $i -le 8; $i++)
        {
            $randompass += $(get-random -InputObject $char -Count 1)
        }
        $userpass = ConvertTo-SecureString -String "${randompass}" –AsPlainText -Force
        $emailbody = "There has been a new AD user created on the network.

Full Name: ${FullName}
User Name: ${LoginName}
Password: ${randompass}
Teacher Code: ${TeacherCode}
Position: ${Position}
Phone Ext: ${Telephone}

 * If any of the above fields are blank they must be filled in ASAP.
 * An email account will be created in the next hour.
 * To allow printing, a group of cost codes must be given to IT.
 * An ID card is required for:
    Flexi Schools
    Printing
    Door Locks
 * Many groups are automated but some additional information may be required.
   (Information about who this person is replacing can speed this up)
 * Access to TASS.Web, Teacher Kiosk and Web.Book to be set up as required.
 * Any changes to these details must be made in TASS.Web payroll.

###########################
This is an automated email.
###########################"

        $errorbody = "Incomplete Staff Data

Full Name: ${FullName}
User Name: ${LoginName}
Password: ${randompass}
Teacher Code: ${TeacherCode}
Position: ${Position}
Phone Ext: ${Telephone}

Please check this out ASAP"

        ######################################
        ### Create / Modify Staff Accounts ###
        ######################################

        # create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            #check values from data
            If ((!($LoginName)) -or (!($FullName)) -or (!($PreferredName)) -or (!($Surname)) -or ($UserPrincipalName -eq "@vnc.qld.edu.au")) {
                write-host "Error in data for ${LoginName}"
                # log it
                $LogContents += "Incomplete Data!"
                $LogContents += ""
                $LogContents += "LoginName = ${LoginName}"
                $LogContents += "PreferredName = ${PreferredName}"
                $LogContents += "Surname = ${Surname}"
                $LogContents += "FullName = ${FullName}"
                $LogContents += "UserPrincipalName = ${UserPrincipalName}"
                $LogContents += ""
                $LogContents += "Please check this out ASAP"
                # mail it
                Send-MailMessage -From $fromnotification -Subject "User Data Error" -To $tonotification -Body $errorbody -SmtpServer $smtpserver
            }
            # create user when not found
            Try {
                New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword $userpass -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $False
                # removing # -homedrive "H" -homedirectory $HomeDrive
                $LogContents += "New User ${LoginName} created for ${FullName}"
                Send-MailMessage -From $fromnotification -Subject "${emailsubject} ${LoginName}" -To $tonotification -Body $emailbody -SmtpServer $smtpserver
            }
            Catch {
                $LogContents += "${LoginName} already exists for ${FullName}"
            }
            # try to set basic position info
            Try {
                Set-ADUser -Identity $LoginName -Description $Position -Office $Position -Title $Position
            }
            Catch {
                $LogContents += "${LoginName} couldn't set position for ${FullName}"
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
                $LogContents += "${TestAccountName} Changed UPN to: ${TestEmail}"
            }

            # Check Name Information
            If ($TestGiven -cne $PreferredName) {
                Set-ADUser -Identity $LoginName -GivenName $PreferredName
                write-host "${TestAccountName} Changed Given Name to: ${PreferredName}"
            }
            If ($TestSurname -cne $Surname) {
                Set-ADUser -Identity $LoginName -Surname $Surname
                write-host "${TestAccountName} Changed Surname to: ${SurName}"
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
                write-host "${TestAccountName} Changed homedirectory to: ${HomeDrive}"
            }

            # create home folder if it doesn't exist
            If (!(Test-Path $HomeDrive)) {
                New-Item -ItemType Directory -Force -Path $HomeDrive
            }

            # Add Position if there is one
            If (!($Position -ceq $TestDescription) -and (!($Position.length -eq 0))) {
                Set-ADUser -Identity $LoginName -Description $Position
                write-host $TestAccountName, "Changed position to:", $Position
                write-host
            }

            # Add Office title
            If (!("Villanova College" -ceq $TestOffice)) {
                Set-ADUser -Identity $LoginName -Office "Villanova College"
                write-host $TestAccountName, "Changed Office to: Villanova College"
                write-host
            }

            # Add title
            If (!($Position -ceq $TestTitle) -and (!($Position.length -eq 0))) {
                Set-ADUser -Identity $LoginName -Title $Position
                write-host $TestUser.Name, "Changed Title to:", $Position
                write-host
            }
            
            # Set Department to identify current staff
            If (!(($TestUser.Department) -ceq ("Staff"))) {
                Set-ADUser -Identity $LoginName -Department "Staff"
                write-host $TestUser.Name, "Changed Position to:", $TestUser.Department
                write-host
            }

            # Add Telephone number if there is one
            If ($Telephone -ne $TestPhone) {
                If ($Telephone -eq $null) {
                    If ($TestPhone -ne "690") {
                        Set-ADUser -Identity $LoginName -OfficePhone "690"
                        write-host $TestAccountName, "Changed Telephone to: Default (690)"
                        write-host
                    }
                }
                Else {
                    Set-ADUser -Identity $LoginName -OfficePhone $Telephone
                    write-host $TestAccountName, "Changed Telephone to:", $Telephone
                    write-host
                }
            }

            # Move user to their default OU if not already there
            If ($TestUser.description -eq $null) {
                If (!($TestUser.distinguishedname.Contains($OtherPath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $OtherPath
                    $LogContents += "Moving to staff\other OU: no description for ${LoginName}"
                }
            }
            ElseIf ($TestUser.distinguishedname.Contains($DisablePath)) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $UserPath
                $LogContents += "${TestAccountName} moved out of Disabled OU"
            }
            ElseIf (($TestCompany -ceq "Relief Teacher") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                $LogContents += "${TestAccountName} moved to Relief Teacher OU"
            }
            ElseIf ($TestCompany -ceq "Teacher") {
                If ($TestUser.distinguishedname.Contains($ReliefTeacherPath)) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                    $LogContents += "${TestAccountName} moved to Teacher OU from Relief Teachers"
                }
                ElseIf (!($TestUser.distinguishedname.Contains($TeacherPath))) {
                    Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                    $LogContents += "${TestAccountName} moved to Teacher OU"
                }
            }
            ElseIf (($TestCompany -ceq "Tutors") -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                $LogContents += "${TestAccountName} moved to Music Tutor OU"
            }
            ElseIf (($TestPath -and (!($TestTeacherPath))) -and (!($TestNonTeacherPath)) -and (!($TestITPath)) -and (!($TestTutorPath)) -and (!($TestReliefTeacherPath))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $NonTeacherPath
                $LogContents += "${TestAccountName} moved to non-teaching"
            }

            # Set company for automatic mail group filtering
            If (($TestUser.distinguishedname.Contains($NonTeacherPath)) -and (!($TeacherCode))) {
                If ((!($TestCompany -ceq "Admin")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Admin"
                    write-host $TestUser.Name "set company to Admin"
                }
            }
            If ($TestUser.distinguishedname.Contains($ITPath)) {
                If ((!($TestCompany -ceq "ICT")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "ICT"
                    write-host $TestUser.Name "set company to ICT"
                }
            }
            If ($TestUser.distinguishedname.Contains($TeacherPath)) {
                If ((!($TestCompany -ceq "Teacher")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Teacher"
                    write-host $TestUser.Name "set company to Teacher"
                }
            }
            If ($TestUser.distinguishedname.Contains($ReliefTeacherPath)) {
                If ((!($TestCompany -ceq "Relief")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Relief"
                    write-host $TestUser.Name "set company to Teacher"
                }
            }
            If ($TestUser.distinguishedname.Contains($TutorPath)) {
                If ((!($TestCompany -ceq "Tutors")) -or ($TestCompany -eq $null)) {
                    Set-ADUser -Identity $LoginName -Company "Tutors"
                    write-host $TestUser.Name "set company to Tutors"
                }
            }

            # Check Group Membership
            If (!($TestStaff.SamAccountName.contains($LoginName))) {
                        Add-ADGroupMember -Identity "Staff" -Member $TestAccountName
                        write-host $TestAccountName "added Staff"
            }
            If (!($TestPrintGroup.SamAccountName.contains($LoginName))) {
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

        ###################################################################
        ### Create / Edit Teacher Info for Staff with existing accounts ###
        ###################################################################

        If (($TeacherCode -ne $null) -and ($TeacherCode -ne "")) {

            # Set user to confirm details
            $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
            $TestDescription = $TestUser.Description
            $TestCompany = $TestUser.Company

            #Make sure Teachers have the correct Company
            If (($TestDescription) -and ($TestCompany)) {
                If ((!($TestCompany -ceq "Teacher")) -and (!($TestDescription.Contains("Relief Teacher")))) {
                    Set-ADUser -Identity $LoginName -Company "Teacher"
                    write-host "Changing Company for ${LoginName} to Teacher"
                    write-host
                    #refresh details again
                    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
                    $TestDescription = $TestUser.Description
                    $TestCompany = $TestUser.Company
                }
            }

            # Get additional details
            $TestAccountName = $TestUser.SamAccountName
            $TestDN = $TestUser.distinguishedname
            $TestTitle = $TestUser.Title
            $TestOffice = $TestUser.Office
            $TestDepartment = $TestUser.Department

            If ($TestUser.Enabled) {

                # Move to Teacher OU if not already there
                If ($TestDescription) {
                    If ($TestDN.Contains($UserPath) -and (!($TestDN.Contains($TeacherPath))) -and (!($TestDN.Contains($ReliefTeacherPath)))) {
                        If ($TestDescription.Contains("Relief Teacher") -and (!($TestDN.Contains($ReliefTeacherPath)))) {
                            Get-ADUser $TestAccountName | Move-ADObject -TargetPath $ReliefTeacherPath
                            write-host $TestAccountName "moved to Relief Teacher OU"
                        }
                        ElseIf (($TestDescription.Contains("Tutor")) -and (!($TestDN.Contains($TutorPath)))) {
                            Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TutorPath
                            write-host $TestAccountName "moved to Music Tutor OU"
                        }
                        ElseIf ((!($TestDescription.Contains("Tutor"))) -and (!($TestDN.Contains($TeacherPath)))) {
                            Get-ADUser $TestAccountName | Move-ADObject -TargetPath $TeacherPath
                            write-host $TestAccountName "moved to Teacher OU"
                        }
                    }
                }
                # Check Group Membership
                If (!($TestTeachers.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $TeacherName -Member $TestAccountName
                    write-host $TestAccountName "ADDED to Teachers Group"
                }
                If (!($TestMoodleTeachers.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $MoodleName -Member $TestAccountName
                    write-host $TestAccountName "ADDED to MoodleTeachers Group"
                }
                If (!($TestMapTeachers.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $TeacherMapName -Member $TestAccountName
                    write-host $TestAccountName "ADDED to Map-Teachers Group"
                }
                # $TestMoodlePlaypen
                If (!($TestMoodlePlaypen.SamAccountName.contains($LoginName))) {
                    Add-ADGroupMember -Identity $MoodlePlaypen -Member $TestAccountName
                    write-host $TestAccountName "ADDED to MoodlePlaypen Group"
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
                                If ($tmpsubtitle -eq "Junior School Pastoral") {
                                    $classjuniorpastoral = $true
                                }
                            }
                            ElseIf (($tmpyear -eq "6") -and (!($classin6))) {
                                $classin6 = $true
                                If ($tmpsubtitle -eq "Junior School Pastoral") {
                                    $classjuniorpastoral = $true
                                }
                            }
                            ElseIf (($tmpyear -eq "7") -and (!($classin7))) {
                                $classin7 = $true
                                If ($tmpsubtitle -eq "Crane Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Goold Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Heavey Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Murray Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                            }
                            ElseIf (($tmpyear -eq "8") -and (!($classin8))) {
                                $classin8 = $true
                                If ($tmpsubtitle -eq "Crane Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Goold Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Heavey Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Murray Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                            }
                            ElseIf (($tmpyear -eq "9") -and (!($classin9))) {
                                $classin9 = $true
                                If ($tmpsubtitle -eq "Crane Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Goold Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Heavey Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                                If ($tmpsubtitle -eq "Murray Pastoral") {
                                    $classmiddlepastoral = $true
                                }
                            }
                            ElseIf ($tmpyear -eq "10") {
                                $classin10 = $true
                                If ($tmpsubtitle -eq "Crane Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Goold Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Heavey Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Murray Pastoral") {
                                    $classseniorpastoral = $true
                                }
                            }
                            ElseIf ($tmpyear -eq "11") {
                                $classin11 = $true
                                If ($tmpsubtitle -eq "Crane Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Goold Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Heavey Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Murray Pastoral") {
                                    $classseniorpastoral = $true
                                }
                            }
                            ElseIf ($tmpyear -eq "12") {
                                $classin12 = $true
                                If ($tmpsubtitle -eq "Crane Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Goold Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Heavey Pastoral") {
                                    $classseniorpastoral = $true
                                }
                                If ($tmpsubtitle -eq "Murray Pastoral") {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach5Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach6Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach7Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach8Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach9Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach10Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach11Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $Teach12Name -Member $TestAccountName
                    }
                    Catch {
                    }
                    Try {
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
                    Try {
                        Add-ADGroupMember -Identity $JunPastName -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $JunPastName -Member $TestAccountName -Confirm:$false
                }
                If ($classmiddlepastoral) {
                    Try {
                        Add-ADGroupMember -Identity $MidPastName -Member $TestAccountName
                    }
                    Catch {
                    }
                }
                Else {
                    Remove-ADGroupMember -Identity $MidPastName -Member $TestAccountName -Confirm:$false
                }
                If ($classseniorpastoral) {
                    Try {
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
        $TestTitle = $TestUser.Title
        $TestCompany = $TestUser.Company
        $TestOffice = $TestUser.Office
        $TestDescription = $TestUser.Description
        $TestDepartment = $TestUser.Department
        $TestEnabled = $TestUser.Enabled
        $TestAccountName = $TestUser.SamAccountName
        $TestMembership = $TestUser.MemberOf
        $disableemailbody = "A current AD user has been disabled.

This email is sent when an existing user is disabled.
Full Name: ${FullName}
User Name: ${LoginName}
Teacher Code: ${TeacherCode}
Position: ${Position}
Phone Ext: ${Telephone}

 * If this is a mistake please check the Termination Date in TASS.Web payroll.
 * Any other queries about this email can be forwarded to IT.

###########################
This is an automated email.
###########################"

        # Disable users with a termination date if they are still enabled
        If ($TestEnabled) {

            # Don't disable users we want to keep
            If ($TestDescription -eq "keep") {
                $LogContents += "${LoginName} Keeping terminated user"
            }
            # Terminate Staff AFTER their Termination date
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
            If (!($TestUser.distinguishedname.Contains($DisablePath))) {
                Get-ADUser $TestAccountName | Move-ADObject -TargetPath $DisablePath
                $LogContents += "Moving: ${TestAccountName} to Disabled Staff OU"
            }
            else {
                # Set Department to "Disabled" to help identify current staff
                If (!(($TestDepartment) -ceq ("Disabled"))) {
                    Set-ADUser -Identity $LoginName -Department "Disabled"
                    write-host "${LoginName} Setting Position from ${TestDepartment} to Disabled"
                }
                # Set Company to "Disabled" to help identify current staff
                If (!($TestCompany -ceq "Disabled")) {
                    Set-ADUser -Identity $LoginName -Company "Disabled"
                    write-host "${LoginName} set company from ${TestCompany} to Disabled"
                }
                # Set Title to "Disabled" to help identify current staff
                If (!($TestTitle -ceq "Disabled")) {
                    Set-ADUser -Identity $LoginName -Title "Disabled"
                    write-host "${LoginName} Title change from ${TestTitle} to: Disabled"
                }
            }

            # Remove groups if they are a member of any additional groups
            If ($TestMembership) {
                write-host "Removing groups for ${TestAccountName}"
                write-host
                # Remove All Villanova Groups
                Foreach($GroupName In $VillanovaGroups) {
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
