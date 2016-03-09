###############################################################################
###                                                                         ###
###  Create Staff and Student Email Accounts From TASS.web Data             ###
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

### Active Directory
import-module activedirectory

### On Premise Exchange
If (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
    #write-host "Exchange is imported"
}
Else {
    #write-host "importing exchange module"
    . 'D:\Program Files\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto
}

### Office 365 Connector (check for Add-RecipientPermission)
###
### http://www.adminarsenal.com/admin-arsenal-blog/secure-password-with-powershell-encrypting-credentials-part-1/
###
If (Get-Command Add-RecipientPermission -ErrorAction SilentlyContinue) {
    #write-host "365 is imported"
}
Else {
    $pass = cat C:\DATA\365securestring.txt | convertto-securestring
    $mycred = new-object -typename System.Management.Automation.PSCredential -argumentlist "generic.admin@vnc.qld.edu.au",$pass
    Import-Module MSOnline
    $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection -Credential $mycred
    Import-PSSession $O365Session
    Connect-MsolService -Credential $mycred
}

### Input CSV's
$input = Import-CSV "C:\DATA\csv\fim_staffALL.csv" -Encoding UTF8
$inputcount = (Import-CSV "C:\DATA\csv\fim_staffALL.csv" -Encoding UTF8 | Measure-Object).Count
$StudentInput = Import-CSV "C:\DATA\csv\fim_student.csv" -Encoding UTF8
$StudentInputcount = (Import-CSV "C:\DATA\csv\fim_student.csv" -Encoding UTF8 | Measure-Object).Count
$enrolledinput = Import-CSV "C:\DATA\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8
$enrolledinputcount = (Import-CSV "C:\DATA\csv\fim_enrolled_students-ALL.csv" -Encoding UTF8 | Measure-Object).Count

# Deny Access group is used to remove problem students from important services
$DenyAccessGroup = "CN=S-G_Deny-Access,OU=security,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$DenyAccessGroupMembers = Get-ADGroupMember -Identity $DenyAccessGroup

$userdomain = "VILLANOVA"

write-host "### Starting Mailbox Creation Script"
write-host

# Set the date to match the termination date for each type of user
$YEAR = [string](Get-Date).Year
$MONTH = [string](Get-Date).Month
If ($MONTH.length -eq 1) {
    $MONTH = "0${MONTH}"
}
$DAY = [string](Get-Date).Day
If ($DAY.length -eq 1) {
    $DAY = "0${DAY}"
}
# Student date format
$DATE = "${YEAR}/${MONTH}/${DAY}"
# Staff date format
$FULLDATE = "${DATE} 00:00:00"
$LogDate = "${YEAR}-${MONTH}-${DAY}"

# check log path
If (!(Test-Path "C:\DATA\log")) {
    mkdir "C:\DATA\log"
}

# set log file
$LogFile = "C:\DATA\log\Email-Creation-${LogDate}.log"
$LogContents = @()

#####################################
### Set 365 Location to Australia ###
#####################################

write-host "### Checking all users are assigned to the AU 365 Location"
write-host

Get-MsolUser -all | Where-Object { $_.isLicensed -ne "TRUE" }| Set-MsolUser -UsageLocation AU

##############################################
### Create Staff 365 Accounts and licenses ###
##############################################

write-host
write-host "### Parsing Staff File"
write-host

$tmpcount = 0
$lastprogress = $NULL

foreach($line in $Input) {
    $progress = ((($tmpcount / $Inputcount) * 100) -as [int]) -as [string]
    If (((((($tmpcount / $Inputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
        Write-Host "Progress: ${progress}%"
    }
    $tmpcount = $tmpcount + 1
    $lastprogress = $progress

    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())

    # Check for staff who have left
    $Termination = $line.term_date.Trim()

    # Set user to confirm details
    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
    $TestEnabled = $TestUser.Enabled
    $TestDescription = $TestUser.Description
    $SchoolEmail = $TestUser.UserPrincipalName

    ### Process Current Users ###
    If ($Termination.length -eq 0) {
        # get location and check for existing license
        $tempUser = Get-MsolUser -UserPrincipalName $SchoolEmail -ErrorAction SilentlyContinue
        $testusagelocation = $tempUser.UsageLocation
        $test365license = $tempUser.isLicensed
        $testblock = $tempUser.BlockCredential
        $testpassexpiry = $tempUser.PasswordNeverExpires

        If (($tempUser) -and (!($testpassexpiry))) {
            write-host "Disabling password expiry for: ${LoginName}"
            Set-MsolUser -UserPrincipalName $SchoolEmail -PasswordNeverExpires $true
        }

        # Check for Deny-Access members
        If (!($DenyAccessGroupMembers)) {
            #Empty Group
        }
        # Remove User from groups if they are a member of Deny Access
        ElseIf ($DenyAccessGroupMembers.SamAccountName.contains($LoginName)) {
            # User is denied from accessing their account
            If (!($testblock)) {
                write-host "blocking user account"
                Set-MsolUser -UserPrincipalName $SchoolEmail -BlockCredential $true
            }
        }
        ElseIf (($TestUser) -and ($SchoolEmail)) {

            # Only Add licenses when users have a location set to AU
            If (!($test365license) -and ($testusagelocation -ceq "AU")) {
                write-host "Missing 365 License for ${LoginName}... Adding."
                Get-MsolUser -UserPrincipalName $SchoolEmail| Where-Object { $_.isLicensed -ne "TRUE" }| Set-MsolUserLicense -AddLicenses "vnc4:AAD_PREMIUM"
                $LogContents += "Added Faculty 365 License for: ${LoginName}" #| Out-File $LogFile -Append
            }
            If ($testblock) {
                write-host "unblocking user account"
                Set-MsolUser -UserPrincipalName $SchoolEmail -BlockCredential $false
            }
        }
    }
}

write-host
write-host "### Parsing Student File"
write-host

###############################################
### Create / Disable Student Email Accounts ###
###############################################

$tmpcount = 0
$lastprogress = $NULL

foreach($line in $StudentInput) {
    $progress = ((($tmpcount / $StudentInputcount) * 100) -as [int]) -as [string]
    If (((((($tmpcount / $StudentInputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
        Write-Host "Progress: ${progress}%"
    }
    $tmpcount = $tmpcount + 1
    $lastprogress = $progress

    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())

    # Check for students who have left
    $Termination = $line.dol.Trim()

    # Set user to confirm details
    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
    $TestEnabled = $TestUser.Enabled
    $TestDescription = $TestUser.Description
    $SchoolEmail = $TestUser.UserPrincipalName

    ### Process Current Users ###
    If ($Termination.length -eq 0) {
        # get location and check for existing license
        $tempUser = Get-MsolUser -UserPrincipalName $SchoolEmail -ErrorAction SilentlyContinue
        $testusagelocation = $tempUser.UsageLocation
        $test365license = $tempUser.isLicensed
        $testblock = $tempUser.BlockCredential
        $testpassexpiry = $tempUser.PasswordNeverExpires

        If (($tempUser) -and (!($testpassexpiry))) {
            write-host "Disabling password expiry for: ${LoginName}"
            Set-MsolUser -UserPrincipalName $SchoolEmail -PasswordNeverExpires $true
        }

        # Check for Deny-Access members
        If (!($DenyAccessGroupMembers)) {
            #Empty Group
        }
        # Remove User from groups if they are a member of Deny Access
        ElseIf ($DenyAccessGroupMembers.SamAccountName.contains($LoginName)) {
            # User is denied from accessing their account
            If (!($testblock)) {
                write-host "blocking user account ${LoginName}"
                Set-MsolUser -UserPrincipalName $SchoolEmail -BlockCredential $true
            }
        }
        ElseIf ($TestUser) {
            # Check for 365 mail account
            Try {
                $testmailenabled = Get-RemoteMailbox -Identity $LoginName
            }
            Catch {
                $testmailenabled = $null
            }
            $testusagelocation = (Get-MsolUser -UserPrincipalName "${LoginName}@vnc.qld.edu.au").UsageLocation
            $test365license = (Get-MsolUser -UserPrincipalName "${LoginName}@vnc.qld.edu.au").isLicensed

            # Only work on users that have a 365 mail account
            If ($testmailenabled) {
                #write-host "365 Mailbox Found for ${LoginName}"
                # Only Add licenses when users have a location set to AU
                If (!($test365license) -and ($testusagelocation -ceq "AU")) {
                    write-host "Missing 365 License for ${LoginName}... Adding."
                    Get-MsolUser -UserPrincipalName "${LoginName}@vnc.qld.edu.au"| Where-Object { $_.isLicensed -ne "TRUE" }| Set-MsolUserLicense -AddLicenses "vnc4:STANDARDWOFFPACK_IW_STUDENT"
                    $LogContents += "Added Student 365 License for: ${LoginName}" #| Out-File $LogFile -Append
                }
            }
            Else {
                write-host "365 Mailbox Not Found for ${LoginName}... Creating"
                Enable-RemoteMailbox ${LoginName} -alias ${LoginName} -RemoteRoutingAddress "${LoginName}@VNC4.mail.onmicrosoft.com"
                $LogContents += "Created mailbox for ${LoginName}" #| Out-File $LogFile -Append
            }
            If ($testblock) {
                write-host "unblocking user account"
                Set-MsolUser -UserPrincipalName $SchoolEmail -BlockCredential $false
            }
        }
    }
}

write-host "### Parsing Future Student File"
write-host

############################################
### Create FUTURE Student Email Accounts ###
############################################

$tmpcount = 0
$lastprogress = $NULL

foreach($line in $enrolledinput) {
    $progress = ((($tmpcount / $enrolledinputcount) * 100) -as [int]) -as [string]
    If (((((($tmpcount / $enrolledinputcount) * 100) -as [int]) / 10) -is [int]) -and (!(($progress) -eq ($lastprogress)))) {
        Write-Host "Progress: ${progress}%"
    }
    $tmpcount = $tmpcount + 1
    $lastprogress = $progress

    # Get login name for processing
    $LoginName = (Get-Culture).TextInfo.ToLower($line.stud_code.Trim())

    # Set user to confirm details
    $TestUser = (Get-ADUser -Filter { (SamAccountName -eq $LoginName) } -Properties *)
    $TestEnabled = $TestUser.Enabled
    $TestMail = $TestUser.mail

    ### Process Current Users ###
    If (($TestUser) -and ($TestEnabled)) {
            # Check for 365 mail account
            Try {
                $testmailenabled = Get-RemoteMailbox -Identity $LoginName
            }
            Catch {
                $testmailenabled = $null
            }
            $testusagelocation = (Get-MsolUser -UserPrincipalName "${LoginName}@vnc.qld.edu.au").UsageLocation
            $test365license = (Get-MsolUser -UserPrincipalName "${LoginName}@vnc.qld.edu.au").isLicensed

            # Only work on users that have a 365 mail account
            If ($testmailenabled) {
                #write-host "365 Mailbox Found for ${LoginName}"
                # Only Add licenses when users have a location set to AU
                If (!($test365license) -and ($testusagelocation -ceq "AU")) {
                    write-host "Missing 365 License for ${LoginName}... Adding."
                    Get-MsolUser -UserPrincipalName "${LoginName}@vnc.qld.edu.au"| Where-Object { $_.isLicensed -ne "TRUE" }| Set-MsolUserLicense -AddLicenses "vnc4:STANDARDWOFFPACK_IW_STUDENT"
                    $LogContents += "Added Student 365 License for: ${LoginName}" #| Out-File $LogFile -Append
                }
            }
            Else {
                write-host "365 Mailbox Not Found for ${LoginName}... Creating"
                Enable-RemoteMailbox ${LoginName} -alias ${LoginName} -RemoteRoutingAddress "${LoginName}@VNC4.mail.onmicrosoft.com"
                $LogContents += "Created mailbox for ${LoginName}" #| Out-File $LogFile -Append
            }
            If ($test365license) {
                #write-host "365 already licensed for ${LoginName}"
            }
    }
}

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
    Write-Host "No Changes occurred"
}

write-host
write-host "### Mailbox Creation Script Finished"
write-host
