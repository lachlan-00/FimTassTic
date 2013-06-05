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
$input = Import-CSV .\csv\student.csv

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
###$TestDomainUser = Get-ADGroupMember -Identity "Domain Users"
$StudentGroup = Get-ADGroupMember -Identity "Students"
$5Group = Get-ADGroupMember -Identity "S-G_Group H"
$6Group = Get-ADGroupMember -Identity "S-G_Group G"
$7Group = Get-ADGroupMember -Identity "S-G_Group F"
$8Group = Get-ADGroupMember -Identity "S-G_Group E"
$9Group = Get-ADGroupMember -Identity "S-G_Group D"
$10Group = Get-ADGroupMember -Identity "S-G_Group C"
$11Group = Get-ADGroupMember -Identity "S-G_Group B"
$12Group = Get-ADGroupMember -Identity "S-G_Group A"

###############################################
### Create / Edit /Disable student accounts ###
###############################################

foreach($line in $input) {
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

    # pull remaining details
    $PreferredName = (Get-Culture).TextInfo.ToLower($line.preferred_name.Trim())
    $GivenName = (Get-Culture).TextInfo.ToLower($line.given_name.Trim())
    $Surname = (Get-Culture).TextInfo.ToLower($line.surname.Trim())
    $Position = 'Year', $YearGroup.Trim()
    $FullName =  (Get-Culture).TextInfo.ToTitleCase($PreferredName + " " + $Surname)
    if (((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "").length -gt 3)
        {
        $LoginName = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "").substring(0,4) + $PreferredName.substring(0,2)
        $Test0 = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "").substring(0,4) + $GivenName.substring(0,2)
        }
    else
        {
        $LoginName = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "") + $PreferredName.substring(0,2)
        $Test0 = $YearIdent + '-' + ((($Surname -replace "\s+", "") -replace "'", "") -replace "-", "") + $GivenName.substring(0,2)
        }
    $UserPrincipalName = $LoginName + "@villanova.vnc.qld.edu.au"
    $UserCode = (Get-Culture).TextInfo.ToUpper($line.stud_code.Trim())
    $HomeDrive = "\\vncfs01\studentdata$\" + $YearIdent + "\" + $LoginName
    $PreferredName = (Get-Culture).TextInfo.ToTitleCase($PreferredName)
    $Surname = (Get-Culture).TextInfo.ToTitleCase($Surname)
    $Termination = $line.dol.Trim()

    # Check for given name before testing  year identities
    If ($LoginName -notcontains $Test0) {
        If (Get-ADUser -Filter { SamAccountName -eq $Test0 }) {
            #write-host "missing login", $LoginName
            $LoginName = $Test0
        }
    }
    #check against old login name style (helps for end of year rollover)
    $TestA = ("a" + $LoginName.Substring(1))
    $TestB = ("b" + $LoginName.Substring(1))
    $TestC = ("c" + $LoginName.Substring(1))
    $TestD = ("d" + $LoginName.Substring(1))
    $TestE = ("e" + $LoginName.Substring(1))
    $TestF = ("f" + $LoginName.Substring(1))
    $TestG = ("g" + $LoginName.Substring(1))
    $TestH = ("h" + $LoginName.Substring(1))
    If ($LoginName -notcontains $TestA) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestA }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestA
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\A\" + $LoginName
            $Position = 'Year 12'
        }
    }
    If ($LoginName -notcontains $TestB) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestB }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestB
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\B\" + $LoginName
            $Position = 'Year 11'
        }
    }
    If ($LoginName -notcontains $TestC) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestC }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestC
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\C\" + $LoginName
            $Position = 'Year 10'
        }
    }
    If ($LoginName -notcontains $TestD) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestD }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestD
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\D\" + $LoginName
            $Position = 'Year 9'
        }
    }
    If ($LoginName -notcontains $TestE) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestE }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestE
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\E\" + $LoginName
            $Position = 'Year 8'
        }
    }
    If ($LoginName -notcontains $TestF) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestF }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestF
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\F\" + $LoginName
            $Position = 'Year 7'
        }
    }
    If ($LoginName -notcontains $TestG) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestG }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestG
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\G\" + $LoginName
            $Position = 'Year 6'
        }
    }
    If ($LoginName -notcontains $TestH) {
        If (Get-ADUser -Filter { SamAccountName -eq $TestH }) {
            #write-host "missing login", $LoginName
            $LoginName = $TestH
            #write-host "missing login", $LoginName
            $HomeDrive = "\\vncfs01\studentdata$\H\" + $LoginName
            $Position ='Year 5'
        }
    }

    ########################################
    ### Create / Modify Student Accounts ###
    ########################################

    If ($Termination.length -eq 0) {
        # Check for existing users before creating new ones.
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            write-host "missing login", $LoginName
            $LoginName = $UserCode
            $UserPrincipalName = $LoginName + "@villanova.vnc.qld.edu.au"
            $HomeDrive = "\\vncfs01\studentdata$\" + $YearIdent + "\" + $LoginName
            $Position = 'Year', $YearGroup
        }
        
        # Create basic user if you can't find one
        If (!(Get-ADUser -Filter { (SamAccountName -eq $LoginName)})) {
            #New-ADUser -SamAccountName $LoginName -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "Abc123" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
            write-host "NEW USER", $($LoginName)
        }
        
        # set additional user details if the user exists
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode))}) {
            # Enable use if disabled
            If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode) -and (Enabled -eq "False")) }) {
                #Set-ADUser -Identity $LoginName -Enabled $true
                write-host "RE-ENABLING", $($LoginName)
            }

            # Set updateable object values
            #Set-ADUser -Identity $LoginName -Enabled $true -Description $UserCode -Office $YearGroup -Title "Student"
            
            # Set user to confirm details
            $TestUser = Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode)) }

            # Move user to the default OU if not already there
            if (!($TestUser.distinguishedname.Contains($UserPath))) {
                #Get-ADUser $LoginName | Move-ADObject -TargetPath $UserPath
                write-host $LoginName "Taking From:" $TestUser.distinguishedname
                write-host "Moving To:" $UserPath
            }

            # Check Group Membership
            #if (!($TestDomainUser.name.contains($TestUser.name))) {
            #    #Add-ADGroupMember -Identity "Domain Users" -Member $LoginName
            #    write-host $LoginName "added Domain Users"
            #}
            if (!($StudentGroup.name.contains($TestUser.name))) {
                #Add-ADGroupMember -Identity Students -Member $LoginName
                write-host $LoginName "added Students Group"
            }
            # Remove groups for other grades and add the correct grade
            IF ($YearGroup -eq "5") {
                # Add Correct Year Level
                if (!($5Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $5Name -Member $LoginName
                    write-host $LoginName "added 5"
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "6") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($6Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $6Name -Member $LoginName
                    write-host $LoginName "added 6"
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "7") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($7Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $7Name -Member $LoginName
                    write-host $LoginName "added 7"
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "8") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($8Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $8Name -Member $LoginName
                    write-host $LoginName "added 8"
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "9") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($9Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $9Name -Member $LoginName
                    write-host $LoginName "added 9"
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "10") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($10Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $10Name -Member $LoginName
                    write-host $LoginName "added 10"
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "11") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($11Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $11Name -Member $LoginName
                    write-host $LoginName "added 11"
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                }
            }
            IF ($YearGroup -eq "12") {
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                }
                # Add Correct Year Level
                if (!($12Group.name.contains($TestUser.name))) {
                    #Add-ADGroupMember -Identity $12Name -Member $LoginName
                    write-host $LoginName "added 12"
                }
            }
        }
    }

    ######################################
    ### Disable Students who have left ###
    ######################################

    Else {
        # Disable users with a termination date if they are still enabled
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode) -and (Enabled -eq "True")) }) {
            $YEAR = [string](Get-Date).Year
            $MONTH = [string](Get-Date).Month
            If ($MONTH.length -eq 1) {
                $MONTH = "0${MONTH}"
            }
            $DAY = [string](Get-Date).Day
            If ($DAY.length -eq 1) {
                $DAY = "0${DAY}"
            }

            $DATE = "${YEAR}-${MONTH}-${DAY}"
            $DATE = $DATE, "00:00:00"
            If ($DATE -gt $Termination) {
                write-host "DISABLING ACCOUNT, '$($LoginName)'"

                # Set user to confirm details
                $TestUser = Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Description -eq $UserCode)) }

                if (!($TestUser.distinguishedname.Contains($DisablePath))) {
                    # Move to disabled user OU if not already there
                    #Get-ADUser $LoginName | Move-ADObject -TargetPath $DisablePath
                    write-host $LoginName "MOVED to Disabled OU"
                }

                # Check Group Membership
                #if ($TestDomainUser.name.contains($TestUser.name)) {
                #    #Remove-ADGroupMember -Force -Identity "Domain users" -Member $LoginName
                #    write-host $LoginName "REMOVED Domain Users"
                #}
                if ($StudentGroup.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity "Students" -Member $LoginName
                    write-host $LoginName "REMOVED Students"
                }
                if ($5Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $5Name -Member $LoginName
                    write-host $LoginName "REMOVED 5"
                }
                if ($6Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $6Name -Member $LoginName
                    write-host $LoginName "REMOVED 6"
                }
                if ($7Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $7Name -Member $LoginName
                    write-host $LoginName "REMOVED 7"
                }
                if ($8Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $8Name -Member $LoginName
                    write-host $LoginName "REMOVED 8"
                }
                if ($9Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $9Name -Member $LoginName
                    write-host $LoginName "REMOVED 9"
                }
                if ($10Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $10Name -Member $LoginName
                    write-host $LoginName "REMOVED 10"
                }
                if ($11Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $11Name -Member $LoginName
                    write-host $LoginName "REMOVED 11"
                }
                if ($12Group.name.contains($TestUser.name)) {
                    #Remove-ADGroupMember -Force -Identity $12Name -Member $LoginName
                    write-host $LoginName "REMOVED 12"
                }

                # Disable The account
                #Set-ADUser -Identity $LoginName -Enabled $false
            }
        }
    }
}
