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
$input = Import-CSV  .\csv\telemf.csv
$teacherinput = Import-CSV  .\csv\teacher.csv

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$UserPath = "OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TeacherPath = "OU=Teachers,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$NonTeacherPath = "OU=Teachers,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$ReliefTeacherPath = "OU=Relief and Preservice Teachers,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TutorPath = "OU=Tutors,OU=Staff,OU=Curriculum Users,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$DisablePath = "OU=staff,OU=Disabled,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
# Get membership for group Membership Tests
###$TestDomainUser = Get-ADGroupMember -Identity "Domain Users"
$TestStaff = Get-ADGroupMember -Identity "Staff"
$TestAllStaff = Get-ADGroupMember -Identity "All Staff"
$TestTeachers = Get-ADGroupMember -Identity "Teachers"
$TestAllTeachers = Get-ADGroupMember -Identity "Teachers - All"

##############################################
### Create / Edit / Disable Staff accounts ###
##############################################

foreach($line in $input) {

    # Lower set lower case because powershell ignores uppercase words
    $PreferredName = (Get-Culture).TextInfo.ToLower($line.prefer_name_text.Trim())
    $GivenName = (Get-Culture).TextInfo.ToLower($line.given_names_text.Trim())
    $Surname = (Get-Culture).TextInfo.ToLower($line.surname_text.Trim())
    If (($line.position_title.Trim() -ne $null) -and ($line.position_title.Trim() -ne '')) {
        $Position = (Get-Culture).TextInfo.ToLower($line.position_title.Trim())
    }
    Elseif (($line.position_text.Trim() -ne $null) -and ($line.position_text.Trim() -ne '')) {
        $Position = (Get-Culture).TextInfo.ToLower($line.position_text.Trim())
    }

    # Set Login and display names/text
    $FullName =  (Get-Culture).TextInfo.ToTitleCase($PreferredName + " " + $Surname)
    $LoginName = (Get-Culture).TextInfo.ToUpper($line.emp_code.Trim())
    $UserPrincipalName = $LoginName + "@villanova.vnc.qld.edu.au"
    $UserCode = $LoginName
    $HomeDrive = "\\vncfs01\staffdata$\" + $LoginName
    $PreferredName = (Get-Culture).TextInfo.ToTitleCase($PreferredName)
    $Surname = (Get-Culture).TextInfo.ToTitleCase($Surname)
    $Position = (Get-Culture).TextInfo.ToTitleCase($Position.Trim())

    # pull remaining details
    $Termination = $line.term_date.Trim()
    $Telephone = $line.phone_w_text.Trim()
    If ($Telephone.length -le 1) {
        $Telephone = $null
    }

    ######################################
    ### Create / Modify Staff Accounts ###
    ######################################

    If ($Termination.length -eq 0) {
        # create basic user if you can't find one
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            $TempLogin = (Get-Culture).TextInfo.ToLower($LoginName)
            New-ADUser -SamAccountName $TempLogin -Name $FullName -AccountPassword (ConvertTo-SecureString -AsPlainText "Abc123" -Force) -Enabled $true -Path $UserPath -DisplayName $FullName -GivenName $PreferredName -Surname $Surname -UserPrincipalName $UserPrincipalName -ChangePasswordAtLogon $True -homedrive "H" -homedirectory $HomeDrive
            Set-ADUser -Identity $LoginName -Description $Position -Office $Position -Title $Position
            write-host $LoginName, " created for ", $FullName
        }

        # set additional user details if the user exists
        If (Get-ADUser -Filter { SamAccountName -eq $LoginName }) {
            # Enable use if disabled
            If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "False")) }) {
                Set-ADUser -Identity $LoginName -Enabled $true
            }

            # Set updateable object values
            Set-ADUser -Identity $LoginName -Description $Position
            ########Set-ADUser -Identity $LoginName -Description $Position -Office $Position -Title $Position

            # Add Telephone number if there is one
            if ($Telephone) {
                Set-ADUser -Identity $LoginName -OfficePhone $Telephone
            }
            
            # Set user to confirm details
            $TestUser = Get-ADUser -Identity $LoginName

            # Move user to the default OU if not already there
            if ($TestUser.distinguishedname.Contains($DisablePath)) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $UserPath
                write-host $LoginName "moved out of Disabled OU"
            }
            ElseIf ($Position.Contains("Relief") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $ReliefTeacherPath
                write-host $LoginName "moved to Relief Teacher OU"
            }
            ElseIf (($Position.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $TutorPath
                write-host $LoginName "moved to Music Tutor OU"
            }
            
            # Check Group Membership
            #if (!($TestDomainUser.name.contains($TestUser.name))) {
            #    try {
            #            Add-ADGroupMember -Identity "Domain Users" -Member $LoginName
            #            write-host $LoginName "added Domain Users"
            #        }
            #        catch {}
            #        finally {}
            #}
            if (!($TestStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "Staff" -Member $LoginName
                        write-host $LoginName "added Staff"
            }
            if (!($TestAllStaff.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "All Staff" -Member $LoginName
                        write-host $LoginName "added allstaff"
            }
        }
    }

    ###################################
    ### Disable Staff who have left ###
    ###################################
    
    Else {

        # Disable users with a termination date if they are still enabled
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }) {
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
            If ($DATE -ge $Termination) {
                write-host "DISABLING ACCOUNT, '$($LoginName)'"

                # Set user to confirm details
                $TestUser = Get-ADUser -Identity $LoginName

                if (!($TestUser.distinguishedname.Contains($DisablePath))) {
                    # Move to disabled user OU if not already there
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $DisablePath
                }

                # Check Group Membership
                #if ($TestDomainUser.name.contains($TestUser.name)) {
                #    Remove-ADGroupMember -Force -Identity "Domain users" -Member $LoginName
                #    write-host $LoginName "REMOVED Domain Users"
                #}
                if ($TestStaff.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "Staff" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED Staff"
                }
                if ($TestAllStaff.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "All Staff" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED allstaff"
                }
                if ($TestTeachers.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "Teachers" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED Teachers"
                }
                if ($TestAllTeachers.name.contains($TestUser.name)) {
                    Remove-ADGroupMember -Identity "Teachers - All" -Member $LoginName -Confirm:$false
                    write-host $LoginName "REMOVED Teachers - All"
                }
            
                # Disable The account
                Set-ADUser -Identity $LoginName -Enabled $false
            }
        }
    }
}

###################################################################
### Create / Edit Teacher Info for Staff with existing accounts ###
###################################################################

foreach($line in $teacherinput)
{
    $LoginName = (Get-Culture).TextInfo.ToLower($line.emp_code.Trim())

    If (($LoginName -ne $null) -and ($LoginName -ne '')) {
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }) {

            # Set user to confirm details
            $TestUser = Get-ADUser -Identity $LoginName
            $Description = (Get-ADUser -Identity $LoginName -Properties Description).Description
            # Move to Teacher OU if not already there
            if ($TestUser.distinguishedname.Contains($UserPath) -and (!($TestUser.distinguishedname.Contains($TeacherPath))) -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                If ($Description.Contains("Relief") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                Get-ADUser $LoginName | Move-ADObject -TargetPath $ReliefTeacherPath
                write-host $LoginName "moved to Relief Teacher OU"
                }
                ElseIf (($Description.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $TutorPath
                    write-host $LoginName "moved to Music Tutor OU"
                }
                ElseIf (!($TestUser.distinguishedname.Contains($TeacherPath)) -and (!($Description.Contains("Tutor")))) {
                    Get-ADUser $LoginName | Move-ADObject -TargetPath $TeacherPath
                    write-host $LoginName "moved to Teacher OU"
                    # Check Group Membership
                    if (!($TestTeachers.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "Teachers" -Member $LoginName
                        write-host $LoginName "ADDED Teachers"
                    }
                    if (!($TestAllTeachers.name.contains($TestUser.name))) {
                        Add-ADGroupMember -Identity "Teachers - All" -Member $LoginName
                        write-host $LoginName "ADDED Teachers - All"
                    }
                }
            }  
        }
    }
}