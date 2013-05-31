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
        # Check so we don't overwrite existing users
        If (!(Get-ADUser -Filter { SamAccountName -eq $LoginName })) {
            write-host $LoginName, " created for ", $FullName
            #write-host $LoginName
            #write-host $PreferredName
            #write-host $Surname
            #write-host $FullName
            #write-host $UserPrincipalName
            #write-host $UserCode
            #write-host $HomeDrive
        }
        Else {
            #write-host "Existing user", $LoginName
        }
            # Set user to confirm details
            $TestUser = Get-ADUser -Identity $LoginName

            If ($Position.Contains("Relief") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                #Get-ADUser $LoginName | Move-ADObject -TargetPath $ReliefTeacherPath
                write-host $LoginName, "Moved to RELIEF TEACHER"
            }
            ElseIf (($Position.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                #Get-ADUser $LoginName | Move-ADObject -TargetPath $TutorPath
                write-host $LoginName, "Staff moved to Music Tutor OU"
            }
            #Else {
            #    write-host $Position
            #}
            #if ($Telephone) {
            #    write-host $Telephone
            #}
        
    }

    ###################################
    ### Disable Staff who have left ###
    ###################################
    
    Else {
        If (Get-ADUser -Filter { ((SamAccountName -eq $LoginName) -and (Enabled -eq "True")) }) {
            write-host "USER HAS LEFT", $FullName
            write-host $Termination
            write-host
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

            $TestUser = Get-ADUser -Identity $LoginName
            $Description = (Get-ADUser -Identity $LoginName -Properties Description).Description
            if ($TestUser.distinguishedname.Contains($UserPath) -and (!($TestUser.distinguishedname.Contains($TeacherPath))) -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                If ($Description.Contains("Relief") -and (!($TestUser.distinguishedname.Contains($ReliefTeacherPath)))) {
                    #Get-ADUser $LoginName | Move-ADObject -TargetPath $ReliefTeacherPath
                    write-host $LoginName, "moved to Relief Teacher OU"
                }
                ElseIf (($Description.Contains("Tutor")) -and (!($TestUser.distinguishedname.Contains($TutorPath)))) {
                    #Get-ADUser $LoginName | Move-ADObject -TargetPath $TutorPath
                    write-host $LoginName, "Teacher moved to Music Tutor OU"
                }
                ElseIf ((!($TestUser.distinguishedname.Contains($TeacherPath)) -and ($Description.Contains("Teacher")) -and (!($Description.Contains("Tutor"))))) {
                    #Get-ADUser $LoginName | Move-ADObject -TargetPath $TeacherPath
                    write-host $LoginName, "moved to Teacher OU"
                    write-host $Description
                    $TEST = Get-ADUser -Identity $LoginName -Properties MemberOf
                    # Check Group Membership
                    if (!($TEST.memberof.Contains("CN=Teachers,OU=Security,OU=Curriculum Groups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"))) {
                        #Add-ADGroupMember -Identity Teachers -Member $LoginName
                        write-host "adding teachers group", $LoginName
                    }
                    if (!($TEST.memberof.Contains("CN=Teachers - All,OU=Distribution,OU=Curriculum Groups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"))) {
                        #Add-ADGroupMember -Identity "Teachers - All" -Member $LoginName
                        write-host "adding teachers - all ", $LoginName
                    }
                    write-host
                }
            }

            
        }
    }
}