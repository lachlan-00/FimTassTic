###############################################################################
###                                                                         ###
###  Create Moodle Class Groups From TASS.web Data                          ###
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

write-host
write-host "### Starting Moodle Membership Add Script"
write-host

###############
### GLOBALS ###
###############

$members = Import-CSV ".\csv\fim_MoodleUsersCombined.csv" -Encoding UTF8
# Check for the length of the import so you don't overwrite the content
$memberCount = (Import-CSV ".\csv\fim_MoodleUsersCombined.csv").count

$CUSTOMmembers = Import-CSV ".\csv\MoodleUSERS_CUSTOM.csv" -Encoding UTF8

#####################################################
### Check group membership for staff and students ###
#####################################################
If ($memberCount -lt 500) {
    write-host "Not enough users enrolled in courses"
}
### Read data for enrolled students ###
Else {
    foreach($line in $members) {

        # Read data for Timetabled Courses
        $courseid = $line.CLASS_id
        $classGroup = $line.stud_class
        $usercode = (Get-Culture).TextInfo.ToUpper($line.USER_code)
        $title = $line.TITLE_line

        ### Add User to their regular course groups FIRST ###

        if ((!($courseid -like "05-*")) -and (!($courseid -like "06-*"))) {
            # Split membership depending on whether the user is a student or a teacher
            if ($title -ceq "Manager") {
                $courseid = "${courseid}-teachers"
            }
            elseif ($title -ceq "student") {
                $courseid = "${courseid}-students"
            }

            Try {
                #Add to course group
                Add-ADGroupMember -Identity "${courseid}" -Member $usercode
            }
            Catch {
                #Error with course
            }
            Finally {
                #end of line note
                Write-Host "Added ${usercode} to: ${courseid}"
            }
        }
        ### Add users to year 5 student groups

        if ($courseid -like "05-*") {

            # Split membership depending on whether the user is a student or a teacher
            if ($title -ceq "Manager") {
                $courseid = "05-Year-5-teachers"
            }
            elseif ($title -ceq "student") {
                $courseid = "05-Year-5-students"
            }

            Try {
                #Add to course group
                Add-ADGroupMember -Identity "${courseid}" -Member $usercode
            }
            Catch {
                #Error with course
            }
            Finally {
                #end of line note
                Write-Host "Added ${usercode} to: ${courseid}"
            }
        }

        ### Add users to year 6 student groups

        if ($courseid -like "06-*") {

            # Split membership depending on whether the user is a student or a teacher
            if ($title -ceq "Manager") {
                $courseid = "06-Year-6-teachers"
            }
            elseif ($title -ceq "student") {
                $courseid = "06-Year-6-students"
            }

            Try {
                #Add to course group
                Add-ADGroupMember -Identity "${courseid}" -Member $usercode
            }
            Catch {
                #Error with course
            }
            Finally {
                #end of line note
                Write-Host "Added ${usercode} to: ${courseid}"
            }
        }

        ### Add users to year 7 student colour groups
        if (($courseid -like "07-*") -and (($classGroup -ceq "BL") -or ($classGroup -ceq "GO") -or ($classGroup -ceq "GR") -or ($classGroup -ceq "MA") -or ($classGroup -ceq "RE") -or ($classGroup -ceq "WH"))) {

            # Split membership depending on whether the user is a student or a teacher
            if ($title -ceq "Manager") {
                $courseid = "07-${classGroup}-teachers"
            }
            elseif ($title -ceq "student") {
                $courseid = "07-${classGroup}-students"
            }

            Try {
                #Add to course group
                Add-ADGroupMember -Identity "${courseid}" -Member $usercode
            }
            Catch {
                #Error with course
            }
            Finally {
                #end of line note
                Write-Host "Added ${usercode} to: ${courseid}"
            }
        }
    }
}

Write-Host
Write-Host "TASS search finished"
Write-Host

Write-Host "Getting accounts from Custom File"
Write-Host

##############################################
### Add users to cources manually from csv ###
##############################################

foreach($line in $CUSTOMmembers) {

    # Read data for Custom Courses
    # Using this method the $courseid must include "-students" or "-teachers"
    $courseid = $line.CLASS_id
    $usercode = (Get-Culture).TextInfo.ToUpper($line.USER_code)

    Try {
        #Add user to course group
        Add-ADGroupMember -Identity "${courseid}" -Member $usercode
    }
    Catch {
        #Error with course
    }
    Finally {
        #end of line note
        Write-Host "Custom Add ${usercode} to: ${courseid}"
    }
}

write-host
write-host "### Moodle Membership Creation Script Finished"
write-host
