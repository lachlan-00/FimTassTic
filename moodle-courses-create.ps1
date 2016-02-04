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
write-host "### Starting Moodle Group Creation Script"
write-host

###############
### GLOBALS ###
###############

$input = Import-CSV ".\csv\fim_MoodleCourses.csv" -Encoding UTF8
# Check for the length of the import so you don't overwrite the content
$inputCount = (Import-CSV ".\csv\fim_MoodleCourses.csv").count

# OU paths for different user types
$StudentPath = "OU=student,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"
$TeacherPath = "OU=teacher,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=villanova,DC=vnc,DC=qld,DC=edu,DC=au"

######################################
### Create / Edit / Disable groups ###
######################################
If ($inputCount -lt 50) {
    write-host "Not enough courses"
}
Else {
    write-host "Found ${inputCount} courses"
    foreach($line in $input) {

        # Read data for Timetabled Courses
        $courseid = $line.idnumber
        $classGroup = $line.stud_class
        $fullname = $line.fullname
        $category = $line.category_idnumber
        #if ($category -eq "TS") {
        #    write-host "TS Course Found"
        #    write-host "${courseid}"
        #    write-host "${fullname}"
        #}
    
        # Create colour groups for middle school
        if (($courseid -like "07-*") -and (($classGroup -ceq "BL") -or ($classGroup -ceq "GO") -or ($classGroup -ceq "GR") -or ($classGroup -ceq "MA") -or ($classGroup -ceq "RE") -or ($classGroup -ceq "WH"))) {

            $coursestudent = "07-${classGroup}-students"
            $courseteacher = "07-${classGroup}-teachers"

            ### Create Groups for year 7 student colour groups
            Try {
                #Check if the Group already exists
                $exists = Get-ADGroup $coursestudent
            }
            Catch {
                #Create the group if it doesn't exist
                $create = New-ADGroup -Name $coursestudent -GroupScope Global -Path $StudentPath -Description "07-${classGroup}"
                Write-Host "Student Group 07-${classGroup} created"
                Write-Host
            }
            ### Create Groups for year 7 teacher colour groups
            Try {
                #Check if the teacher Group already exists
                $exists = Get-ADGroup $courseteacher
            }
            Catch {
                #Create the group if it doesn't exist
                $create = New-ADGroup -Name $courseteacher -GroupScope Global -Path $TeacherPath -Description "07-${classGroup}"
                Write-Host "Teacher Group 07-${classGroup} created"
                Write-Host
            }
        }

        # Search for regular course names
        $coursestudent = "${courseid}-students"
        $courseteacher = "${courseid}-teachers"

        ### STUDENT AD Groups

        # REMOVED 2016-01-04 # if (!($category -ceq 'TS')) {
        if (($courseid -like "05-*") -or ($courseid -like "06-*")) {
            # nothing
            #write-host "${courseid} found; ignoring junior class"
        }
        else {
            ### Create Groups for students
            Try {
                #Check if the Group already exists
                $exists = Get-ADGroup $coursestudent
            }
            Catch {
                #Create the group if it doesn't exist
                $create = New-ADGroup -Name $coursestudent -GroupScope Global -Path $StudentPath -Description $courseid
                Write-Host "Student Group ${courseid} created"
            }
        }
        # REMOVED 2016-01-04 # }
    
        ### STAFF AD Groups

        # REMOVED 2016-01-04 # if (!($category -ceq 'TS')) {
        if (($courseid -like "05-*") -or ($courseid -like "06-*")) {
            #nothing
        }
        else {
            ### Create Groups for staff
            Try {
                #Check if the Group already exists
                $exists = Get-ADGroup $courseteacher
            }
            Catch {
                #Create the group if it doesn't exist
                $create = New-ADGroup -Name $courseteacher -GroupScope Global -Path $TeacherPath -Description $courseid
                Write-Host "TeacherGroup ${courseid} created"
            }
        }
        # REMOVED 2016-01-04 # }
    }
}

write-host
write-host "### Moodle Group Creation Script Finished"
write-host
