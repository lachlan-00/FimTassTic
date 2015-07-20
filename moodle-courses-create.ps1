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
$input = Import-CSV  ".\csv\fim_MoodleCourses.csv" -Encoding UTF8

write-host
write-host "### Starting Moodle Group Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$ClassPath = "OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au"
$StudentPath = "OU=student,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au"
$TeacherPath = "OU=teacher,OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au"

######################################
### Create / Edit / Disable groups ###
######################################

foreach($line in $input) {

    # The ID is Unique for all courses.
    $courseid = $line.idnumber
    $coursestudent = "${courseid}-students"
    $courseteacher = "${courseid}-teachers"
    $fullname = $line.fullname
    $category = $line.category_idnumber
    if (!($category -ceq 'TS')) {
        if ((!($courseid -like "05-*"))-or (!($courseid -like "06-*"))) {
            #nothing
        }
        else {
            ### Create Groups for students
            Try
            {
                #Check if the Group already exists
                $exists = Get-ADGroup $coursestudent
            }
            Catch
            {
                #Create the group if it doesn't exist
                $create = New-ADGroup -Name $coursestudent -GroupScope Global -Path $StudentPath -Description $courseid
                Write-Host "Student Group ${courseid} created"
            }
        }
    }
    if (!($category -ceq 'TS')) {
        if ((!($courseid -like "05-*"))-or (!($courseid -like "06-*"))) {
            #nothing
        }
        else {
            ### Create Groups for staff
            Try
            {
                #Check if the Group already exists
                $exists = Get-ADGroup $courseteacher
            }
            Catch
            {
                #Create the group if it doesn't exist
                $create = New-ADGroup -Name $courseteacher -GroupScope Global -Path $TeacherPath -Description $courseid
                Write-Host "TeacherGroup ${courseid} created"
            }
        }
    }
}

write-host
write-host "### Moodle Group Creation Script Finished"
write-host

#USER_code,CLASS_id
