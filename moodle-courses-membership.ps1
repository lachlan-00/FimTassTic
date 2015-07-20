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
$members = Import-CSV  ".\csv\fim_MoodleALLUSERS.csv" -Encoding UTF8
$CUSTOMmembers = Import-CSV  ".\csv\MoodleUSERS_CUSTOM.csv" -Encoding UTF8

write-host
write-host "### Starting Moodle Membership Creation Script"
write-host

###############
### GLOBALS ###
###############

# OU paths for differnt user types
$ClassPath = "OU=ClassEnrolment,OU=moodle,OU=UserGroups,DC=example,DC=com,DC=au"

#######################################################
### Empty Groups so only current members will exist ###
#######################################################

foreach($line in $input) {

    # The ID is Unique for all courses.
    $courseid = $line.idnumber
    $coursestudent = "${courseid}-students"
    $courseteacher = "${courseid}-teachers"
    $fullname = $line.fullname
    $category = $line.category_idnumber
    if (!($category -ceq 'TS')) {
        Try
        {
            #Check group membership
            Get-ADGroupMember -Identity "${coursestudent}" | %{Remove-ADGroupMember -Identity "${coursestudent}" -Members $_ -Confirm:$false}
        }
        Catch
        {
        }
    }
    if (!($category -ceq 'TS')) {
        Try
        {
            #Check group membership
            Get-ADGroupMember -Identity "${courseteacher}" | %{Remove-ADGroupMember -Identity "${courseteacher}" -Members $_ -Confirm:$false}
        }
        Catch
        {
        }
    }
    # Remove membership from manual courses for junior school
    Try
    {
        #Check group membership
        Get-ADGroupMember -Identity "05-Year-5-students" | %{Remove-ADGroupMember -Identity "05-Year-5-students" -Members $_ -Confirm:$false}
    }
    Catch
    {
    }
    Try
    {
        #Check group membership
        Get-ADGroupMember -Identity "05-Year-5-teachers" | %{Remove-ADGroupMember -Identity "05-Year-5-teachers" -Members $_ -Confirm:$false}
    }
    Catch
    {
    }

    # Remove membership from manual courses for junior school
    Try
    {
        #Check group membership
        Get-ADGroupMember -Identity "06-Year-6-students" | %{Remove-ADGroupMember -Identity "06-Year-6-students" -Members $_ -Confirm:$false}
    }
    Catch
    {
    }
    Try
    {
        #Check group membership
        Get-ADGroupMember -Identity "06-Year-6-teachers" | %{Remove-ADGroupMember -Identity "06-Year-6-teachers" -Members $_ -Confirm:$false}
    }
    Catch
    {
    }
}

Write-Host "All Class Groups emptied"
Write-Host

Write-Host "Adding membership from TASS"
Write-Host

#################################################
### Sort membership by 'Manager' or 'student' ###
#################################################

foreach($line in $members) {

    # The ID is Unique for all courses.
    $courseid = $line.CLASS_id
    $usercode = (Get-Culture).TextInfo.ToUpper($line.USER_code)
    $title = $line.TITLE_line

    ### Read data for Timetabled Courses ###

    # Split membership depending on whether the user is a student or a teacher
    if ($title -ceq "Manager") {
        $courseid = "${courseid}-teachers"
    }
    elseif ($title -ceq "student") {
        $courseid = "${courseid}-students"
    }

    ### Add User to their groups ###

    # course names are set in SQL so it's just a simple add
    Try {
        #Add user to course group
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

Write-Host
Write-Host "TASS search finished"
Write-Host

Write-Host "Getting accounts from Custom File"
Write-Host

##############################################
### Add users to cources manually from csv ###
##############################################

foreach($line in $CUSTOMmembers) {

    #Using this method the $courseid must include "-students" or "-teachers"
    $courseid = $line.CLASS_id
    $usercode = (Get-Culture).TextInfo.ToUpper($line.USER_code)

    ### Read data for Custom Course membership ###

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

