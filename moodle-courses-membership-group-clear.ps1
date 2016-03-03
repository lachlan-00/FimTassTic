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
write-host "### Starting Moodle Membership Clear Script"
write-host

###############
### GLOBALS ###
###############

$input = Import-CSV "C:\DATA\csv\fim_MoodleCourses.csv" -Encoding UTF8
# Check for the length of the import so you don't overwrite the content
$inputCount = (Import-CSV "C:\DATA\csv\fim_MoodleCourses.csv").count

#######################################################
### Empty Groups so only current members will exist ###
#######################################################

If ($inputCount -lt 50) {
    write-host "Not enough courses"
}
### Read data for Timetabled Courses ###
Else {
    foreach($line in $input) {

        # CSV information
        $courseid = $line.idnumber
        $classGroup = $line.stud_class
        $fullname = $line.fullname
        $category = $line.category_idnumber

        ### Remove membership from regular Courses

        $coursestudent = "${courseid}-students"
        $courseteacher = "${courseid}-teachers"

        ### Students
        # REMOVED 2016-01-04 # if (!($category -ceq 'TS')) {
        Try {
            #Check group membership
            Get-ADGroupMember -Identity "${coursestudent}" | %{Remove-ADGroupMember -Identity "${coursestudent}" -Members $_ -Confirm:$false}
        }
        Catch {
            #Error with course
        }
        Finally {
            #end of line note
            Write-Host "Cleared: ${coursestudent}"
        }
        # REMOVED 2016-01-04 # }

        ### Teachers
        # REMOVED 2016-01-04 # if (!($category -ceq 'TS')) {
        Try {
            #Check group membership
            Get-ADGroupMember -Identity "${courseteacher}" | %{Remove-ADGroupMember -Identity "${courseteacher}" -Members $_ -Confirm:$false}
        }
        Catch {
            #Error with course
        }
        Finally {
            #end of line note
            Write-Host "Cleared: ${courseteacher}"
        }
        # REMOVED 2016-01-04 # }

        ### Remove membership from middle school colour groups

        # Create colour groups for middle school
        if (($courseid -like "07-*") -and (($classGroup -ceq "BL") -or ($classGroup -ceq "GO") -or ($classGroup -ceq "GR") -or ($classGroup -ceq "MA") -or ($classGroup -ceq "RE") -or ($classGroup -ceq "WH"))) {
            $coursestudent = "07-${classGroup}-students"
            $courseteacher = "07-${classGroup}-teachers"

            ### for year 7 student colour groups

            # Students
            Try {
                #Check group membership
                Get-ADGroupMember -Identity "${coursestudent}" | %{Remove-ADGroupMember -Identity "${coursestudent}" -Members $_ -Confirm:$false}
            }
            Catch {
                #Error with course
            }
            Finally {
                #end of line note
                Write-Host "Cleared: ${coursestudent}"
            }

            # Teachers
            Try {
                #Check group membership
                Get-ADGroupMember -Identity "${courseteacher}" | %{Remove-ADGroupMember -Identity "${courseteacher}" -Members $_ -Confirm:$false}
            }
            Catch {
                #Error with course
            }
            Finally {
                #end of line note
                Write-Host "Cleared: ${courseteacher}"
            }
        }
        Write-Host
    }

    #### Remove membership from manual courses for junior school

    # Students
    Try {
        #Check group membership
        Get-ADGroupMember -Identity "05-Year-5-students" | %{Remove-ADGroupMember -Identity "05-Year-5-students" -Members $_ -Confirm:$false}
    }
    Catch {
    }
    Try {
        #Check group membership
        Get-ADGroupMember -Identity "05-Year-5-teachers" | %{Remove-ADGroupMember -Identity "05-Year-5-teachers" -Members $_ -Confirm:$false}
    }
    Catch {
    }

    # Teachers
    Try {
        #Check group membership
        Get-ADGroupMember -Identity "06-Year-6-students" | %{Remove-ADGroupMember -Identity "06-Year-6-students" -Members $_ -Confirm:$false}
    }
    Catch {
    }
    Try {
        #Check group membership
        Get-ADGroupMember -Identity "06-Year-6-teachers" | %{Remove-ADGroupMember -Identity "06-Year-6-teachers" -Members $_ -Confirm:$false}
    }
    Catch {
    }

    Write-Host
    Write-Host "### Active Directory Class Groups emptied"
    Write-Host
}
