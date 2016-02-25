------------------------------------------------------------------
FimTassTic: Automatic User Management using CSV data from TASS.web
------------------------------------------------------------------

v2016.2
-------
 * Add staff to 365 (no mailbox)
 * Add import checks to avoid reimporting the same modules
 * cleanup windows and unicode characters

v2016.1
-------
 * Add email notification on account creation
 * Add text file logging to allow emailing changes.
 * Remove student mail creation from all-mailbox-create
 * Add future students. (enrolled but not started yet)
 * In moodle-courses-create don't filter category
 * Add custom filter to reverse student data in export
 * Use a secure string file to take password out of user scripts
 * 365 mailbox creation for students
 * email log files created by scripts
 * moodle membership split into clear and add
   (allow clear once per day and add multiple
   times to allow for time changes.)

v2015.1
-------
 * Removed CSV files (you can create your own with the required columns)
 * Updated to most recent version
 * Obfuscated passwords and names (This means you will obviously edit to fill your own values.)


Why do this?
------------
This is a group of scripts I use to pull data from a SQL database using python and then fully automate user account creation using PowerShell.


About
-----
This is primarily a group of PowerShell scripts to automate user creation.

I have included the python scripts I use to pull our data from the employee/student database (TASS.web)


Requires
--------
 * pyodbc (If you use the pull scripts)


Links
-----
* https://github.com/lachlan-00/FimTassTic
* http://www.tassweb.com.au/
* https://code.google.com/p/pyodbc/
