###############################################################################
###                                                                         ###
###  Generate Secure password files for default locations                   ###
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

# Create secure file for default user password
Read-Host -Prompt "Enter your default user password" -AsSecureString | ConvertFrom-SecureString | Out-File "C:\DATA\DefaultPassword.txt"

# Create secure string to remember password
Read-Host -Prompt "Enter your 365 administrator password" -AsSecureString | ConvertFrom-SecureString | Out-File "C:\DATA\365securestring.txt"