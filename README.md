# Localisation

Localisation is a simple powershell program which can be executed with a simple clic on picture file
View on only one clic the position on google map about the place of the picture
Work / Test on Windows 10 PC

# Pre requisite :
* install the register key with the file qsd.reg ( double clic )
* set powershell policy Set-ExecutionPolicy bypass
	Get-ExecutionPolicy -List
	Set-ExecutionPolicy Bypass -Scope MachinePolicy
	Set-ExecutionPolicy Bypass -Scope UserPolicy
	Set-ExecutionPolicy Bypass -Scope Process
	Set-ExecutionPolicy Bypass -Scope CurrentUser	
* copy the directory in D:\POWERSHELL

   
# Directory

Example of tree :
D:\POWERSHELL
│   LICENSE
│   localisation.ps1
│   qsd.reg
│  README.md
├──Image


# Need 
PowerShell

# Instructions / Usage
    1/ Execute qsd.reg to install context menu
    2/ copy on D:\powershell
    3/ right clic on the picture
    4/ google chrome will open on the right place
    
