# AD_Inventory
PowerShell script to collect basic PC data within a domain and output the results as an Excel table.

This script was tested on Windows 10 PowerShell 5.1 but should work on both newer and older versions.

The script collects PC information, specifically: CPU details (including core and thread counts); RAM specifications; details on the currently logged-in user and all users who have previously signed in on the machine; OS and build version; computer name; IP address and all network interfaces; manufacturer and model information; and details on peripherals (excluding mice and keyboards).

All results are exported to an Excel spreadsheet, with functionality included to append data to an existing file.

Connection diagnostics are performed when errors occur, and the diagnostic results and error details are recorded on a separate sheet within the Excel spreadsheet.

# Screenshots:

## Sheet Computers
<div align="center">
  <img src="images/Computers.png" alt="Computers sheet" width="100%" />
</div>

## Sheet Network
<div align="center">
  <img src="images/Network.png" alt="Network sheet" width="100%" />
</div>

## Sheet Users
<div align="center">
  <img src="images/Users.png" alt="Users sheet" width="100%" />
</div>

## Sheet Printers
<div align="center">
  <img src="images/Printers.png" alt="Printers sheet" width="100%" />
</div>

## Sheet Devices
<div align="center">
  <img src="images/Devices.png" alt="Devies sheet" width="100%" />
</div>

## Sheet Diagnostics
<div align="center">
  <img src="images/Diagnostics.png" alt="Diagnostics sheet" width="100%" />
</div>

## Sheet Errors
<div align="center">
  <img src="images/Errors.png" alt="Errors sheet" width="100%" />
</div>
