By default, most servers are configured to not allow PowerShell scripts. To run this script, you will first need to allow scripts using:

Set-Executionpolicy Unrestricted

Then you can run the following script:
.\GetPISystemInfo.ps1

To specify a time period for the PI Message Logs and Event Logs, input the start and end times in PI time format as arguments.

For example,

.\GetPISystemInfo.ps1 *-8h *

will return the the PI Message Logs and Event Logs for the most recent 8 hours.