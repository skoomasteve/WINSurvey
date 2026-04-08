# WINSurvey
## A simple Powershell-based Windows inventory and survey tool with UI. 

Scans defined any windows devices specified in the host field, .txt file or .CSV file and returns datapoints/properties for those hosts into ui and file output.  Especially useful for Server instances but compatible with any winrm-enabled windows machine.

## Datapoints Returned

- Winrm enabled/reachable
- User folders list
- Scheduled Tasks (non-system)
- Open web/mail ports (443, 80, 8443, 8080, 25)
- ICMP polling (ping)
- IIS installed?
- SQL installed? + instance Name(s)

-Runs best in ISE or VScode
-Winrm required for datapoint polling. 
