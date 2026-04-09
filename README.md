# WINSurvey
## A simple PowerShell-based Windows inventory and survey tool with UI
WINSurvey is useful for domain environment intelligence or pre-OS-upgrade reconnaissance.  The tool scans domain-joined machines and lists informative datapoints which could help identify their use case, common users, and upgradability. Non-domain/Windows machines can also be scanned for open ports or ICMP availability, however,  the tool is especially useful for Windows instances with winrm-enabled.

## Use

- Run the WINSurvey.ps1 script with Powershell 5.1 or newer (ISE or VScode recommended)
- In the UI, pecify scan targets in the host field or by selecting a .txt/.csv file; optionally, WINSurvey can scan Active Directory for all machines active in the past 10 days and generate a file for survey input. 
- Datapoints/properties for the specified hosts are shown in ui and/or file output written to the current users's desktop.  Each host is bracketed by preliminary and end summary lines. 



## Datapoints Returned

- Winrm enabled/reachable
- Windows Version
- OS Guess
- User folders list
- Scheduled Tasks (non-system)
- Open web/mail ports (443, 80, 8443, 8080, 25)
- ICMP polling (ping)
- IIS installed?
- SQL installed? + instance name(s)
- List local users
- List local groups
- Determine last logged on user

## Considerations

-- Winrm required for non-ICMP datapoint polling. 

-- RSAT required for domain device scan + DomainMachines.txt output

-- Port scan values are processesed as: 
   - If this handshake completes in time → Open 
   - If the server refuses it → Closed 
   - If nothing answers → Filtered

-- The first line of txt/csv import files will be treated as a hostname
